// Merges captured data into chrome.storage.local.
// Layout: { vacations: { "<listId>:<ID>": {...fields} }, holidays: {...}, orgTree, refresh, updatedAt }

let queue = Promise.resolve();

// no number badge on the toolbar icon (also clears one left over from older versions)
chrome.action.setBadgeText({ text: "" });

const POWERAPPS_URL =
  "https://apps.powerapps.com/play/e/default-fc6c5276-13d7-46c6-a396-d18ca58b5ad9/a/d46f58f0-e3d2-4138-b41d-bad3e0ba351b";
const REFRESH_TIMEOUT_MIN = 2;

/* ---------- auto refresh ----------
   1) If a Power Apps tab is already open, ask that tab to re-run the leave queries (nothing opens).
   2) Otherwise (or if that fails) open Power Apps in an inactive background tab of the current
      window and close it when done. */

const REUSE_TIMEOUT_MIN = 0.5; // 30 s for an already-open tab

async function setRefresh(status, extra = {}) {
  await chrome.storage.local.set({ refresh: { status, at: Date.now(), ...extra } });
}

async function startRefresh() {
  const { refreshRun } = await chrome.storage.session.get("refreshRun");
  if (refreshRun && Date.now() - refreshRun.started < (REFRESH_TIMEOUT_MIN + 1) * 60_000) {
    try { await chrome.tabs.get(refreshRun.tabId); return; } catch { /* tab gone: start over */ }
  }
  // 1) reuse an open Power Apps tab. Tab URLs are not visible without the "tabs" permission, so ask
  //    every tab: only Power Apps tabs run relay.js and answer the ping.
  let open = [];
  try {
    const tabs = await chrome.tabs.query({});
    const answers = await Promise.all(tabs.map((t) =>
      chrome.tabs.sendMessage(t.id, { type: "vb-ping" }, { frameId: 0 }).then((r) => (r?.powerapps ? t : null), () => null)));
    open = answers.filter(Boolean);
  } catch {}
  for (const tab of open) {
    try {
      await chrome.tabs.sendMessage(tab.id, { type: "vb-replay" }, { frameId: 0 });
      await chrome.storage.session.set({ refreshRun: { mode: "reuse", tabId: tab.id, started: Date.now() } });
      await setRefresh("running", { method: "reuse" });
      chrome.alarms.create("refresh-timeout", { delayInMinutes: REUSE_TIMEOUT_MIN });
      return;
    } catch { /* no collector in that tab (e.g. opened before the extension loaded): try the next */ }
  }
  // 2) background tab
  await startTabRefresh();
}

async function startTabRefresh() {
  let windowId;
  try { windowId = (await chrome.windows.getLastFocused({ windowTypes: ["normal"] })).id; } catch {}
  const tab = await chrome.tabs.create({ url: POWERAPPS_URL, active: false, ...(windowId != null ? { windowId } : {}) });
  await chrome.storage.session.set({ refreshRun: { mode: "tab", tabId: tab.id, started: Date.now() } });
  await setRefresh("running", { method: "tab" });
  chrome.alarms.create("refresh-timeout", { delayInMinutes: REFRESH_TIMEOUT_MIN });
}

async function finishRefresh(ok, reason) {
  const { refreshRun } = await chrome.storage.session.get("refreshRun");
  await chrome.storage.session.remove("refreshRun");
  chrome.alarms.clear("refresh-timeout");
  // close only the tab we opened ourselves, never the user's own Power Apps tab
  if (refreshRun?.mode === "tab") { try { await chrome.tabs.remove(refreshRun.tabId); } catch {} }
  await setRefresh(ok ? "done" : "failed", { reason, method: refreshRun?.mode });
}

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== "refresh-timeout") return;
  const { refreshRun } = await chrome.storage.session.get("refreshRun");
  if (refreshRun?.mode === "reuse") await startTabRefresh(); // open tab did not answer in time
  else finishRefresh(false, "timeout");
});

chrome.tabs.onRemoved.addListener(async (tabId) => {
  const { refreshRun } = await chrome.storage.session.get("refreshRun");
  if (refreshRun?.tabId !== tabId) return;
  if (refreshRun.mode === "reuse") await startTabRefresh(); // user closed it meanwhile
  else finishRefresh(false, "closed");
});

/* ---------- org chart refresh ----------
   Open the Teams org chart app (deep link straight to the app's tab) in a background tab;
   orgchart.js reads the department tree inside it, then the tab is closed. */

const ORG_URL = "https://teams.cloud.microsoft/_#/l/entity/e4b811fd-b165-4805-a3fa-e5403d507da9/orgchart";
const ORG_TIMEOUT_MIN = 1.5;

async function setOrgRefresh(status, extra = {}) {
  await chrome.storage.local.set({ orgRefresh: { status, at: Date.now(), ...extra } });
}

async function startOrgRefresh() {
  const { orgRun } = await chrome.storage.session.get("orgRun");
  if (orgRun && Date.now() - orgRun.started < (ORG_TIMEOUT_MIN + 0.5) * 60_000) {
    try { await chrome.tabs.get(orgRun.tabId); return; } catch { /* gone: start over */ }
  }
  let windowId;
  try { windowId = (await chrome.windows.getLastFocused({ windowTypes: ["normal"] })).id; } catch {}
  const tab = await chrome.tabs.create({ url: ORG_URL, active: false, ...(windowId != null ? { windowId } : {}) });
  await chrome.storage.session.set({ orgRun: { tabId: tab.id, started: Date.now() } });
  await setOrgRefresh("running");
  chrome.alarms.create("org-timeout", { delayInMinutes: ORG_TIMEOUT_MIN });
}

async function finishOrgRefresh(ok, reason) {
  const { orgRun } = await chrome.storage.session.get("orgRun");
  if (!orgRun) return;
  await chrome.storage.session.remove("orgRun");
  chrome.alarms.clear("org-timeout");
  try { await chrome.tabs.remove(orgRun.tabId); } catch {}
  await setOrgRefresh(ok ? "done" : "failed", { reason });
}

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "org-timeout") finishOrgRefresh(false, "timeout");
});

chrome.tabs.onRemoved.addListener(async (tabId) => {
  const { orgRun } = await chrome.storage.session.get("orgRun");
  if (orgRun?.tabId === tabId) finishOrgRefresh(false, "closed");
});

/* ---------- merge captured list items ---------- */

function merge(payload) {
  return chrome.storage.local.get(["vacations", "holidays"]).then((store) => {
    const bucketName = payload.kind === "vacation" ? "vacations" : "holidays";
    const bucket = store[bucketName] ?? {};
    const seen = new Set();

    for (const item of payload.items) {
      if (item.ID == null) continue;
      const key = `${payload.listId}:${item.ID}`;
      seen.add(key);
      // different queries return different fields for the same item, so merge field by field
      bucket[key] = { ...(bucket[key] ?? {}), ...item, listId: payload.listId };
    }

    // an unfiltered, untruncated query is a full snapshot: drop items that are no longer there
    const isFullSnapshot =
      !payload.filtered &&
      !payload.hasNextLink &&
      (payload.top == null || payload.items.length < payload.top) &&
      (payload.kind === "holiday" || payload.items.some((i) => "UserName" in i));
    if (isFullSnapshot) {
      for (const key of Object.keys(bucket)) {
        if (key.startsWith(payload.listId + ":") && !seen.has(key)) delete bucket[key];
      }
    }

    return chrome.storage.local.set({ [bucketName]: bucket, updatedAt: Date.now() });
  });
}

// diagnostics: which frames / requests / workers were seen (shapes only, no values)
function recordDiag(p) {
  return chrome.storage.local.get("diag").then(({ diag = { counters: {}, log: [] } }) => {
    const key = `${p.origin}|${p.type}${p.data?.via ? ":" + p.data.via : ""}`;
    diag.counters[key] = (diag.counters[key] ?? 0) + 1;
    diag.log.push({ t: Date.now(), origin: p.origin, type: p.type, data: p.data });
    if (diag.log.length > 80) diag.log = diag.log.slice(-80);
    return chrome.storage.local.set({ diag });
  });
}

chrome.runtime.onMessage.addListener((msg, sender) => {
  if (msg?.type === "refresh") { startRefresh(); return; }
  if (msg?.type === "refresh-org") { startOrgRefresh(); return; }
  if (msg?.type === "refresh-done") {
    chrome.storage.session.get("refreshRun").then(({ refreshRun }) => {
      if (!refreshRun || sender.tab?.id !== refreshRun.tabId || sender.frameId !== 0) return;
      const ok = (msg.payload?.ok ?? 0) > 0;
      // an already-open tab could not refresh (e.g. its sign-in expired): use a background tab
      if (refreshRun.mode === "reuse" && !ok) { chrome.alarms.clear("refresh-timeout"); startTabRefresh(); return; }
      // let pending captures finish first
      queue = queue.then(() => finishRefresh(ok, "done"));
    });
    return;
  }
  if (!msg?.payload) return;
  if (msg.type === "capture") {
    queue = queue.then(() => merge(msg.payload)).catch((err) => console.error(err));
  } else if (msg.type === "me") {
    queue = queue.then(() => chrome.storage.local.set({ me: { email: msg.payload.email } }));
  } else if (msg.type === "balance") {
    // own leave balance only (annual / longService)
    const p = msg.payload;
    queue = queue.then(async () => {
      const { balance = {} } = await chrome.storage.local.get("balance");
      balance[p.balanceKind] = { year: p.year, total: p.total, used: p.used, remain: p.remain, at: Date.now() };
      await chrome.storage.local.set({ balance });
    }).catch((err) => console.error(err));
  } else if (msg.type === "balances") {
    // everyone's leave balance (annual / longService), replaced as a snapshot per list
    const p = msg.payload;
    queue = queue.then(async () => {
      const { balances = {} } = await chrome.storage.local.get("balances");
      balances[p.balanceKind] = { rows: p.rows, at: Date.now() };
      await chrome.storage.local.set({ balances });
    }).catch((err) => console.error(err));
  } else if (msg.type === "orgtree") {
    // org chart: keep department structure only (replace with latest)
    queue = queue.then(() => chrome.storage.local.set({
      orgTree: { ...msg.payload, updatedAt: Date.now() },
    })).catch((err) => console.error(err));
    // came from the background org chart tab we opened: done, close it
    chrome.storage.session.get("orgRun").then(({ orgRun }) => {
      if (orgRun && sender.tab?.id === orgRun.tabId) queue = queue.then(() => finishOrgRefresh(true, "done"));
    });
  } else if (msg.type === "diag") {
    queue = queue.then(() => recordDiag(msg.payload)).catch((err) => console.error(err));
  }
});

chrome.action.onClicked.addListener(() => {
  chrome.tabs.create({ url: chrome.runtime.getURL("calendar.html") });
});
