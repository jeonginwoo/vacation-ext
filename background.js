// Merges captured data into chrome.storage.local.
// Layout: { vacations: { "<listId>:<ID>": {...fields} }, holidays: {...}, orgTree, refresh, updatedAt }

let queue = Promise.resolve();

const POWERAPPS_URL =
  "https://apps.powerapps.com/play/e/default-fc6c5276-13d7-46c6-a396-d18ca58b5ad9/a/d46f58f0-e3d2-4138-b41d-bad3e0ba351b";
const REFRESH_TIMEOUT_MIN = 2;

/* ---------- auto refresh: open Power Apps in a minimized window, then close it ---------- */

async function setRefresh(status, extra = {}) {
  await chrome.storage.local.set({ refresh: { status, at: Date.now(), ...extra } });
}

async function startRefresh() {
  const { refreshWin } = await chrome.storage.session.get("refreshWin");
  if (refreshWin) {
    // still running unless the window is gone
    try { await chrome.windows.get(refreshWin.id); return; } catch {}
  }
  const win = await chrome.windows.create({ url: POWERAPPS_URL, state: "minimized", focused: false });
  await chrome.storage.session.set({ refreshWin: { id: win.id, started: Date.now() } });
  await setRefresh("running");
  chrome.alarms.create("refresh-timeout", { delayInMinutes: REFRESH_TIMEOUT_MIN });
}

async function finishRefresh(ok, reason) {
  const { refreshWin } = await chrome.storage.session.get("refreshWin");
  await chrome.storage.session.remove("refreshWin");
  chrome.alarms.clear("refresh-timeout");
  if (refreshWin) { try { await chrome.windows.remove(refreshWin.id); } catch {} }
  await setRefresh(ok ? "done" : "failed", { reason });
}

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "refresh-timeout") finishRefresh(false, "timeout");
});

chrome.windows.onRemoved.addListener(async (id) => {
  const { refreshWin } = await chrome.storage.session.get("refreshWin");
  if (refreshWin?.id === id) finishRefresh(false, "closed");
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

    return chrome.storage.local.set({ [bucketName]: bucket, updatedAt: Date.now() }).then(() => {
      if (bucketName !== "vacations") return;
      // badge = number of collected leave entries (shows that collection works)
      const count = Object.values(bucket).filter((v) => v.UserName && v.UseDate).length;
      chrome.action.setBadgeBackgroundColor({ color: "#2F64D6" });
      chrome.action.setBadgeText({ text: count > 999 ? "999+" : String(count) });
    });
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
  if (msg?.type === "refresh-done") {
    chrome.storage.session.get("refreshWin").then(({ refreshWin }) => {
      if (refreshWin && sender.tab?.windowId === refreshWin.id) {
        // let pending captures finish first
        queue = queue.then(() => finishRefresh((msg.payload?.ok ?? 0) > 0, "done"));
      }
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
  } else if (msg.type === "orgtree") {
    // org chart: keep department structure only (replace with latest)
    queue = queue.then(() => chrome.storage.local.set({
      orgTree: { ...msg.payload, updatedAt: Date.now() },
    })).catch((err) => console.error(err));
  } else if (msg.type === "diag") {
    queue = queue.then(() => recordDiag(msg.payload)).catch((err) => console.error(err));
  }
});

chrome.action.onClicked.addListener(() => {
  chrome.tabs.create({ url: chrome.runtime.getURL("calendar.html") });
});
