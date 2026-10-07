// Power Apps 페이지(MAIN world)에서 SharePoint 커넥터 응답을 가로챕니다.
// 1) 페이지가 직접 하는 fetch/XHR 응답
// 2) Web Worker가 페이지로 보내는 메시지 (워커 코드는 건드리지 않고 메시지만 읽음)
// 휴가/공휴일 리스트에 필요한 필드만 추려서 relay.js로 넘기고, 진단 정보도 함께 보냅니다.
(() => {
  if (window.__vacationBoardHooked) return;
  window.__vacationBoardHooked = true;

  const ITEMS_RE = /\/apis\/sharepointonline\/.*?\/tables\/([0-9a-f-]{36})\/items/i;
  const CONNECTOR_RE = /environment\.api\.powerplatform\.com|\/apis\/sharepointonline\//i;
  const VACATION_FIELDS = ["ID", "UserName", "UserDeptName", "TypeName", "HalfName",
    "UseDate", "Period", "UseYear", "CancelYN"];
  const HOLIDAY_FIELDS = ["ID", "Title", "HolidayDate", "HolidayType"];

  // Auto refresh: the same list queries the app sends when the leave screen is opened.
  // They are sent from this page with the app's own request headers (kept in memory only).
  const REFRESH_QUERIES = [
    ["7a9a67ad-a6e6-4242-9e82-9b84c9cd74dc", "HalfName,ID,Period,Title,TypeName,UseDate,UserDeptName,UserName,UseYear"],
    ["7a9a67ad-a6e6-4242-9e82-9b84c9cd74dc", "CancelYN,ID,UseYear"],
    ["ce4b7dda-cece-42d2-a009-d257973fef03", "HalfName,ID,Period,Title,TypeName,UseDate,UserDeptName,UserName,UseYear"],
    ["ce4b7dda-cece-42d2-a009-d257973fef03", "CancelYN,ID,UseYear"],
    ["ed13b1cc-b5ff-4f5b-a210-bd286adc5f57", "HolidayDate,HolidayType,ID,Title"],
  ];

  function post(msg) {
    try { window.postMessage({ __vacationBoard: true, ...msg }, "*"); } catch {}
  }
  function diag(type, data) {
    post({ kind: "diag", origin: location.origin, type, data });
  }

  function pick(obj, fields) {
    const out = {};
    for (const f of fields) if (f in obj) out[f] = obj[f];
    return out;
  }

  function kindOf(sample) {
    if (!sample || typeof sample !== "object") return null;
    if ("UseDate" in sample || "UseYear" in sample) return "vacation";
    if ("HolidayDate" in sample) return "holiday";
    return null;
  }

  function urlInfo(url) {
    const m = url ? ITEMS_RE.exec(url) : null;
    let filtered = true, top = null;
    if (m) {
      try {
        const u = new URL(url, location.href);
        filtered = u.searchParams.has("$filter");
        top = Number(u.searchParams.get("$top")) || null;
      } catch {}
    }
    return { listId: m ? m[1].toLowerCase() : null, filtered, top };
  }

  function emit(items, url, hasNextLink, via) {
    const kind = kindOf(items[0]);
    if (!kind) return false;
    const info = urlInfo(url);
    post({
      kind,
      listId: info.listId ?? `unknown-${kind}`,
      // 리스트를 알 수 없으면 정리(삭제) 판단을 하지 않도록 filtered로 취급
      filtered: info.listId ? info.filtered : true,
      top: info.top,
      hasNextLink: Boolean(hasNextLink),
      items: items.map((v) => pick(v, kind === "vacation" ? VACATION_FIELDS : HOLIDAY_FIELDS)),
    });
    diag("captured", { via, kind, listId: info.listId, count: items.length });
    return true;
  }

  const MARKERS = ['"UseDate"', '"UseYear"', '"HolidayDate"'];
  const hasMarker = (text) => MARKERS.some((m) => text.includes(m));

  function parse(text) {
    if (typeof text !== "string") return text;
    try { return JSON.parse(text); } catch {}
    // BOM, XSSI 접두어 등이 붙은 경우 첫 '{' 부터 다시 시도
    const i = text.indexOf("{");
    if (i > 0) { try { return JSON.parse(text.slice(i)); } catch {} }
    return null;
  }

  // ArrayBuffer/TypedArray → 문자열 (gzip이면 풀어서)
  async function bytesToText(buf) {
    const bytes = buf instanceof ArrayBuffer ? new Uint8Array(buf)
      : new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    if (bytes[0] === 0x1f && bytes[1] === 0x8b && typeof DecompressionStream !== "undefined") {
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
      return await new Response(stream).text();
    }
    return new TextDecoder().decode(bytes);
  }

  // XHR 응답을 responseType과 상관없이 JSON으로 읽습니다.
  async function readXhrBody(xhr) {
    const rt = xhr.responseType;
    const r = xhr.response;
    const meta = { rt, ctor: r?.constructor?.name ?? null, ct: xhr.getResponseHeader("content-type") };
    let text = null;
    if (rt === "json") return { body: r, meta };
    if (rt === "" || rt === "text") text = xhr.responseText;
    else if (rt === "arraybuffer" && r) text = await bytesToText(r);
    else if (rt === "blob" && r) text = await bytesToText(await r.arrayBuffer());
    if (text != null) {
      meta.len = text.length;
      meta.code0 = text.charCodeAt(0);
    }
    return { body: parse(text), meta };
  }

  function handleResponse(url, body, via, meta) {
    const info = urlInfo(url);
    if (!body || !Array.isArray(body.value)) {
      diag("parsed", {
        via, listId: info.listId, ok: false,
        shape: body && typeof body === "object" ? Object.keys(body).slice(0, 10) : typeof body,
        ...(meta ?? {}),
      });
      return;
    }
    if (body.value.length === 0) return;
    if (!emit(body.value, url, body["@odata.nextLink"], via)) {
      diag("parsed", { via, listId: info.listId, ok: true, kind: null, keys: Object.keys(body.value[0]).slice(0, 15) });
    }
  }

  /* ---------- 1) fetch / XHR ---------- */

  const origFetch = window.fetch;
  if (origFetch) {
    window.fetch = async function (input, init) {
      const res = await origFetch.apply(this, arguments);
      try {
        const url = typeof input === "string" ? input : input?.url ?? String(input);
        if (CONNECTOR_RE.test(url)) {
          diag("request", { via: "fetch", items: ITEMS_RE.test(url), status: res.status });
          if (ITEMS_RE.test(url) && res.ok) {
            res.clone().text().then((t) => handleResponse(url, parse(t), "fetch")).catch(() => {});
          }
        }
      } catch {}
      return res;
    };
  }

  const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    if (this.__vbUrl && CONNECTOR_RE.test(this.__vbUrl)) (this.__vbHeaders ||= {})[name] = value;
    return origSetHeader.apply(this, arguments);
  };

  let replayed = false;
  function replay(templateUrl, headers) {
    if (replayed) return;
    replayed = true;
    const base = templateUrl.slice(0, templateUrl.indexOf("/tables/") + "/tables/".length);
    let pending = REFRESH_QUERIES.length;
    let ok = 0;
    const done = () => {
      // small delay so the captured responses reach the extension before the window is closed
      if (--pending === 0) setTimeout(() => post({ kind: "refresh-done", ok, total: REFRESH_QUERIES.length }), 1500);
    };
    diag("replay", { count: REFRESH_QUERIES.length });
    for (const [listId, select] of REFRESH_QUERIES) {
      const xhr = new XMLHttpRequest();
      // goes through the patched open/send above, so the response is captured like any other
      xhr.open("GET", `${base}${listId}/items?%24select=${encodeURIComponent(select)}&%24top=500`);
      for (const [k, v] of Object.entries(headers)) {
        try { xhr.setRequestHeader(k, v); } catch {}
      }
      xhr.addEventListener("loadend", () => { if (xhr.status >= 200 && xhr.status < 300) ok++; done(); });
      xhr.send();
    }
  }

  const origOpen = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (method, url) {
    this.__vbUrl = String(url);
    return origOpen.apply(this, arguments);
  };
  const origSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function () {
    const url = this.__vbUrl;
    if (url && CONNECTOR_RE.test(url)) {
      this.addEventListener("load", () => {
        try {
          diag("request", { via: "xhr", items: ITEMS_RE.test(url), status: this.status });
          if (!ITEMS_RE.test(url) || this.status < 200 || this.status >= 300) return;
          // first successful list request from the app: use it as the template for the refresh
          if (!replayed && /\/apis\/sharepointonline\//i.test(url) && this.__vbHeaders) {
            const headers = this.__vbHeaders;
            setTimeout(() => replay(url, headers), 0);
          }
          readXhrBody(this)
            .then(({ body, meta }) => handleResponse(url, body, "xhr", meta))
            .catch((err) => diag("parsed", { via: "xhr", ok: false, error: String(err).slice(0, 200) }));
        } catch (err) {
          diag("parsed", { via: "xhr", ok: false, error: String(err).slice(0, 200) });
        }
      });
    }
    return origSend.apply(this, arguments);
  };

  /* ---------- 2) Web Worker 메시지 읽기 ---------- */

  // 메시지 안을 훑어서 휴가/공휴일 항목 배열과, 같이 들어있는 리스트 URL을 찾습니다.
  function sniff(data, via) {
    const arrays = [];
    const urls = [];
    let budget = 20000;

    function walk(node, depth) {
      if (budget-- <= 0 || node == null || depth > 10) return;
      if (typeof node === "string") {
        if (node.length < 4096 && ITEMS_RE.test(node)) urls.push(node);
        else if (node.length < 20_000_000 && hasMarker(node)) {
          const parsed = parse(node);
          if (parsed) walk(parsed, depth + 1);
        }
        return;
      }
      if (typeof node !== "object") return;
      if (node instanceof ArrayBuffer || ArrayBuffer.isView(node)) {
        try {
          const text = new TextDecoder().decode(node);
          if (hasMarker(text)) walk(parse(text), depth + 1);
        } catch {}
        return;
      }
      if (Array.isArray(node)) {
        if (node.length && kindOf(node[0])) { arrays.push(node); return; }
        for (const v of node) walk(v, depth + 1);
        return;
      }
      for (const k in node) {
        let v;
        try { v = node[k]; } catch { continue; }
        walk(v, depth + 1);
      }
    }

    try { walk(data, 0); } catch {}
    if (!arrays.length) return;
    const url = urls.length === 1 ? urls[0] : null;
    for (const arr of arrays) emit(arr, url, false, via);
  }

  function wrapWorker(name) {
    const Orig = window[name];
    if (!Orig) return;
    window[name] = new Proxy(Orig, {
      construct(target, args, newTarget) {
        const w = Reflect.construct(target, args, newTarget);
        try {
          const raw = String(args[0] ?? "");
          diag("worker", { name, url: raw.startsWith("blob:") ? "blob" : raw.split("?")[0].slice(0, 200) });
          const port = name === "SharedWorker" ? w.port : w;
          port.addEventListener("message", (ev) => sniff(ev.data, name));
        } catch {}
        return w;
      },
    });
  }
  wrapWorker("Worker");
  wrapWorker("SharedWorker");

  // 다른 프레임에서 오는 메시지에도 데이터가 실려 있을 수 있습니다.
  window.addEventListener("message", (ev) => {
    if (ev.data && ev.data.__vacationBoard) return;
    if (ev.source === window) return;
    sniff(ev.data, "frame-message");
  });

  diag("hello", { path: location.pathname.slice(0, 120) });
})();
