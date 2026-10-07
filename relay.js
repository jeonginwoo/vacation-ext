// 페이지(MAIN world)에서 받은 메시지를 확장 백그라운드로 전달합니다.
window.addEventListener("message", (event) => {
  if (event.source !== window) return;
  const data = event.data;
  if (!data || data.__vacationBoard !== true) return;
  try {
    const type = ["diag", "refresh-done", "me", "balance"].includes(data.kind) ? data.kind : "capture";
    chrome.runtime.sendMessage({ type, payload: data });
  } catch {
    // 확장이 새로고침된 직후 등에는 무시
  }
});

// pass the saved own email to the page so the auto refresh can ask for the user's own balance only
chrome.storage.local.get("me").then(({ me }) => {
  if (me?.email) window.postMessage({ __vacationBoardCfg: true, email: me.email }, "*");
}).catch(() => {});

// 최상위 Power Apps 페이지에 진단 요약을 표시해 둡니다 (문제 확인용, 화면에는 안 보임).
if (window.top === window) {
  const root = document.documentElement;
  const show = async () => {
    try {
      const { diag = {}, vacations = {} } = await chrome.storage.local.get(["diag", "vacations"]);
      root.dataset.vbExt = chrome.runtime.id;
      root.dataset.vbDiag = JSON.stringify({
        vacations: Object.keys(vacations).length,
        counters: diag.counters ?? {},
        recent: (diag.log ?? []).slice(-15),
      });
    } catch {}
  };
  show();
  chrome.storage.onChanged.addListener(show);
}
