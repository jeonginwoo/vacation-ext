// 확장 저장소에 모인 휴가 데이터를 월간 캘린더로 그립니다. (빌드 도구 없는 순수 JS)
(() => {
  const POWERAPPS_URL =
    "https://apps.powerapps.com/play/e/default-fc6c5276-13d7-46c6-a396-d18ca58b5ad9/a/d46f58f0-e3d2-4138-b41d-bad3e0ba351b";
  const WEEKDAYS = ["월", "화", "수", "목", "금"];
  const DAY_NAMES = ["일", "월", "화", "수", "목", "금", "토"];
  const TYPE_COLORS = [
    ["반반차", "#2E9BB0"], ["반차", "#0F8A6E"], ["근속", "#7A4FD0"],
    ["공가", "#C27C0E"], ["경조", "#C2456B"], ["연차", "#2F64D6"],
  ];

  /* ---------- 데이터 정규화 ---------- */

  // SharePoint 커넥터는 선택 필드를 {Value: "..."} 객체로 줄 때가 있습니다.
  function val(v) {
    if (v == null) return "";
    if (typeof v === "object") return String(v.Value ?? v.DisplayName ?? v.Title ?? "").trim();
    return String(v).trim();
  }

  function toKstDate(v) {
    const s = val(v);
    if (!s) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return s.slice(0, 10);
    return new Date(d.getTime() + 9 * 3600e3).toISOString().slice(0, 10);
  }

  function normalize(store) {
    const entries = [];
    for (const [key, r] of Object.entries(store.vacations ?? {})) {
      const name = val(r.UserName);
      const date = toKstDate(r.UseDate);
      if (!name || !date) continue;
      if (val(r.CancelYN).toUpperCase() === "Y" || r.CancelYN === true) continue;
      entries.push({
        id: key,
        name,
        dept: val(r.UserDeptName),
        type: val(r.TypeName) || "기타",
        half: val(r.HalfName) || null,
        date,
        period: val(r.Period),
      });
    }
    entries.sort((a, b) => a.date.localeCompare(b.date) || a.name.localeCompare(b.name, "ko"));

    const holidays = new Map();
    for (const r of Object.values(store.holidays ?? {})) {
      const date = toKstDate(r.HolidayDate);
      if (!date) continue;
      const list = holidays.get(date) ?? [];
      const name = val(r.Title) || val(r.HolidayType);
      if (!list.some((h) => h.name === name)) list.push({ name, type: val(r.HolidayType) });
      holidays.set(date, list);
    }
    return { entries, holidays, updatedAt: store.updatedAt ?? null };
  }

  /* ---------- 날짜 유틸 ---------- */

  function ymd(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function monthGrid(year, month) {
    const first = new Date(year, month, 1);
    const last = new Date(year, month + 1, 0);
    const start = new Date(year, month, 1 - ((first.getDay() + 6) % 7));
    const end = new Date(year, month + 1, (7 - last.getDay()) % 7);
    const weeks = [];
    for (const d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      if (!weeks.length || weeks[weeks.length - 1].length === 7) weeks.push([]);
      weeks[weeks.length - 1].push(new Date(d));
    }
    // 토·일은 빼고, 평일 중 이번 달 날짜가 하나도 없는 주는 제외
    return weeks
      .map((w) => w.slice(0, 5))
      .filter((w) => w.some((d) => d.getMonth() === month));
  }

  function typeColor(type) {
    return (TYPE_COLORS.find(([k]) => type.includes(k)) ?? [null, "#6B7280"])[1];
  }

  function entryLabel(e) {
    return e.half ? `${e.half} ${e.type}` : e.type;
  }

  // 주말이면 다음 월요일로
  function toWeekday(dateStr) {
    const d = new Date(dateStr + "T00:00:00");
    const dow = d.getDay();
    if (dow === 6) d.setDate(d.getDate() + 2);
    if (dow === 0) d.setDate(d.getDate() + 1);
    return ymd(d);
  }

  /* ---------- DOM 헬퍼 (데이터는 항상 textContent로 넣습니다) ---------- */

  function h(tag, props, ...children) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(props ?? {})) {
      if (v == null || v === false) continue;
      if (k === "class") el.className = v;
      else if (k === "style") for (const [sk, sv] of Object.entries(v)) el.style.setProperty(sk, sv);
      else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? "" : v);
    }
    for (const c of children.flat()) {
      if (c == null || c === false) continue;
      el.append(c instanceof Node ? c : document.createTextNode(String(c)));
    }
    return el;
  }

  /* ---------- 상태 ---------- */

  const now = new Date();
  const today = ymd(now);
  const state = {
    year: now.getFullYear(),
    month: now.getMonth(),
    selected: toWeekday(today),
    query: "",
    depts: new Set(loadDepts()),
    hiddenTypes: new Set(),
    expanded: new Map(),   // org-tree node open state (code -> bool); default when absent
    org: null,             // { company, nodes, byCode, children, roots, byName }
    data: { entries: [], holidays: new Map(), updatedAt: null },
    refresh: null,
    balance: null,          // own leave balance { annual, longService }
    diag: null,
  };

  // 선택한 부서는 다음에 열 때도 유지
  function loadDepts() {
    try { return JSON.parse(localStorage.getItem("vb.depts") || "[]"); } catch { return []; }
  }
  function saveDepts() {
    try { localStorage.setItem("vb.depts", JSON.stringify([...state.depts])); } catch {}
  }

  function move(delta) {
    const d = new Date(state.year, state.month + delta, 1);
    state.year = d.getFullYear();
    state.month = d.getMonth();
    render();
  }

  /* ---------- 연속 일정 묶기 ---------- */

  /* ---------- org chart ---------- */

  function buildOrg(raw) {
    if (!raw?.nodes?.length) return null;
    const nodes = [...raw.nodes].sort((a, b) => a.order - b.order);
    const byCode = new Map(nodes.map((n) => [n.code, n]));
    const children = new Map();
    const roots = [];
    for (const n of nodes) {
      if (n.parent && byCode.has(n.parent)) {
        const list = children.get(n.parent) ?? [];
        list.push(n);
        children.set(n.parent, list);
      } else roots.push(n);
    }
    const byName = new Map();
    for (const n of nodes) {
      const list = byName.get(n.name) ?? [];
      list.push(n);
      byName.set(n.name, list);
    }
    return { company: raw.company, updatedAt: raw.updatedAt, diag: raw.diag, nodes, byCode, children, roots, byName };
  }

  function depthOf(org, node) {
    let d = 0, cur = node;
    while (cur.parent && org.byCode.has(cur.parent)) { d++; cur = org.byCode.get(cur.parent); }
    return d;
  }

  // node name plus every sub-department name
  function subtreeNames(org, node, out = new Set()) {
    out.add(node.name);
    for (const c of org.children.get(node.code) ?? []) subtreeNames(org, c, out);
    return out;
  }

  // selected departments (incl. parent units) -> actual department names to filter by
  function effectiveDepts(org, selected) {
    if (!org) return selected;
    const out = new Set();
    for (const name of selected) {
      const nodes = org.byName.get(name);
      if (nodes) for (const n of nodes) subtreeNames(org, n, out);
      else out.add(name);
    }
    return out;
  }

  function ancestorSelected(org, node) {
    let cur = node;
    while (cur.parent && org.byCode.has(cur.parent)) {
      cur = org.byCode.get(cur.parent);
      if (state.depts.has(cur.name)) return true;
    }
    return false;
  }

  function toggleDept(name) {
    if (state.depts.has(name)) state.depts.delete(name);
    else {
      state.depts.add(name);
      // selecting a parent unit drops sub-departments that were picked individually
      const org = state.org;
      if (org) for (const n of org.byName.get(name) ?? []) {
        for (const sub of subtreeNames(org, n)) if (sub !== name) state.depts.delete(sub);
      }
    }
    saveDepts();
    render();
  }

  // department tree with this month's number of people on leave per node (subtree, unique names)
  function renderOrgTree(org, peopleByDept) {
    const countCache = new Map();
    const countFor = (node) => {
      if (countCache.has(node.code)) return countCache.get(node.code);
      const people = new Set();
      for (const name of subtreeNames(org, node)) for (const p of peopleByDept.get(name) ?? []) people.add(p);
      countCache.set(node.code, people.size);
      return people.size;
    };

    const renderNode = (node) => {
      const kids = org.children.get(node.code) ?? [];
      const depth = depthOf(org, node);
      const open = state.expanded.has(node.code) ? state.expanded.get(node.code) : depth < 2;
      const selected = state.depts.has(node.name);
      const implied = !selected && ancestorSelected(org, node);
      const count = countFor(node);
      return h("li", { class: "org-node", role: "treeitem", "aria-expanded": kids.length ? String(open) : null },
        h("div", { class: ["org-row", selected && "selected", implied && "implied"].filter(Boolean).join(" ") },
          kids.length
            ? h("button", {
                type: "button", class: "org-caret", "aria-label": open ? "접기" : "펼치기",
                onclick: () => { state.expanded.set(node.code, !open); render(); },
              }, open ? "▾" : "▸")
            : h("span", { class: "org-caret-space" }),
          h("button", {
            type: "button", class: "org-name",
            "aria-pressed": String(selected || implied),
            disabled: implied,
            title: implied ? "상위 부서가 선택되어 함께 표시 중입니다" : null,
            onclick: () => toggleDept(node.name),
          },
            h("span", { class: "org-check" }),
            h("span", { class: "org-label" }, node.name),
            count ? h("span", { class: "org-count" }, count) : null),
        ),
        kids.length && open ? h("ul", { role: "group" }, kids.map(renderNode)) : null,
      );
    };

    return h("ul", { class: "org-tree", role: "tree", "aria-label": "조직도" }, org.roots.map(renderNode));
  }

  const MAX_LANES = Infinity; // show every lane, never collapse into "N more"

  function addDays(dateStr, n) {
    const d = new Date(dateStr + "T00:00:00");
    d.setDate(d.getDate() + n);
    return ymd(d);
  }

  // 다음 영업일 (주말·공휴일 건너뜀)
  function nextBusinessDay(dateStr, holidays) {
    let cur = addDays(dateStr, 1);
    for (let i = 0; i < 14; i++) {
      const dow = new Date(cur + "T00:00:00").getDay();
      if (dow !== 0 && dow !== 6 && !holidays.has(cur)) return cur;
      cur = addDays(cur, 1);
    }
    return cur;
  }

  // 같은 사람·같은 구분(반차는 오전/오후까지)이 영업일 기준으로 이어지면 하나의 일정으로 묶습니다.
  function buildSpans(list, holidays) {
    const groups = new Map();
    for (const e of list) {
      const key = `${e.name}|${e.dept}|${e.type}|${e.half ?? ""}`;
      const g = groups.get(key) ?? [];
      g.push(e);
      groups.set(key, g);
    }
    const spans = [];
    for (const g of groups.values()) {
      const byDay = new Map(g.map((e) => [e.date, e]));
      const dates = [...byDay.keys()].sort();
      let cur = null;
      for (const date of dates) {
        if (cur && nextBusinessDay(cur.end, holidays) === date) {
          cur.end = date;
          cur.days += 1;
        } else {
          const e = byDay.get(date);
          cur = { ...e, id: e.id, start: date, end: date, days: 1 };
          spans.push(cur);
        }
      }
    }
    // 긴 일정이 위로 오도록
    return spans.sort((a, b) => a.start.localeCompare(b.start) || b.days - a.days || a.name.localeCompare(b.name, "ko"));
  }

  function renderWeek(week, spans, byDate, holidays) {
    const keys = week.map(ymd);
    const first = keys[0], last = keys[keys.length - 1];

    // 이번 주에 걸치는 구간만 잘라서 열 위치 계산
    const segs = [];
    for (const sp of spans) {
      if (sp.end < first || sp.start > last) continue;
      let s = keys.findIndex((k) => k >= sp.start);
      let e = -1;
      for (let i = keys.length - 1; i >= 0; i--) if (keys[i] <= sp.end) { e = i; break; }
      if (s < 0 || e < s) continue;
      segs.push({ sp, s, e, contL: sp.start < keys[s], contR: sp.end > keys[e] });
    }
    segs.sort((a, b) => a.s - b.s || (b.e - b.s) - (a.e - a.s));

    // 겡치지 않게 줄(lane) 배정
    const lanes = [];
    for (const seg of segs) {
      let lane = lanes.findIndex((occ) => occ.slice(seg.s, seg.e + 1).every((x) => !x));
      if (lane < 0) { lane = lanes.length; lanes.push(Array(5).fill(false)); }
      for (let i = seg.s; i <= seg.e; i++) lanes[lane][i] = true;
      seg.lane = lane;
    }
    const shownLanes = Math.min(lanes.length, MAX_LANES);
    const hiddenPerDay = Array(5).fill(0);
    for (const seg of segs) {
      if (seg.lane < MAX_LANES) continue;
      for (let i = seg.s; i <= seg.e; i++) hiddenPerDay[i] += 1;
    }

    const select = (key) => { state.selected = key; render(); };
    const children = [];

    week.forEach((d, i) => {
      const key = keys[i];
      const list = byDate.get(key) ?? [];
      const hol = holidays.get(key);
      const cls = ["day",
        d.getMonth() !== state.month && "outside",
        key === today && "today",
        key === state.selected && "selected",
        hol && "holiday",
      ].filter(Boolean).join(" ");
      // 배경 칸 (클릭 영역)
      children.push(h("button", {
        type: "button", class: cls + (i === 0 ? " first" : ""), style: { "grid-column": String(i + 1) },
        "aria-label": `${d.getMonth() + 1}월 ${d.getDate()}일${hol ? " " + hol[0].name : ""}, 휴가 ${list.length}명`,
        onclick: () => select(key),
      }));
      // 날짜 머리
      children.push(h("div", { class: `day-head${cls.replace("day", "")}`, style: { "grid-column": String(i + 1) } },
        h("span", { class: "day-top" },
          h("span", { class: "day-num" }, d.getDate()),
          list.length ? h("span", { class: "day-count" }, list.length) : null),
        hol ? h("span", { class: "holiday-name", title: hol.map((x) => x.name).join(", ") }, hol[0].name) : null,
      ));
      if (hiddenPerDay[i]) {
        children.push(h("button", {
          type: "button", class: "more",
          style: { "grid-column": String(i + 1), "grid-row": String(shownLanes + 2) },
          onclick: () => select(key),
        }, `${hiddenPerDay[i]}명 더`));
      }
    });

    for (const seg of segs) {
      if (seg.lane >= MAX_LANES) continue;
      const sp = seg.sp;
      const label = sp.half ?? sp.type;
      const range = sp.start === sp.end ? sp.start : `${sp.start} ~ ${sp.end}`;
      children.push(h("button", {
        type: "button",
        class: ["bar", seg.contL && "cont-left", seg.contR && "cont-right"].filter(Boolean).join(" "),
        style: {
          "--type-color": typeColor(sp.type),
          "grid-column": `${seg.s + 1} / ${seg.e + 2}`,
          "grid-row": String(seg.lane + 2),
        },
        title: `${sp.name} (${sp.dept}) ${entryLabel(sp)}, ${range}`,
        // 막대를 누르면 누른 위치의 날짜를 선택
        onclick: (ev) => {
          const rect = ev.currentTarget.getBoundingClientRect();
          const cols = seg.e - seg.s + 1;
          const idx = Math.min(cols - 1, Math.max(0, Math.floor(((ev.clientX - rect.left) / rect.width) * cols)));
          select(keys[seg.s + idx]);
        },
      },
        h("span", { class: "chip-name" }, sp.name),
        h("span", { class: "chip-type" }, sp.days > 1 && !seg.contL ? `${label} ${sp.days}일` : label),
      ));
    }

    return h("div", {
      class: "week",
      style: { "grid-template-rows": `auto repeat(${shownLanes}, 22px) auto 1fr` },
    }, children);
  }

  /* ---------- 렌더링 ---------- */

  const app = document.getElementById("app");

  // The top bar (with the search box) is built once and only updated in place.
  // Re-creating the search input on every keystroke breaks Korean IME composition
  // (characters end up split into jamo).
  const searchInput = h("input", {
    id: "search", type: "search", placeholder: "이름 또는 부서 검색",
    oninput: (ev) => { state.query = ev.target.value; render(); },
  });
  const titleEl = h("h1", { class: "month-title" });
  const metaEl = h("div", { class: "meta" }); // last collected · refresh · clear
  const balanceEl = h("div", { class: "balance", "aria-live": "polite" });
  // month navigation, placed just above the calendar's top-right corner
  const navEl = h("div", { class: "nav-buttons" },
    h("button", { type: "button", "aria-label": "이전 달", onclick: () => move(-1) }, "‹"),
    h("button", {
      type: "button",
      onclick: () => {
        state.year = now.getFullYear(); state.month = now.getMonth(); state.selected = toWeekday(today); render();
      },
    }, "오늘"),
    h("button", { type: "button", "aria-label": "다음 달", onclick: () => move(1) }, "›"),
  );
  const legendSlot = h("span", { class: "legend-slot" });
  const toolbarEl = h("header", { class: "toolbar" },
    h("div", { class: "filters" }, balanceEl),
  );
  // Row above the three panels:
  //   above the org chart  -> search
  //   above the calendar   -> month title + collection status (left), month navigation (right)
  //   above the day detail -> leave-type checkboxes
  const searchWrap = h("div", { class: "search-wrap" }, searchInput);
  const headEl = h("div", { class: "calendar-head" },
    h("div", { class: "calendar-head-left" }, titleEl, metaEl),
    navEl,
  );
  const legendWrap = h("div", { class: "legend-wrap" }, legendSlot);
  // The layout and the head row stay mounted; only the sidebar / calendar / detail panels are swapped
  // on each render, so the search input is never detached while typing.
  let sideEl = h("aside", { class: "org" });
  let calEl = h("section", { class: "calendar" });
  let detailEl = h("aside", { class: "detail" });
  const layoutEl = h("div", { class: "layout" }, searchWrap, headEl, legendWrap, sideEl, calEl, detailEl);
  app.replaceChildren(toolbarEl, layoutEl);

  function render() {
    const { entries, holidays, updatedAt } = state.data;
    const weeks = monthGrid(state.year, state.month);
    const from = ymd(weeks[0][0]);
    const to = ymd(weeks[weeks.length - 1][4]);

    const depts = [...new Set(entries.map((e) => e.dept).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
    const types = [...new Set(entries.map((e) => e.type))].sort((a, b) => a.localeCompare(b, "ko"));

    const q = state.query.trim();
    // 검색/필터는 전체 기간에 적용 (달을 넘나드는 일정도 일수가 정확하도록)
    const base = entries.filter((e) =>
      (!q || e.name.includes(q) || e.dept.includes(q)) &&
      !state.hiddenTypes.has(e.type)
    );
    const deptSet = effectiveDepts(state.org, state.depts);
    const matched = state.depts.size === 0 ? base : base.filter((e) => deptSet.has(e.dept));
    const filtered = matched.filter((e) => e.date >= from && e.date <= to);
    const byDate = new Map();
    for (const e of filtered) {
      const list = byDate.get(e.date) ?? [];
      list.push(e);
      byDate.set(e.date, list);
    }
    const monthPrefix = `${state.year}-${String(state.month + 1).padStart(2, "0")}`;


    // update the persistent top bar in place (the search input itself is never replaced)
    titleEl.replaceChildren(`${state.year}년 `, h("strong", {}, `${state.month + 1}월`));
    if (searchInput.value !== state.query) searchInput.value = state.query;
    renderBalance();

    const rf = state.refresh;
    const running = rf?.status === "running" && Date.now() - rf.at < 3 * 60_000;
    // collection status lives in the top bar, next to the month title
    metaEl.replaceChildren(...[
      updatedAt
        ? h("span", {}, `마지막 수집 ${new Date(updatedAt).toLocaleString("ko-KR", {
            month: "long", day: "numeric", hour: "2-digit", minute: "2-digit",
          })}`)
        : null,
      h("button", {
        type: "button", class: "refresh-btn", disabled: running,
        onclick: () => requestRefresh(),
      }, running ? "갱신 중…" : "휴가 정보 갱신"),
      rf?.status === "failed" && !running
        ? h("span", { class: "refresh-failed" },
            "자동으로 받지 못했어요. ",
            h("a", { href: POWERAPPS_URL, target: "_blank", rel: "noopener" }, "Power Apps"),
            "에 로그인되어 있는지 확인하고 직접 열어주세요.")
        : null,
      updatedAt
        ? h("button", {
            type: "button",
            onclick: async () => {
              if (!confirm("수집한 휴가 데이터를 모두 지울까요?")) return;
              // department selection belongs to the collected org chart, so reset it too
              state.depts.clear();
              saveDepts();
              await chrome.storage.local.clear();
              chrome.action.setBadgeText({ text: "" });
            },
          }, "수집 데이터 지우기")
        : null,
    ].filter(Boolean));

    const legend = types.length
      ? h("div", { class: "legend", role: "group", "aria-label": "휴가 구분 필터" },
          types.map((t) => h("button", {
            type: "button", class: "legend-item",
            "aria-pressed": String(!state.hiddenTypes.has(t)),
            style: { "--type-color": typeColor(t) },
            onclick: () => {
              if (state.hiddenTypes.has(t)) state.hiddenTypes.delete(t); else state.hiddenTypes.add(t);
              render();
            },
          }, h("span", { class: "swatch" }), t)))
      : null;

    const spans = buildSpans(matched, holidays);
    const calendar = h("section", { class: "calendar", "aria-label": "월간 휴가 캘린더" },
      h("div", { class: "weekday-row" },
        WEEKDAYS.map((w) => h("div", { class: "weekday" }, w))),
      weeks.map((week) => renderWeek(week, spans, byDate, holidays)),
    );

    const sel = state.selected;
    const selEntries = byDate.get(sel) ?? [];
    const selHol = holidays.get(sel);
    const detail = h("aside", { class: "detail", "aria-live": "polite" },
      h("h2", {}, `${Number(sel.slice(5, 7))}월 ${Number(sel.slice(8, 10))}일`,
        h("span", { class: "detail-weekday" }, `${DAY_NAMES[new Date(sel + "T00:00:00").getDay()]}요일`)),
      selHol ? h("p", { class: "detail-holiday" }, selHol.map((x) => x.name).join(", ")) : null,
      selEntries.length === 0
        ? h("p", { class: "empty" }, "이 날은 휴가자가 없습니다.")
        : h("ul", { class: "detail-list" },
            selEntries.map((e) => h("li", { style: { "--type-color": typeColor(e.type) } },
              h("div", { class: "detail-main" },
                h("span", { class: "detail-name" }, e.name),
                h("span", { class: "detail-type" }, entryLabel(e))),
              h("div", { class: "detail-sub" },
                h("span", {}, e.dept),
                e.period ? h("span", { class: "detail-period" }, e.period) : null)))),
    );

    // org-chart sidebar: people on leave this month per department
    const peopleByDept = new Map();
    for (const e of base) {
      if (!e.date.startsWith(monthPrefix)) continue;
      const set = peopleByDept.get(e.dept) ?? new Set();
      set.add(e.name);
      peopleByDept.set(e.dept, set);
    }
    // flat checkbox row (departments without an org-tree node, or every department before the tree is loaded)
    const deptRow = (d) => h("li", { class: "org-node" },
      h("div", { class: ["org-row", state.depts.has(d) && "selected"].filter(Boolean).join(" ") },
        h("span", { class: "org-caret-space" }),
        h("button", {
          type: "button", class: "org-name", "aria-pressed": String(state.depts.has(d)),
          onclick: () => toggleDept(d),
        },
          h("span", { class: "org-check" }),
          h("span", { class: "org-label" }, d),
          peopleByDept.get(d)?.size ? h("span", { class: "org-count" }, peopleByDept.get(d).size) : null)));
    const clearBtn = state.depts.size
      ? h("button", {
          type: "button", class: "tag-clear",
          onclick: () => { state.depts.clear(); saveDepts(); render(); },
        }, "선택 해제")
      : null;

    let sidebar;
    if (state.org) {
      const known = new Set(state.org.nodes.map((n) => n.name));
      const others = depts.filter((d) => !known.has(d));
      sidebar = h("aside", { class: "org" },
        h("div", { class: "org-head" }, h("h2", {}, state.org.company || "조직도"), clearBtn),
        renderOrgTree(state.org, peopleByDept),
        // structure could not be read as a hierarchy -> show what the page looked like (no names)
        state.org.nodes.length > 3 && state.org.nodes.every((n) => !n.parent)
          ? h("details", { class: "diag org-diag" },
              h("summary", {}, "조직도 구조를 읽지 못했어요 (진단 정보)"),
              h("pre", {}, JSON.stringify(state.org.diag ?? {}, null, 2)))
          : null,
        others.length
          ? h("div", { class: "org-others" },
              h("p", { class: "org-others-title" }, "조직도에 없는 부서"),
              h("ul", { class: "org-tree" }, others.map(deptRow)))
          : null,
      );
    } else {
      // no org chart yet: same checkbox UI, as a flat list of departments found in the leave data
      const list = [...new Set([...depts, ...state.depts])].sort((a, b) => a.localeCompare(b, "ko"));
      sidebar = h("aside", { class: "org" },
        h("div", { class: "org-head" }, h("h2", {}, "부서"), clearBtn),
        list.length
          ? h("ul", { class: "org-tree" }, list.map(deptRow))
          : h("p", { class: "org-hint-empty" }, "휴가 데이터가 들어오면 부서 목록이 표시됩니다."),
        h("p", { class: "org-hint" }, "크롬에서 ",
          h("a", { href: "https://teams.cloud.microsoft/", target: "_blank", rel: "noopener" }, "Teams 웹"),
          "을 열고 조직도 탭을 한 번 누르면 상위 부서까지 트리로 볼 수 있어요. Teams 데스크톱 앱에서는 가져올 수 없어요."),
      );
    }

    legendSlot.replaceChildren(...(legend ? [legend] : []));
    sideEl.replaceWith(sidebar); sideEl = sidebar;
    calEl.replaceWith(calendar); calEl = calendar;
    detailEl.replaceWith(detail); detailEl = detail;
  }

  /* ---------- 시작 ---------- */

  /* ---------- own leave balance ---------- */

  // 0.5 -> "0.5", 6 -> "6"
  const fmtDays = (n) => (n == null ? "-" : String(Math.round(n * 100) / 100));

  function renderBalance() {
    const b = state.balance ?? {};
    const parts = [];
    const item = (label, v, cls) => h("span", {
      class: `balance-item ${cls}`,
      title: `${v.year ? v.year + "년 · " : ""}부여 ${fmtDays(v.total)}일 · 사용 ${fmtDays(v.used)}일 · 잔여 ${fmtDays(v.remain)}일`,
    },
      h("span", { class: "balance-label" }, label),
      h("strong", {}, fmtDays(v.remain)),
      h("span", { class: "balance-total" }, `/ ${fmtDays(v.total)}일`));
    if (b.annual) parts.push(item("내 잔여 연차", b.annual, "annual"));
    // long-service leave only when the user actually has some
    if (b.longService && (b.longService.total ?? 0) > 0) parts.push(item("근속연차", b.longService, "long"));
    balanceEl.replaceChildren(...parts);
    balanceEl.hidden = parts.length === 0;
  }

  /* ---------- auto refresh ---------- */

  const STALE_MS = 60 * 60_000; // refresh automatically when data is older than 1 hour

  function requestRefresh() {
    state.refresh = { status: "running", at: Date.now() };
    render();
    chrome.runtime.sendMessage({ type: "refresh" });
  }

  let autoChecked = false;
  function maybeAutoRefresh() {
    if (autoChecked) return;
    autoChecked = true;
    const rf = state.refresh;
    const running = rf?.status === "running" && Date.now() - rf.at < 3 * 60_000;
    const recentlyFailed = rf?.status === "failed" && Date.now() - rf.at < 10 * 60_000;
    const empty = state.data.entries.length === 0;
    const stale = !state.data.updatedAt || Date.now() - state.data.updatedAt > STALE_MS;
    if (running) return;
    // no leave data yet -> always fetch; old data -> fetch unless it just failed
    if (empty || (stale && !recentlyFailed)) requestRefresh();
  }

  async function load() {
    const store = await chrome.storage.local.get(["vacations", "holidays", "updatedAt", "diag", "orgTree", "refresh", "balance"]);
    state.refresh = store.refresh ?? null;
    state.balance = store.balance ?? null;
    state.org = buildOrg(store.orgTree);
    state.data = normalize(store);
    state.diag = store.diag ?? null;
    render();
  }

  window.addEventListener("keydown", (e) => {
    if (e.target instanceof HTMLInputElement || e.target instanceof HTMLSelectElement) return;
    if (e.key === "ArrowLeft") move(-1);
    if (e.key === "ArrowRight") move(1);
  });

  // Power Apps에서 새 데이터가 들어오면 자동으로 다시 그립니다.
  chrome.storage.onChanged.addListener((_changes, area) => { if (area === "local") load(); });

  load().then(maybeAutoRefresh);
})();
