// Reads only the department tree from the Teams org-chart app (iframe) and saves it to the extension.
// Stores department code / name / parent only. Employee names and contact info are never read.
(() => {
  let lastSent = "";
  let timer = null;

  const isList = (el) => el.tagName === "UL" || el.tagName === "OL";
  const HAS_WORD = /[\p{L}\p{N}]/u;

  // the li's own label (ignores anything inside nested lists)
  function ownLabel(li) {
    const attrs = {};
    let text = "";
    for (const child of li.children) {
      if (isList(child)) continue;
      const tagged = child.matches?.("[data-treedata-groupname],[data-treedata-groupcode]")
        ? child : child.querySelector?.("[data-treedata-groupname],[data-treedata-groupcode]");
      if (tagged && !Object.keys(attrs).length) {
        for (const at of tagged.attributes) attrs[at.name.toLowerCase()] = at.value;
      }
      // skip pieces without letters/digits (expand buttons like "+", "-", icons)
      const piece = (child.textContent || "").trim();
      if (HAS_WORD.test(piece)) text += " " + piece;
    }
    // text placed directly inside the li
    for (const node of li.childNodes) {
      if (node.nodeType === Node.TEXT_NODE && HAS_WORD.test(node.textContent)) text += " " + node.textContent.trim();
    }
    const name = (attrs["data-treedata-groupname"] || text).replace(/\s+/g, " ").trim();
    return { name, code: attrs["data-treedata-groupcode"] || null };
  }

  // hierarchy from nested lists (any li, whatever its class)
  function extractFromLists(root) {
    const lis = [...root.querySelectorAll("li")];
    const info = new Map();
    lis.forEach((li, i) => {
      const { name, code } = ownLabel(li);
      if (name) info.set(li, { code: code || `node-${i}`, name, order: i });
    });
    const nodes = [];
    for (const [li, n] of info) {
      let p = li.parentElement?.closest("li");
      while (p && root.contains(p) && !info.has(p)) p = p.parentElement?.closest("li");
      n.parent = p && info.has(p) ? info.get(p).code : null;
      nodes.push(n);
    }
    return nodes;
  }

  // fallback when lists are not nested: infer hierarchy from label indentation
  function extractByIndent(root) {
    const els = [...root.querySelectorAll("[data-treedata-groupname],[data-treedata-groupcode]")];
    const stack = [];
    const nodes = [];
    els.forEach((el, i) => {
      const name = (el.getAttribute("data-treedata-groupname") || el.textContent || "").replace(/\s+/g, " ").trim();
      if (!name) return;
      const x = el.getBoundingClientRect().left;
      const node = { code: el.getAttribute("data-treedata-groupcode") || `node-${i}`, name, order: i, parent: null };
      while (stack.length && stack[stack.length - 1].x >= x) stack.pop();
      node.parent = stack.length ? stack[stack.length - 1].node.code : null;
      stack.push({ x, node });
      nodes.push(node);
    });
    return nodes;
  }

  // structure diagnostics: tag / class / attribute names only, no text
  function skeleton(root) {
    const sample = [];
    const walk = (el, depth) => {
      if (sample.length >= 40 || depth > 8) return;
      if (el !== root) {
        const attrs = [...el.attributes].map((a) => a.name).filter((n) => n !== "class" && n !== "style").join(",");
        const cls = typeof el.className === "string" && el.className.trim() ? "." + el.className.trim().split(/\s+/).join(".") : "";
        sample.push(`${"  ".repeat(depth)}${el.tagName.toLowerCase()}${cls}${attrs ? ` [${attrs}]` : ""}`);
      }
      for (const c of el.children) walk(c, depth + 1);
    };
    walk(root, -1);
    return sample;
  }

  function extract() {
    const root = document.querySelector("#Tree-dataList");
    if (!root) return null;
    let nodes = extractFromLists(root);
    let method = "lists";
    const flat = nodes.length > 3 && nodes.every((n) => !n.parent);
    if (!nodes.length || flat) {
      const alt = extractByIndent(root);
      if (alt.some((n) => n.parent)) { nodes = alt; method = "indent"; }
    }
    if (!nodes.length) return null;
    const company = document.querySelector("#Tree-companySelect option:checked")?.textContent?.trim()
      || document.querySelector("#Tree-companyName")?.textContent?.trim() || "";
    return {
      company,
      nodes,
      diag: { method, liCount: root.querySelectorAll("li").length, skeleton: skeleton(root) },
    };
  }

  function send() {
    const tree = extract();
    if (!tree) return;
    const key = JSON.stringify(tree.nodes) + tree.company;
    if (key === lastSent) return;
    lastSent = key;
    try { chrome.runtime.sendMessage({ type: "orgtree", payload: tree }); } catch {}
  }

  const schedule = () => { clearTimeout(timer); timer = setTimeout(send, 500); };
  new MutationObserver(schedule).observe(document.documentElement, { childList: true, subtree: true });
  schedule();
})();
