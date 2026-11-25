// Filtre uygulanmis Qlik tablosunu EKRANDAN okur (indirme yok).
// Tablo hem asagi (satirlar) hem saga (ekrana sigmayan kolonlar) kaydirilir; her hucre kendi kolon
// basligiyla eslestirilir. Sanal kaydirmali (sadece gorunen kismi ciziyor) tablolarda da calisir.
async ({ oid, tableSel, maxRows, waitMs, knownHeaders }) => {
  const root =
    (oid && document.querySelector(`[tid="${oid}"], [data-qid="${oid}"]`)) || document.querySelector(tableSel);
  if (!root) throw new Error("Tablo ekranda bulunamadi");
  const txt = (e) => (e.innerText || e.textContent || "").replace(/\s+/g, " ").trim();
  const norm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const known = new Set((knownHeaders || []).map(norm));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const scrollable = (axis) =>
    [root, ...root.querySelectorAll("*")].find((e) => {
      const s = getComputedStyle(e);
      return axis === "y"
        ? e.scrollHeight > e.clientHeight + 5 && /(auto|scroll)/.test(s.overflowY)
        : e.scrollWidth > e.clientWidth + 5 && /(auto|scroll)/.test(s.overflowX);
    });
  const vS = scrollable("y");
  const hS = scrollable("x");

  let page = 0; // sayfalamali tablolarda sayfa numarasi (satir anahtarina girer)
  let mode = ""; // "yapi": th/role isaretleri var | "goz": yazilari konumundan oku
  const headerPos = new Map(); // baslik -> tablodaki yatay konum (siralama icin)
  const rows = new Map(); // satir anahtari -> Map(baslik -> deger)
  const isHeaderRow = (tr) => tr.querySelector("th, [role=columnheader]");
  const visible = (r) => r.width > 0 && r.height > 0;
  const CELL = "th, td, [role=columnheader], [role=gridcell], [role=cell], [class*='cell' i], [class*='header' i]";

  // --- 1) Yapisal okuma: <th>/<td> veya role isaretleri
  const structural = () => {
    const heads = [...root.querySelectorAll("thead th, [role=columnheader]")]
      .map((th) => {
        const r = th.getBoundingClientRect();
        return { text: txt(th), left: r.left, right: r.right, ci: th.getAttribute("aria-colindex") };
      })
      .filter((h) => h.text && h.right > h.left);
    const lines = [];
    root.querySelectorAll("tbody tr, [role=row]").forEach((tr) => {
      if (isHeaderRow(tr)) return;
      const cells = [...tr.querySelectorAll("td, [role=gridcell], [role=cell]")];
      if (!cells.length) return;
      lines.push({ top: tr.getBoundingClientRect().top, ri: tr.getAttribute("aria-rowindex"),
                   cells: cells.map((td) => ({ el: td, r: td.getBoundingClientRect(), ci: td.getAttribute("aria-colindex") })) });
    });
    return { heads, lines };
  };

  // --- 2) Gozle okuma: bilinen kolon adlarini ekrandaki yazilardan bul, altindaki yazilari
  //        konumlarina gore satir/kolona yerlestir (DOM isaretlerine bagimli degil)
  const leaves = () =>
    [...root.querySelectorAll("*")].filter((e) => {
      if ([...e.children].some((c) => txt(c))) return false;
      const t = txt(e);
      return t && visible(e.getBoundingClientRect());
    });
  const byEye = () => {
    const ls = leaves();
    const hits = ls.filter((e) => known.has(norm(txt(e))));
    if (!hits.length) return { heads: [], lines: [] };
    // ayni yatay bantta en cok bilinen baslik olan satir = baslik satiri
    const bands = [];
    for (const e of hits) {
      const r = e.getBoundingClientRect();
      const cy = (r.top + r.bottom) / 2;
      let b = bands.find((x) => Math.abs(x.cy - cy) <= 8);
      if (!b) bands.push((b = { cy, els: [] }));
      b.els.push(e);
    }
    bands.sort((a, b) => b.els.length - a.els.length);
    const band = bands[0];
    const headEls = new Set();
    const heads = band.els.map((e) => {
      let c = e.closest(CELL);
      if (!c || !root.contains(c) || c.getBoundingClientRect().width > 700) c = e;
      headEls.add(c);
      const r = c.getBoundingClientRect();
      return { text: txt(e), left: r.left, right: r.right, ci: null, bottom: r.bottom };
    });
    const headBottom = Math.max(...heads.map((h) => h.bottom));
    const cellsAll = ls
      .filter((e) => ![...headEls].some((h) => h.contains(e)))
      .map((e) => ({ el: e, r: e.getBoundingClientRect(), ci: null }))
      .filter((c) => (c.r.top + c.r.bottom) / 2 > headBottom + 1);
    cellsAll.sort((a, b) => a.r.top - b.r.top);
    const lines = [];
    for (const c of cellsAll) {
      const cy = (c.r.top + c.r.bottom) / 2;
      let ln = lines.find((l) => Math.abs(l.cy - cy) <= 5);
      if (!ln) lines.push((ln = { cy, top: cy, ri: null, cells: [] }));
      ln.cells.push(c);
    }
    return { heads, lines };
  };

  const pickHead = (heads, c) => {
    if (c.ci !== null) {
      const h = heads.find((x) => x.ci === c.ci);
      if (h) return h;
    }
    let best = null, bestOv = 0;
    for (const h of heads) {
      const ov = Math.min(h.right, c.r.right) - Math.max(h.left, c.r.left);
      if (ov > bestOv) { bestOv = ov; best = h; }
    }
    if (best) return best;
    return heads.find((h) => Math.abs(h.left - c.r.left) <= 30) || null;
  };

  const snapshot = () => {
    let added = 0;
    const hx = hS ? hS.scrollLeft : 0;
    const vy = vS ? vS.scrollTop : 0;
    const baseTop = (vS || root).getBoundingClientRect().top;
    let res = null;
    if (mode !== "goz") {
      res = structural();
      if (res.heads.length >= 2) mode = "yapi";
      else if (!mode) res = null;
    }
    if (!res || mode === "goz") {
      res = byEye();
      if (res.heads.length && !mode) mode = "goz";
    }
    const { heads, lines } = res;
    for (const h of heads) if (!headerPos.has(h.text)) headerPos.set(h.text, h.left + hx);
    for (const ln of lines) {
      const key = page + ":" + (ln.ri !== null && ln.ri !== undefined ? "r" + ln.ri : "y" + Math.round(ln.top - baseTop + vy));
      let row = rows.get(key);
      if (!row) {
        if (rows.size >= maxRows) continue;
        row = new Map();
        rows.set(key, row);
        added++;
      }
      for (const c of ln.cells) {
        const h = pickHead(heads, c);
        // hucre henuz yuklenmeden bos okunduysa sonraki gorunuste gercek degerle guncellenir
        if (h && (!row.has(h.text) || row.get(h.text) === "")) row.set(h.text, txt(c.el));
      }
    }
    return added;
  };

  const diag = () => {
    const q = (s) => root.querySelectorAll(s).length;
    const sample = leaves().slice(0, 12).map((e) => txt(e).slice(0, 20)).join(" | ");
    return `nesne=${root.getAttribute("tid") || root.className}; th=${q("th")} td=${q("td")} tr=${q("tr")} ` +
      `row=${q("[role=row]")} columnheader=${q("[role=columnheader]")} gridcell=${q("[role=gridcell]")} ` +
      `div=${q("div")} canvas=${q("canvas")}; ekrandaki ilk yazilar: ${sample}`;
  };

  // Bir dikey konumda: tablo saga dogru kaydirilarak tum kolonlar okunur
  const sweepRight = async () => {
    let n = snapshot();
    if (!hS) return n;
    let guard = 0;
    while (hS.scrollLeft + hS.clientWidth < hS.scrollWidth - 2 && guard++ < 100) {
      hS.scrollLeft = hS.scrollLeft + Math.max(50, hS.clientWidth * 0.6);
      await sleep(waitMs);
      n += snapshot();
    }
    hS.scrollLeft = 0;
    await sleep(waitMs);
    return n;
  };

  await sleep(300);
  await sweepRight();
  if (vS) {
    let stable = 0;
    while (stable < 3 && rows.size < maxRows) {
      const before = vS.scrollTop;
      // kucuk adim: gorunen alanin yarisi, boylece hicbir satir atlanmaz
      vS.scrollTop = before + Math.max(20, vS.clientHeight * 0.5);
      await sleep(waitMs);
      const n = await sweepRight();
      stable = n ? 0 : stable + 1;
      if (vS.scrollTop === before) break; // en alta gelindi
    }
    vS.scrollTop = 0;
  }
  // Sayfalama olan tablolarda sonraki sayfalar
  const nextSel =
    "button[aria-label*='Next' i], button[title*='Next' i], button[aria-label*='Sonraki' i], button[title*='Sonraki' i]";
  for (let i = 0; i < 500 && rows.size < maxRows; i++) {
    const next = root.querySelector(nextSel);
    if (!next || next.disabled || next.getAttribute("aria-disabled") === "true") break;
    next.click();
    page++;
    await sleep(waitMs + 200);
    if (!(await sweepRight())) break;
  }

  const headers = [...headerPos.entries()].sort((a, b) => a[1] - b[1]).map((e) => e[0]);
  if (!headers.length) throw new Error("Tablo basliklari ekrandan okunamadi | " + diag());
  const out = [];
  for (const row of rows.values()) {
    const vals = headers.map((h) => row.get(h) ?? "");
    if (vals.some((v) => v !== "")) out.push(vals);
  }
  return { headers, rows: out, truncated: rows.size >= maxRows, hscroll: !!hS, vscroll: !!vS, mode };
}
