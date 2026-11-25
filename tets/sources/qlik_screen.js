// Filtre uygulanmis Qlik tablosunu EKRANDAN okur (indirme yok).
// Okuyucu tarayicida window.__tets olarak kurulur; kaydirmayi Python tarafi GERCEK fare tekerlegiyle
// (veya varsa dogal kaydirma alaniyla) yapar ve her adimda "snap" ile gorunen hucreleri toplar.
// Her hucre kendi kolon basligiyla eslestirilir.
// action: init | snap | hstep | hreset | vstep | vreset | next | result
async ({ action, oid, tableSel, maxRows, waitMs, knownHeaders }) => {
  if (action !== "init") {
    if (!window.__tets) throw new Error("okuyucu kurulmamis");
    return window.__tets[action]();
  }
  const root =
    (oid && document.querySelector(`[tid="${oid}"], [data-qid="${oid}"]`)) || document.querySelector(tableSel);
  if (!root) throw new Error("Tablo ekranda bulunamadi");
  const txt = (e) => (e.innerText || e.textContent || "").replace(/\s+/g, " ").trim();
  const norm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const known = new Set((knownHeaders || []).map(norm));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  // Kaydirilabilir alanlar: tablonun ASIL kaydirma alani en buyuk olandir (kucuk ic alanlar degil)
  // tablo ve onu saran kutular (kaydirma alani tablonun disinda, saran kutuda olabilir)
  const ancestors = [];
  for (let p = root.parentElement, i = 0; p && i < 8 && p !== document.body; p = p.parentElement, i++) ancestors.push(p);
  const scrollers = (axis) =>
    [...ancestors, root, ...root.querySelectorAll("*")].filter((e) => {
      const s = getComputedStyle(e);
      return axis === "y"
        ? e.scrollHeight > e.clientHeight + 5 && /(auto|scroll|overlay)/.test(s.overflowY)
        : e.scrollWidth > e.clientWidth + 5 && /(auto|scroll|overlay)/.test(s.overflowX);
    });
  const biggest = (list, axis) =>
    list.sort((a, b) => (axis === "y" ? b.scrollHeight - a.scrollHeight : b.scrollWidth - a.scrollWidth))[0] || null;
  const hAll = scrollers("x");
  const vAll = scrollers("y");
  const hS = biggest([...hAll], "x");
  const vS = biggest([...vAll], "y");
  const hInfo = { adim: 0, alan: hAll.length, genislik: hS ? hS.scrollWidth : 0, gorunen: hS ? hS.clientWidth : 0 };
  // ana alani kaydirirken diger yatay alanlari (orn. ayri baslik alani) da ayni konuma getir
  const setLeft = (x) => {
    hS.scrollLeft = x;
    for (const o of hAll) if (o !== hS) o.scrollLeft = hS.scrollLeft;
  };

  let page = 0; // sayfalamali tablolarda sayfa numarasi (satir anahtarina girer)
  let mode = ""; // "yapi": th/role isaretleri var | "goz": yazilari konumundan oku
  const headerPos = new Map(); // baslik -> tablodaki yatay konum (siralama icin)
  const rows = new Map(); // satir anahtari -> Map(baslik -> deger)
  const isHeaderRow = (tr) => tr.querySelector("th, [role=columnheader]");
  const visible = (r) => r.width > 0 && r.height > 0;
  const TOTAL = /^(Totals?|Toplam|Genel toplam)$/i;
  const CELL = "th, td, [role=columnheader], [role=gridcell], [role=cell], [class*='cell' i], [class*='header' i]";

  // --- 1) Yapisal okuma: <th>/<td> veya role isaretleri
  const structural = () => {
    let heads = [...root.querySelectorAll("thead th, [role=columnheader]")]
      .map((th) => {
        const r = th.getBoundingClientRect();
        return { text: txt(th), left: r.left, right: r.right, top: r.top, ci: th.getAttribute("aria-colindex") };
      })
      .filter((h) => h.text && h.right > h.left && !TOTAL.test(h.text));
    // sadece en ustteki baslik satiri (alttaki "Totals" satiri baslik degildir)
    if (heads.length) {
      const minTop = Math.min(...heads.map((h) => h.top));
      heads = heads.filter((h) => h.top <= minTop + 6);
    }
    const lines = [];
    root.querySelectorAll("tbody tr, [role=row]").forEach((tr) => {
      if (isHeaderRow(tr)) return;
      const cells = [...tr.querySelectorAll("td, [role=gridcell], [role=cell]")];
      if (!cells.length || TOTAL.test(txt(cells[0]))) return;
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
    return { heads, lines: lines.filter((l) => !l.cells.some((c) => TOTAL.test(txt(c.el)))) };
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

  // gorunen basliklarin imzasi: yatay kayma olduysa degisir
  const sig = () => {
    const hs = [...root.querySelectorAll("thead th, [role=columnheader]")];
    const els = hs.length ? hs : leaves().filter((e) => known.has(norm(txt(e))));
    // konum tablonun kendisine gore: sayfanin tamami kaysa da degismez, sadece tablo ici kayarsa degisir
    const L = root.getBoundingClientRect().left;
    return els.map((e) => txt(e) + "@" + Math.round(e.getBoundingClientRect().left - L)).join("|");
  };

  window.__tets = {
    snap: () => ({ added: snapshot(), sig: sig(), rows: rows.size }),
    hstep: () => {
      if (!hS) return false;
      const before = hS.scrollLeft;
      setLeft(before + Math.max(50, hS.clientWidth * 0.6));
      return hS.scrollLeft > before;
    },
    hreset: () => { if (hS) setLeft(0); return true; },
    // Tablonun altindaki yatay kaydirma cubugu (yeri satir sayisina gore degisir): cubuk + tutulacak parca
    hbar: () => {
      const tb = root.getBoundingClientRect();
      const inArea = (r) => r.right > tb.left && r.left < tb.right && r.top >= tb.top - 5 && r.bottom <= tb.bottom + 40;
      const all = [...document.querySelectorAll("*")];
      const tracks = all.filter((e) => {
        const r = e.getBoundingClientRect();
        if (!(r.width >= tb.width * 0.3 && r.height >= 3 && r.height <= 26 && inArea(r))) return false;
        const cls = (e.className && e.className.baseVal !== undefined ? e.className.baseVal : e.className) || "";
        return /scroll|track|rail|bar/i.test(cls) || e.getAttribute("role") === "scrollbar";
      });
      let best = null;
      for (const t of tracks) {
        const tr = t.getBoundingClientRect();
        const kids = [...t.querySelectorAll("*")].map((k) => [k, k.getBoundingClientRect()])
          .filter(([k, r]) => r.width >= 12 && r.width < tr.width - 4 && Math.abs(r.height - tr.height) <= 8 && r.height >= 3);
        kids.sort((a, b) => b[1].width - a[1].width);
        const th = kids.length ? kids[0][1] : null;
        const cand = { track: { x: tr.left, y: tr.top, w: tr.width, h: tr.height },
                       thumb: th ? { x: th.left, y: th.top, w: th.width, h: th.height } : null };
        // tutulacak parcasi olan, en genis ve tablonun en altina yakin cubuk
        const score = (th ? 1e6 : 0) + tr.width + tr.top;
        if (!best || score > best.score) best = { ...cand, score, el: t, thEl: kids.length ? kids[0][0] : null };
      }
      if (!best) return null;
      // cubuk ekranin gorunen kisminin disindaysa gorunur hale getir, sonra yeniden olc
      best.el.scrollIntoView({ block: "nearest", inline: "nearest" });
      const tr = best.el.getBoundingClientRect();
      const th = best.thEl ? best.thEl.getBoundingClientRect() : null;
      return { track: { x: tr.left, y: tr.top, w: tr.width, h: tr.height },
               thumb: th ? { x: th.left, y: th.top, w: th.width, h: th.height } : null };
    },
    vstep: () => {
      if (!vS) return false;
      const before = vS.scrollTop;
      vS.scrollTop = before + Math.max(20, vS.clientHeight * 0.5); // kucuk adim: satir atlanmaz
      return vS.scrollTop > before;
    },
    vreset: () => { if (vS) vS.scrollTop = 0; return true; },
    next: () => {
      const next = root.querySelector(
        "button[aria-label*='Next' i], button[title*='Next' i], button[aria-label*='Sonraki' i], button[title*='Sonraki' i]");
      if (!next || next.disabled || next.getAttribute("aria-disabled") === "true") return false;
      next.click();
      page++;
      return true;
    },
    result: () => {
      const headers = [...headerPos.entries()].sort((a, b) => a[1] - b[1]).map((e) => e[0]);
      if (!headers.length) throw new Error("Tablo basliklari ekrandan okunamadi | " + diag());
      const out = [];
      for (const row of rows.values()) {
        const vals = headers.map((h) => row.get(h) ?? "");
        if (vals.some((v) => v !== "")) out.push(vals);
      }
      return { headers, rows: out, truncated: rows.size >= maxRows, hscroll: !!hS, vscroll: !!vS, mode, hInfo };
    },
  };
  await sleep(300);
  return { hasH: !!hS, hasV: !!vS };
}
