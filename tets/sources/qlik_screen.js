// Filtre uygulanmis Qlik tablosunu EKRANDAN okur (indirme yok).
// Tablo hem asagi (satirlar) hem saga (ekrana sigmayan kolonlar) kaydirilir; her hucre kendi kolon
// basligiyla eslestirilir. Sanal kaydirmali (sadece gorunen kismi ciziyor) tablolarda da calisir.
async ({ oid, tableSel, maxRows, waitMs }) => {
  const root =
    (oid && document.querySelector(`[tid="${oid}"], [data-qid="${oid}"]`)) || document.querySelector(tableSel);
  if (!root) throw new Error("Tablo ekranda bulunamadi");
  const txt = (e) => (e.innerText || e.textContent || "").replace(/\s+/g, " ").trim();
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
  const headerPos = new Map(); // baslik -> tablodaki yatay konum (siralama icin)
  const rows = new Map(); // satir anahtari -> Map(baslik -> deger)
  const isHeaderRow = (tr) => tr.querySelector("th, [role=columnheader]");

  const snapshot = () => {
    let added = 0;
    const hx = hS ? hS.scrollLeft : 0;
    const vy = vS ? vS.scrollTop : 0;
    const baseTop = (vS || root).getBoundingClientRect().top;
    const heads = [...root.querySelectorAll("thead th, [role=columnheader]")]
      .map((th) => {
        const r = th.getBoundingClientRect();
        return { text: txt(th), left: r.left, right: r.right, ci: th.getAttribute("aria-colindex") };
      })
      .filter((h) => h.text && h.right > h.left);
    for (const h of heads) if (!headerPos.has(h.text)) headerPos.set(h.text, h.left + hx);
    root.querySelectorAll("tbody tr, [role=row]").forEach((tr) => {
      if (isHeaderRow(tr)) return;
      const cells = [...tr.querySelectorAll("td, [role=gridcell], [role=cell]")];
      if (!cells.length) return;
      const ri = tr.getAttribute("aria-rowindex");
      const key = page + ":" + (ri !== null ? "r" + ri : "y" + Math.round(tr.getBoundingClientRect().top - baseTop + vy));
      let row = rows.get(key);
      if (!row) {
        if (rows.size >= maxRows) return;
        row = new Map();
        rows.set(key, row);
        added++;
      }
      for (const td of cells) {
        const ci = td.getAttribute("aria-colindex");
        let h = ci !== null ? heads.find((x) => x.ci === ci) : null;
        if (!h) {
          const r = td.getBoundingClientRect();
          const cx = (r.left + r.right) / 2;
          h = heads.find((x) => cx >= x.left && cx <= x.right);
        }
        // hucre henuz yuklenmeden bos okunduysa sonraki gorunuste gercek degerle guncellenir
        if (h && (!row.has(h.text) || row.get(h.text) === "")) row.set(h.text, txt(td));
      }
    });
    return added;
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
  if (!headers.length) throw new Error("Tablo basliklari ekrandan okunamadi");
  const out = [];
  for (const row of rows.values()) {
    const vals = headers.map((h) => row.get(h) ?? "");
    if (vals.some((v) => v !== "")) out.push(vals);
  }
  return { headers, rows: out, truncated: rows.size >= maxRows, hscroll: !!hS, vscroll: !!vS };
}
