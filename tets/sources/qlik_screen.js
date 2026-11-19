// Filtre uygulanmis Qlik tablosunu EKRANDAN okur (insanin gordugu satirlar).
// Tablo kaydirilarak (sanal kaydirma) veya sayfa sayfa (sayfalama) tum satirlar toplanir.
async ({ oid, tableSel, maxRows }) => {
  const root =
    (oid && document.querySelector(`[tid="${oid}"], [data-qid="${oid}"]`)) || document.querySelector(tableSel);
  if (!root) throw new Error("Tablo ekranda bulunamadi");
  const txt = (e) => (e.innerText || e.textContent || "").replace(/\s+/g, " ").trim();
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  const headers = [...root.querySelectorAll("thead th, [role=columnheader]")].map(txt).filter((h, i, a) => h || a.length);
  if (!headers.length) throw new Error("Tablo basliklari ekrandan okunamadi");
  const seen = new Set();
  const rows = [];
  const grab = () => {
    let added = 0;
    root.querySelectorAll("tbody tr, [role=row]").forEach((tr) => {
      if (tr.querySelector("th, [role=columnheader]")) return;
      let cells = [...tr.querySelectorAll("td, [role=gridcell], [role=cell]")].map(txt);
      if (!cells.length) return;
      if (cells.length > headers.length) cells = cells.slice(cells.length - headers.length);
      if (cells.length < headers.length * 0.5) return;
      const k = cells.join("\u0001");
      if (seen.has(k)) return;
      seen.add(k);
      rows.push(cells);
      added++;
    });
    return added;
  };

  await sleep(300);
  grab();
  // sanal kaydirma
  const scroller = [root, ...root.querySelectorAll("*")].find((e) => {
    const s = getComputedStyle(e);
    return e.scrollHeight > e.clientHeight + 5 && /(auto|scroll)/.test(s.overflowY);
  });
  if (scroller) {
    let stable = 0;
    while (stable < 3 && rows.length < maxRows) {
      const before = scroller.scrollTop;
      scroller.scrollTop = before + Math.max(50, scroller.clientHeight * 0.8);
      await sleep(400);
      const n = grab();
      stable = n ? 0 : stable + 1;
      if (scroller.scrollTop === before && !n) break; // en alta gelindi
    }
    scroller.scrollTop = 0;
  }
  // sayfalama
  const nextSel = "button[aria-label*='Next' i], button[title*='Next' i], button[aria-label*='Sonraki' i], button[title*='Sonraki' i]";
  for (let i = 0; i < 500 && rows.length < maxRows; i++) {
    const next = root.querySelector(nextSel);
    if (!next || next.disabled || next.getAttribute("aria-disabled") === "true") break;
    next.click();
    await sleep(600);
    if (!grab()) break;
  }
  return { headers, rows, truncated: rows.length >= maxRows };
}
