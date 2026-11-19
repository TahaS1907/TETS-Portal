// Qlik Sense Engine API (JSON-RPC over WebSocket) ile tablo verisi okuma.
// Giris yapilmis tarayici sayfasinin icinde calisir; oturum cerezi otomatik gonderilir.
// Ekrandaki menulere (Download/Indir) bagimli degildir.
async ({ wsUrl, appId, objectId, sheetId, selections, mustHave, maxRows, timeoutMs }) => {
  const log = [];
  const ws = new WebSocket(wsUrl);
  let seq = 0;
  const pending = {};
  const opened = new Promise((res, rej) => {
    ws.onopen = res;
    ws.onerror = () => rej(new Error("Qlik motoruna baglanilamadi: " + wsUrl));
  });
  ws.onmessage = (ev) => {
    const m = JSON.parse(ev.data);
    if (m.id && pending[m.id]) {
      const p = pending[m.id];
      delete pending[m.id];
      if (m.error) p.rej(new Error(`${m.error.message || "Qlik hatasi"} (${m.error.code}) ${m.error.parameter || ""}`));
      else p.res(m.result);
    }
  };
  const call = (handle, method, params) =>
    new Promise((res, rej) => {
      const id = ++seq;
      pending[id] = { res, rej };
      setTimeout(() => { if (pending[id]) { delete pending[id]; rej(new Error(method + " zaman asimi")); } }, timeoutMs);
      ws.send(JSON.stringify({ jsonrpc: "2.0", id, handle, method, params }));
    });
  const norm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");

  await opened;
  try {
    let doc;
    try {
      doc = (await call(-1, "OpenDoc", { qDocName: appId })).qReturn.qHandle;
    } catch (e) {
      doc = (await call(-1, "GetActiveDoc", {})).qReturn.qHandle; // uygulama bu oturumda zaten acik
    }

    const tableInfo = async (id) => {
      const h = (await call(doc, "GetObject", { qId: id })).qReturn.qHandle;
      const lay = (await call(h, "GetLayout", {})).qLayout;
      if (!lay.qHyperCube) return null;
      const hc = lay.qHyperCube;
      const titles = [...(hc.qDimensionInfo || []), ...(hc.qMeasureInfo || [])].map((d) => d.qFallbackTitle);
      return { id, h, type: lay.qInfo && lay.qInfo.qType, titles, hc };
    };

    // 1) Tablo nesnesini bul: kayitli id gecerliyse onu, degilse sayfadaki en uygun tabloyu kullan
    let table = null;
    if (objectId) {
      try { table = await tableInfo(objectId); } catch (e) { log.push("Kayitli nesne acilamadi: " + objectId); }
    }
    const fits = (t) => t && mustHave.every((c) => t.titles.some((x) => norm(x) === norm(c)));
    if (!fits(table)) {
      let ids = [];
      if (sheetId) {
        try {
          const sh = (await call(doc, "GetObject", { qId: sheetId })).qReturn.qHandle;
          const sl = (await call(sh, "GetLayout", {})).qLayout;
          ids = ((sl.qChildList && sl.qChildList.qItems) || []).map((i) => i.qInfo.qId);
        } catch (e) { log.push("Sayfa okunamadi: " + sheetId); }
      }
      if (!ids.length) {
        const infos = (await call(doc, "GetAllInfos", {})).qInfos || [];
        ids = infos.filter((i) => /table/i.test(i.qType)).map((i) => i.qId);
      }
      let best = null;
      for (const id of ids) {
        let t = null;
        try { t = await tableInfo(id); } catch (e) { continue; }
        if (!t) continue;
        const score = (fits(t) ? 1000 : 0) + t.titles.length;
        if (!best || score > best.score) best = { ...t, score };
      }
      if (!best) throw new Error("Sayfada tablo bulunamadi");
      table = best;
      log.push(`Tablo nesnesi: ${table.id} (${table.type}, ${table.titles.length} kolon)`);
    }

    // 2) Secimler: kolon basligindan gercek alan adini bul (master item olsa bile)
    const labelToField = {};
    [...(table.hc.qDimensionInfo || [])].forEach((d) => {
      const f = (d.qGroupFieldDefs || [])[0];
      if (f && !String(f).startsWith("=")) labelToField[norm(d.qFallbackTitle)] = f;
    });
    await call(doc, "ClearAll", { qLockedAlso: false });
    const before = (await call(table.h, "GetLayout", {})).qLayout.qHyperCube.qSize.qcy;
    let anyMiss = false;
    for (const [label, values] of Object.entries(selections)) {
      const field = labelToField[norm(label)] || label;
      const fh = (await call(doc, "GetField", { qFieldName: field })).qReturn.qHandle;
      if (!fh) throw new Error(`Qlik'te '${field}' alani bulunamadi`);
      const qFieldValues = [];
      for (const v of values) {
        qFieldValues.push({ qText: String(v) });
        const n = Number(String(v).replace(",", "."));
        if (String(v).trim() !== "" && !isNaN(n)) qFieldValues.push({ qText: String(v), qIsNumeric: true, qNumber: n });
      }
      const ok = (await call(fh, "SelectValues", { qFieldValues, qToggleMode: false, qSoftLock: false })).qReturn;
      log.push(`Secim ${field} = ${values.join(", ")} -> ${ok ? "ok" : "eslesme yok"}`);
      if (!ok) anyMiss = true;
    }

    // 3) Veri: secimlerden sonraki layout, 10.000 hucrelik sayfalar halinde
    const lay = (await call(table.h, "GetLayout", {})).qLayout;
    const hc = lay.qHyperCube;
    const headers = [...(hc.qDimensionInfo || []), ...(hc.qMeasureInfo || [])].map((d) => d.qFallbackTitle);
    if (Object.keys(selections).length && (anyMiss || (before > 0 && hc.qSize.qcy === before))) {
      // secim tutmadi: tum tabloyu dondurmek yerine bos sonuc
      log.push("Secilen deger tabloda yok -> bos sonuc");
      await call(doc, "ClearAll", { qLockedAlso: false });
      return { headers, rows: [], objectId: table.id, log };
    }
    const width = hc.qSize.qcx;
    const total = Math.min(hc.qSize.qcy, maxRows);
    if (hc.qSize.qcy > maxRows) log.push(`UYARI: ${hc.qSize.qcy} satirdan ilk ${maxRows} alindi`);
    const pageH = Math.max(1, Math.floor(10000 / Math.max(width, 1)));
    const rows = [];
    for (let top = 0; top < total; top += pageH) {
      const r = await call(table.h, "GetHyperCubeData", {
        qPath: "/qHyperCubeDef",
        qPages: [{ qTop: top, qLeft: 0, qWidth: width, qHeight: Math.min(pageH, total - top) }],
      });
      for (const row of r.qDataPages[0].qMatrix) rows.push(row.map((c) => (c.qIsNull ? "" : c.qText)));
    }
    await call(doc, "ClearAll", { qLockedAlso: false });
    return { headers, rows, objectId: table.id, log };
  } finally {
    ws.close();
  }
}
