// Qlik Sense tablo verisini menuye dokunmadan okur. Giris yapilmis sayfanin icinde calisir.
// mode = "client": acik Qlik sayfasinin kendi baglantisi (js/qlik -> enigmaModel) kullanilir.
// mode = "ws":     verilen WebSocket adresiyle Engine API'ye (JSON-RPC) dogrudan baglanilir.
async ({ mode, wsUrl, appId, objectId, sheetId, selections, mustHave, maxRows, timeoutMs }) => {
  const log = [];
  const norm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const T = (p, what) =>
    Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error(what + " zaman asimi")), timeoutMs))]);
  let api;
  let cleanup = () => {};

  if (mode === "client") {
    if (typeof window.require !== "function") throw new Error("Qlik istemci API'si sayfada yok");
    const qlik = await T(new Promise((res, rej) => window.require(["js/qlik"], res, rej)), "js/qlik");
    const app = qlik.currApp();
    if (!app) throw new Error("Acik Qlik uygulamasi bulunamadi");
    if (appId && app.id && app.id !== appId) throw new Error(`Sayfada farkli uygulama acik (${app.id})`);
    const doc = app.model && app.model.enigmaModel;
    if (!doc) throw new Error("Qlik istemcisinde enigmaModel yok");
    api = {
      getObject: async (id) => {
        const o = await T(doc.getObject(id), "getObject");
        if (!o) throw new Error("Nesne yok: " + id);
        return {
          layout: () => T(o.getLayout(), "getLayout"),
          data: (pages) => T(o.getHyperCubeData("/qHyperCubeDef", pages), "getHyperCubeData"),
        };
      },
      clearAll: () => T(doc.clearAll(false), "clearAll"),
      select: async (field, vals) => {
        const f = await T(doc.getField(field), "getField");
        if (!f) throw new Error(`Qlik'te '${field}' alani bulunamadi`);
        return T(f.selectValues(vals, false, false), "selectValues");
      },
      allInfos: () => T(doc.getAllInfos(), "getAllInfos"),
    };
  } else {
    const ws = new WebSocket(wsUrl);
    let seq = 0;
    const pending = {};
    await T(new Promise((res, rej) => {
      ws.onopen = res;
      ws.onerror = () => rej(new Error("Qlik motoruna baglanilamadi: " + wsUrl));
    }), "WebSocket");
    ws.onmessage = (ev) => {
      let m;
      try { m = JSON.parse(ev.data); } catch (e) { return; }
      if (m.id && pending[m.id]) {
        const p = pending[m.id];
        delete pending[m.id];
        if (m.error) p.rej(new Error(`${m.error.message || "Qlik hatasi"} (${m.error.code}) ${m.error.parameter || ""}`));
        else p.res(m.result);
      }
    };
    cleanup = () => ws.close();
    const call = (handle, method, params) =>
      T(new Promise((res, rej) => {
        const id = ++seq;
        pending[id] = { res, rej };
        ws.send(JSON.stringify({ jsonrpc: "2.0", id, handle, method, params }));
      }), method);
    let doc;
    try {
      doc = (await call(-1, "OpenDoc", { qDocName: appId })).qReturn.qHandle;
    } catch (e) {
      doc = (await call(-1, "GetActiveDoc", {})).qReturn.qHandle; // uygulama bu oturumda zaten acik
    }
    api = {
      getObject: async (id) => {
        const h = (await call(doc, "GetObject", { qId: id })).qReturn.qHandle;
        if (!h) throw new Error("Nesne yok: " + id);
        return {
          layout: async () => (await call(h, "GetLayout", {})).qLayout,
          data: async (pages) => (await call(h, "GetHyperCubeData", { qPath: "/qHyperCubeDef", qPages: pages })).qDataPages,
        };
      },
      clearAll: () => call(doc, "ClearAll", { qLockedAlso: false }),
      select: async (field, vals) => {
        const fh = (await call(doc, "GetField", { qFieldName: field })).qReturn.qHandle;
        if (!fh) throw new Error(`Qlik'te '${field}' alani bulunamadi`);
        return (await call(fh, "SelectValues", { qFieldValues: vals, qToggleMode: false, qSoftLock: false })).qReturn;
      },
      allInfos: async () => (await call(doc, "GetAllInfos", {})).qInfos,
    };
  }

  try {
    const tableInfo = async (id) => {
      const o = await api.getObject(id);
      const lay = await o.layout();
      if (!lay || !lay.qHyperCube) return null;
      const hc = lay.qHyperCube;
      const titles = [...(hc.qDimensionInfo || []), ...(hc.qMeasureInfo || [])].map((d) => d.qFallbackTitle);
      return { id, o, type: lay.qInfo && lay.qInfo.qType, titles, hc };
    };

    // 1) Tablo nesnesi: kayitli id uygunsa o, degilse sayfadaki en uygun tablo
    let table = null;
    if (objectId) {
      try { table = await tableInfo(objectId); } catch (e) { log.push("Kayitli nesne acilamadi: " + objectId); }
    }
    const fits = (t) => t && mustHave.every((c) => t.titles.some((x) => norm(x) === norm(c)));
    if (!fits(table)) {
      let ids = [];
      if (sheetId) {
        try {
          const sh = await api.getObject(sheetId);
          const sl = await sh.layout();
          ids = ((sl.qChildList && sl.qChildList.qItems) || []).map((i) => i.qInfo.qId);
        } catch (e) { log.push("Sayfa okunamadi: " + sheetId); }
      }
      if (!ids.length) {
        const infos = (await api.allInfos()) || [];
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
    (table.hc.qDimensionInfo || []).forEach((d) => {
      const f = (d.qGroupFieldDefs || [])[0];
      if (f && !String(f).startsWith("=")) labelToField[norm(d.qFallbackTitle)] = f;
    });
    await api.clearAll();
    const before = (await table.o.layout()).qHyperCube.qSize.qcy;
    let anyMiss = false;
    for (const [label, values] of Object.entries(selections)) {
      const field = labelToField[norm(label)] || label;
      const vals = [];
      for (const v of values) {
        vals.push({ qText: String(v) });
        const n = Number(String(v).replace(",", "."));
        if (String(v).trim() !== "" && !isNaN(n)) vals.push({ qText: String(v), qIsNumeric: true, qNumber: n });
      }
      const ok = await api.select(field, vals);
      log.push(`Secim ${field} = ${values.join(", ")} -> ${ok ? "ok" : "eslesme yok"}`);
      if (!ok) anyMiss = true;
    }

    // 3) Veri: secimlerden sonraki tablo, 10.000 hucrelik sayfalar halinde
    const hc = (await table.o.layout()).qHyperCube;
    const headers = [...(hc.qDimensionInfo || []), ...(hc.qMeasureInfo || [])].map((d) => d.qFallbackTitle);
    if (Object.keys(selections).length && (anyMiss || (before > 0 && hc.qSize.qcy === before))) {
      log.push("Secilen deger tabloda yok -> bos sonuc");
      await api.clearAll();
      return { headers, rows: [], objectId: table.id, log };
    }
    const width = hc.qSize.qcx;
    const total = Math.min(hc.qSize.qcy, maxRows);
    if (hc.qSize.qcy > maxRows) log.push(`UYARI: ${hc.qSize.qcy} satirdan ilk ${maxRows} alindi`);
    const pageH = Math.max(1, Math.floor(10000 / Math.max(width, 1)));
    const rows = [];
    for (let top = 0; top < total; top += pageH) {
      const pages = await table.o.data([{ qTop: top, qLeft: 0, qWidth: width, qHeight: Math.min(pageH, total - top) }]);
      for (const row of pages[0].qMatrix) rows.push(row.map((c) => (c.qIsNull ? "" : c.qText)));
    }
    await api.clearAll();
    return { headers, rows, objectId: table.id, log };
  } finally {
    cleanup();
  }
}
