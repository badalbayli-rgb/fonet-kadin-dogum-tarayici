(() => {
  "use strict";

  const APP_KEY = "__FONET_KADIN_DOGUM_TARAYICI__";
  const VERSION = "1.1.0";
  if (window[APP_KEY]?.destroy) window[APP_KEY].destroy();

  const state = {
    version: VERSION,
    sourceOperationCount: 0,
    operations: [],
    results: [],
    errors: [],
    running: false,
    paused: false,
    stopped: false,
    processed: 0,
    concurrency: 6,
    panel: null
  };
  window[APP_KEY] = state;

  const clean = value => String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const trLower = value => clean(value).toLocaleLowerCase("tr-TR");
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const clip = (value, size = 300) => clean(value).slice(0, size);
  const escapeHtml = value => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
  const htmlToText = value => {
    const box = document.createElement("div");
    box.innerHTML = String(value || "");
    return clean(box.textContent || box.innerText || "");
  };

  function genderText(value) {
    if (value && typeof value === "object") {
      return clean(value.adi || value.ad || value.aciklama || value.kodu || value.kod || value.value || "");
    }
    return clean(value);
  }

  function isFemaleGender(value) {
    const gender = trLower(genderText(value)).replace(/[^a-z0-9ğüşöçı]/g, "");
    return gender === "kadın" || gender === "kadin" || gender === "k" ||
      gender === "female" || gender === "f" || gender === "2";
  }

  function readPath(object, path) {
    return String(path).split(".").reduce((value, key) => value?.[key], object);
  }

  function flattenObject(object, prefix = "", depth = 0, out = {}) {
    if (!object || typeof object !== "object" || depth > 7) return out;
    for (const [key, value] of Object.entries(object)) {
      const path = prefix ? `${prefix}.${key}` : key;
      if (value == null) continue;
      if (typeof value === "object") flattenObject(value, path, depth + 1, out);
      else out[path] = value;
    }
    return out;
  }

  function normalizedKey(key) {
    return String(key).replace(/[^a-zA-Z0-9ğüşöçıİĞÜŞÖÇ]/g, "").toLocaleLowerCase("tr-TR");
  }

  function findValue(data, candidates) {
    for (const path of candidates) {
      const value = readPath(data, path);
      if (value != null && clean(value)) return value;
    }
    const flat = flattenObject(data);
    const wanted = new Set(candidates.map(normalizedKey));
    for (const [path, value] of Object.entries(flat)) {
      const last = path.split(".").pop();
      if (wanted.has(normalizedKey(last)) && clean(value)) return value;
    }
    return "";
  }

  function deepFindId(object, predicate, depth = 0, seen = new Set()) {
    if (!object || typeof object !== "object" || depth > 7 || seen.has(object)) return "";
    seen.add(object);
    try {
      if (predicate(object) && object.id != null) return clean(object.id);
    } catch {}
    for (const value of Object.values(object)) {
      const found = deepFindId(value, predicate, depth + 1, seen);
      if (found) return found;
    }
    return "";
  }

  function inferIds(data) {
    const hastaGelisId = clean(findValue(data, [
      "hastaGelisId", "HASTA_GELIS_ID", "HASTAGELISID", "gelisId",
      "hastaGelis.id", "birimSevk.hastaGelis.id", "hastaBirimSevk.hastaGelis.id"
    ])) || deepFindId(data, x => x.hasta && (x.kodu || x.muracaatTarihi || x.kabulTarihi));
    const hastaId = clean(findValue(data, [
      "hastaId", "HASTA_ID", "HASTAID", "hasta.id", "hastaGelis.hasta.id",
      "birimSevk.hastaGelis.hasta.id"
    ])) || deepFindId(data, x => x.kimlik || x.tcKimlikNo || x.kimlikNo);
    return { hastaGelisId, hastaId };
  }

  function operationFromRecord(record, grid) {
    const data = record?.data || record || {};
    const ids = inferIds(data);
    const first = clean(findValue(data, ["adi", "ADI", "ad", "HASTA_ADI", "hasta.adi", "hastaGelis.hasta.adi"]));
    const last = clean(findValue(data, ["soyadi", "SOYADI", "soyad", "HASTA_SOYADI", "hasta.soyadi", "hastaGelis.hasta.soyadi"]));
    let patientName = clean(findValue(data, [
      "adiSoyadi", "adSoyad", "ADSOYAD", "AD_SOYAD", "HASTA_ADI_SOYADI",
      "hasta.adiSoyadi", "hastaGelis.hasta.adiSoyadi"
    ]));
    if (!patientName) patientName = clean(`${first} ${last}`);

    return {
      sourceGrid: grid?.id || "",
      operationNo: clean(findValue(data, ["islemNo", "ISLEM_NO", "ISLEMNO", "ameliyatNo", "AMELIYAT_NO", "protokolNo"])),
      operationDate: clean(findValue(data, ["ameliyatTarihi", "AMELIYAT_TARIHI", "islemTarihi", "ISTEM_TARIHI", "tarih", "baslamaTarihi"])),
      operationName: clean(findValue(data, ["ameliyatAdi", "AMELIYAT_ADI", "ameliyat", "AMELIYAT", "hizmetAdi", "HIZMET_ADI", "islemAdi"])),
      patientName,
      identityNo: clean(findValue(data, ["kimlikNo", "KIMLIK_NO", "tcKimlikNo", "TC_KIMLIK_NO", "hasta.kimlikNo", "hastaGelis.hasta.kimlikNo"])),
      gender: genderText(findValue(data, [
        "cinsiyetAdi", "CINSIYET_ADI", "cinsiyet", "CINSIYET", "cinsiyeti",
        "hasta.cinsiyet.adi", "hasta.cinsiyet.kodu", "hasta.cinsiyet",
        "hastaGelis.hasta.cinsiyet.adi", "hastaGelis.hasta.cinsiyet.kodu", "hastaGelis.hasta.cinsiyet"
      ])),
      birthDate: clean(findValue(data, ["dogumTarihi", "DOGUM_TARIHI", "hasta.dogumTarihi", "hastaGelis.hasta.dogumTarihi"])),
      hastaGelisId: ids.hastaGelisId,
      hastaId: ids.hastaId,
      raw: data
    };
  }

  function storeRecords(store) {
    const records = [];
    try { store?.each?.(record => records.push(record)); } catch {}
    try { if (!records.length && store?.data?.items) records.push(...store.data.items); } catch {}
    try { if (!records.length && store?.getRange) records.push(...store.getRange()); } catch {}
    return records;
  }

  function scoreGrid(grid, records) {
    if (!records.length) return -1;
    const keys = Object.keys(records[0]?.data || {}).join(" ");
    const sample = records.slice(0, 8).map(r => JSON.stringify(r.data || {}).slice(0, 2500)).join(" ");
    const text = trLower(`${grid?.title || ""} ${grid?.itemId || ""} ${keys} ${sample}`);
    let score = Math.min(records.length, 2500) / 20;
    if (/ameliyat|operasyon|appendektomi|cerrahi/.test(text)) score += 500;
    if (/işlemno|islemno|ameliyattarihi|ameliyatadi/.test(text)) score += 300;
    if (/hizmet listesi|laboratuvar|radyoloji/.test(text)) score -= 500;
    return score;
  }

  function collectOperations() {
    if (!window.Ext?.ComponentQuery) throw new Error("ExtJS bulunamadı. Script HBYS ana bağlamında çalıştırılmalı.");
    const candidates = [];
    for (const grid of Ext.ComponentQuery.query("gridpanel, grid")) {
      let store;
      try { store = grid.getStore?.(); } catch { continue; }
      const records = storeRecords(store);
      const score = scoreGrid(grid, records);
      if (score > 0) candidates.push({ grid, records, score });
    }
    candidates.sort((a, b) => b.score - a.score);
    const selected = candidates[0];
    if (!selected) throw new Error("Ameliyat veri deposu bulunamadı. Önce ameliyat sorgusunu çalıştırın.");

    const operations = selected.records.map(record => operationFromRecord(record, selected.grid));
    const meaningful = operations.filter(x => x.patientName || x.identityNo || x.hastaGelisId);
    if (!meaningful.length) throw new Error("Seçilen tabloda hasta kayıtları bulunamadı.");
    const femaleOperations = meaningful.filter(x => isFemaleGender(x.gender));
    state.sourceOperationCount = meaningful.length;
    state.operations = femaleOperations;
    if (!femaleOperations.length) {
      throw new Error("Ameliyat listesi bulundu ancak kadın cinsiyetli kayıt saptanamadı. Cinsiyet alanı HBYS kaydında görünür olmalı.");
    }
    return {
      gridId: selected.grid?.id || "-",
      count: femaleOperations.length,
      sourceCount: meaningful.length,
      score: Math.round(selected.score)
    };
  }

  function baseUrl() {
    return `${location.origin}/hbys-rs/hbys`;
  }

  async function apiJson(path) {
    const separator = path.includes("?") ? "&" : "?";
    const response = await fetch(`${baseUrl()}${path}${separator}_dc=${Date.now()}`, {
      credentials: "include",
      headers: { Accept: "application/json, text/plain, */*" }
    });
    const text = await response.text();
    if (!response.ok) throw new Error(`HTTP ${response.status}: ${clip(text, 120)}`);
    try { return JSON.parse(text); }
    catch { throw new Error(`JSON okunamadı: ${clip(text, 120)}`); }
  }

  function consultationFromRaw(raw) {
    const answer = clean([
      htmlToText(raw?.sonucAciklama || ""),
      htmlToText(raw?.sonucAciklama2 || "")
    ].filter(Boolean).join(" "));
    return {
      consultationId: clean(raw?.id),
      consultationDate: clean(raw?.birimSevk?.sevkTarihi || raw?.etar || raw?.tarih || ""),
      consultationUnit: clean(raw?.birimSevk?.birim?.adi || raw?.birim?.adi || raw?.birimAdi || ""),
      requestingUnit: clean(raw?.isteyenBirim?.adi || raw?.istekBirim?.adi || raw?.istemYapanBirim?.adi || ""),
      requestingDoctor: clean(raw?.isteyenPersonel?.adiSoyadi || raw?.istemYapanPersonel?.adiSoyadi || raw?.doktor?.adiSoyadi || ""),
      requestReason: clean(htmlToText(raw?.istemSebebi || raw?.istekAciklama || raw?.aciklama || "")),
      answer,
      status: clean(raw?.durum || raw?.durumAdi || "")
    };
  }

  function isGynecology(consult) {
    return /kadın hastalıkları ve doğum|kadın doğum|jinekoloji|obstetri/i.test(`${consult.consultationUnit} ${consult.requestReason} ${consult.answer}`);
  }

  async function scanOperation(operation) {
    if (!operation.hastaGelisId) throw new Error("hastaGelisId bulunamadı");
    const payload = await apiJson(`/Poliklinik/Poliklinik/getHastaGelisKonsultasyonList/${encodeURIComponent(operation.hastaGelisId)}/1`);
    const rows = Array.isArray(payload?.data) ? payload.data : [];
    return rows.map(consultationFromRaw).filter(isGynecology).map(consult => ({ ...operation, ...consult }));
  }

  async function waitWhilePaused() {
    while (state.paused && !state.stopped) await sleep(250);
  }

  async function worker(queue) {
    while (queue.length && !state.stopped) {
      await waitWhilePaused();
      if (state.stopped) return;
      const operation = queue.shift();
      if (!operation) return;
      try {
        const found = await scanOperation(operation);
        state.results.push(...found);
      } catch (error) {
        state.errors.push({
          operationNo: operation.operationNo,
          patientName: operation.patientName,
          hastaGelisId: operation.hastaGelisId,
          error: clean(error?.message || error)
        });
      } finally {
        state.processed++;
        render();
      }
    }
  }

  async function startScan() {
    if (state.running) return;
    if (!state.operations.length) {
      try { collectOperations(); }
      catch (error) { setMessage(error.message, true); return; }
    }
    state.running = true;
    state.paused = false;
    state.stopped = false;
    state.processed = 0;
    state.results = [];
    state.errors = [];
    render();
    const queue = [...state.operations];
    await Promise.all(Array.from({ length: state.concurrency }, () => worker(queue)));
    state.running = false;
    render();
    setMessage(state.stopped ? "Tarama durduruldu." : "Tarama tamamlandı.");
  }

  function uniqueResults() {
    const seen = new Set();
    return state.results.filter(item => {
      const key = [item.operationNo, item.hastaGelisId, item.consultationId, item.consultationDate].join("|");
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function csvEscape(value) {
    return `"${String(value ?? "").replace(/"/g, '""')}"`;
  }

  function downloadCsv() {
    const headers = [
      "Ameliyat No", "Ameliyat Tarihi", "Ameliyat", "Hasta", "TC Kimlik No", "Cinsiyet", "Doğum Tarihi",
      "Hasta Geliş ID", "Konsültasyon Tarihi", "Kadın Doğum Birimi", "İsteyen Birim", "İsteyen Doktor",
      "İstem Nedeni", "Konsültasyon Yanıtı", "Durum"
    ];
    const rows = uniqueResults().map(x => [
      x.operationNo, x.operationDate, x.operationName, x.patientName, x.identityNo, x.gender, x.birthDate,
      x.hastaGelisId, x.consultationDate, x.consultationUnit, x.requestingUnit, x.requestingDoctor,
      x.requestReason, x.answer, x.status
    ]);
    const csv = "\ufeff" + [headers, ...rows].map(row => row.map(csvEscape).join(";")).join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `fonet-kadin-dogum-konsultasyonlari-${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 3000);
  }

  function setMessage(message, isError = false) {
    const el = document.getElementById("fkd-message");
    if (!el) return;
    el.textContent = message;
    el.style.color = isError ? "#b91c1c" : "#334155";
  }

  function renderRows() {
    const tbody = document.getElementById("fkd-results");
    if (!tbody) return;
    const rows = uniqueResults().slice(-100).reverse();
    tbody.innerHTML = rows.map(x => `
      <tr>
        <td>${escapeHtml(x.patientName || "-")}</td>
        <td>${escapeHtml(x.operationDate || "-")}</td>
        <td>${escapeHtml(x.consultationDate || "-")}</td>
        <td>${escapeHtml(x.consultationUnit || "-")}</td>
        <td title="${escapeHtml(x.requestReason)}">${escapeHtml(clip(x.requestReason, 90) || "-")}</td>
        <td title="${escapeHtml(x.answer)}">${escapeHtml(clip(x.answer, 130) || "-")}</td>
      </tr>`).join("");
  }

  function render() {
    const total = state.operations.length;
    const percent = total ? Math.round(state.processed / total * 100) : 0;
    const stats = document.getElementById("fkd-stats");
    const progress = document.getElementById("fkd-progress-bar");
    if (stats) stats.textContent = `Toplam ameliyat: ${state.sourceOperationCount} | Kadın: ${total} | İşlenen kadın: ${state.processed} | Kadın doğum konsültasyonu: ${uniqueResults().length} | Hata: ${state.errors.length}`;
    if (progress) progress.style.width = `${percent}%`;
    const start = document.getElementById("fkd-start");
    const pause = document.getElementById("fkd-pause");
    const stop = document.getElementById("fkd-stop");
    if (start) start.disabled = state.running || !total;
    if (pause) {
      pause.disabled = !state.running;
      pause.textContent = state.paused ? "Devam Et" : "Duraklat";
    }
    if (stop) stop.disabled = !state.running;
    renderRows();
  }

  function mount() {
    const old = document.getElementById("fonet-kd-app");
    if (old) old.remove();
    const panel = document.createElement("section");
    panel.id = "fonet-kd-app";
    panel.innerHTML = `
      <style>
        #fonet-kd-app{position:fixed;inset:16px;z-index:2147483647;background:#f8fafc;color:#0f172a;border:1px solid #94a3b8;border-radius:14px;box-shadow:0 18px 60px #0008;font:13px Arial,sans-serif;display:flex;flex-direction:column;overflow:hidden}
        #fonet-kd-app *{box-sizing:border-box} #fonet-kd-app header{padding:14px 18px;background:#0f4c81;color:#fff;display:flex;align-items:center;justify-content:space-between}
        #fonet-kd-app h1{font-size:18px;margin:0} #fonet-kd-app .body{padding:14px;display:flex;flex-direction:column;gap:10px;min-height:0;flex:1}
        #fonet-kd-app button{border:0;border-radius:7px;padding:9px 13px;font-weight:700;cursor:pointer;background:#0b79bd;color:#fff;margin-right:6px}
        #fonet-kd-app button.alt{background:#475569} #fonet-kd-app button.stop{background:#b91c1c} #fonet-kd-app button:disabled{opacity:.4;cursor:not-allowed}
        #fonet-kd-app .progress{height:9px;background:#dbe5ec;border-radius:6px;overflow:hidden} #fkd-progress-bar{height:100%;width:0;background:#16a34a;transition:width .2s}
        #fonet-kd-app .table-wrap{overflow:auto;flex:1;background:#fff;border:1px solid #cbd5e1;border-radius:8px}
        #fonet-kd-app table{border-collapse:collapse;width:100%;font-size:12px} #fonet-kd-app th{position:sticky;top:0;background:#e2e8f0;text-align:left;padding:8px;border-bottom:1px solid #94a3b8}
        #fonet-kd-app td{padding:7px;border-bottom:1px solid #e2e8f0;vertical-align:top;max-width:260px} #fkd-message{min-height:18px}
      </style>
      <header><h1>FONET Kadın Doğum Konsültasyon Tarayıcı <small>v${VERSION}</small></h1><button id="fkd-close" class="alt">Kapat</button></header>
      <div class="body">
        <div>
          <button id="fkd-find">Listeyi Bul</button><button id="fkd-start" disabled>Taramayı Başlat</button>
          <button id="fkd-pause" class="alt" disabled>Duraklat</button><button id="fkd-stop" class="stop" disabled>Durdur</button>
          <button id="fkd-csv" class="alt">CSV İndir</button>
        </div>
        <div id="fkd-message">Önce HBYS ameliyat sorgusunu çalıştırın, sonra Listeyi Bul'a basın.</div>
        <strong id="fkd-stats">Toplam ameliyat: 0 | Kadın: 0 | İşlenen kadın: 0 | Kadın doğum konsültasyonu: 0 | Hata: 0</strong>
        <div class="progress"><div id="fkd-progress-bar"></div></div>
        <div class="table-wrap"><table><thead><tr><th>Hasta</th><th>Ameliyat</th><th>Konsültasyon</th><th>Birim</th><th>İstem nedeni</th><th>Yanıt</th></tr></thead><tbody id="fkd-results"></tbody></table></div>
      </div>`;
    document.documentElement.appendChild(panel);
    state.panel = panel;

    document.getElementById("fkd-find").onclick = () => {
      try {
        const found = collectOperations();
        setMessage(`${found.sourceCount} ameliyat kaydının ${found.count} tanesi kadın. Tarama yalnızca bu kadın hastalarda yapılacak (tablo: ${found.gridId}).`);
        render();
      } catch (error) { setMessage(error.message, true); }
    };
    document.getElementById("fkd-start").onclick = startScan;
    document.getElementById("fkd-pause").onclick = () => { state.paused = !state.paused; render(); };
    document.getElementById("fkd-stop").onclick = () => { state.stopped = true; state.paused = false; render(); };
    document.getElementById("fkd-csv").onclick = downloadCsv;
    document.getElementById("fkd-close").onclick = () => panel.remove();
    render();
  }

  state.destroy = () => {
    state.stopped = true;
    document.getElementById("fonet-kd-app")?.remove();
    delete window[APP_KEY];
  };

  mount();
})();
