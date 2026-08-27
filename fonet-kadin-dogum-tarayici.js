(() => {
  "use strict";

  const APP_KEY = "__FONET_KADIN_DOGUM_TARAYICI__";
  const VERSION = "1.3.3";
  if (window[APP_KEY]?.destroy) window[APP_KEY].destroy();

  const state = {
    version: VERSION,
    sourceOperationCount: 0,
    genderProcessed: 0,
    genderResolved: false,
    genderFilter: "all",
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
    return gender === "k" || gender === "f" || gender === "2" ||
      gender.includes("kadın") || gender.includes("kadin") || gender.includes("female");
  }

  function isMaleGender(value) {
    const gender = trLower(genderText(value)).replace(/[^a-z0-9ğüşöçı]/g, "");
    return gender === "e" || gender === "m" || gender === "1" ||
      gender.includes("erkek") || (gender.includes("male") && !gender.includes("female"));
  }

  function findGenderAnywhere(data) {
    const direct = findValue(data, [
      "cinsiyetAdi", "CINSIYET_ADI", "cinsiyet", "CINSIYET", "cinsiyeti",
      "yasCinsiyet", "YAS_CINSIYET", "yasVeCinsiyet", "YAS_VE_CINSIYET",
      "hasta.cinsiyet.adi", "hasta.cinsiyet.kodu", "hasta.cinsiyet",
      "hastaGelis.hasta.cinsiyet.adi", "hastaGelis.hasta.cinsiyet.kodu", "hastaGelis.hasta.cinsiyet"
    ]);
    if (genderText(direct)) return genderText(direct);
    for (const value of Object.values(flattenObject(data))) {
      const text = trLower(value);
      if (text.includes("kadın") || text.includes("kadin") || text.includes("female")) return clean(value);
      if (text.includes("erkek") || text.includes("male")) return clean(value);
    }
    return "";
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
    const birimSevkId = clean(findValue(data, [
      "birimSevkId", "BIRIM_SEVK_ID", "BIRIMSEVKID", "hastaBirimSevkId",
      "HASTA_BIRIM_SEVK_ID", "birimSevk.id", "hastaBirimSevk.id"
    ])) || deepFindId(data, x => x.hastaGelis && (x.birim || x.sevkTarihi));
    return { hastaGelisId, hastaId, birimSevkId };
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
      phone: clean(findValue(data, [
        "telefon", "TELEFON", "telefonNo", "TELEFON_NO", "cepTelefonu", "CEP_TELEFONU", "cepTelefon",
        "gsm", "GSM", "hasta.telefon", "hasta.cepTelefonu", "hasta.kimlik.telefon", "hasta.kimlik.cepTelefonu",
        "hastaGelis.hasta.telefon", "hastaGelis.hasta.cepTelefonu", "hastaGelis.hasta.kimlik.cepTelefonu"
      ])),
      gender: findGenderAnywhere(data),
      birthDate: clean(findValue(data, ["dogumTarihi", "DOGUM_TARIHI", "hasta.dogumTarihi", "hastaGelis.hasta.dogumTarihi"])),
      hastaGelisId: ids.hastaGelisId,
      hastaId: ids.hastaId,
      birimSevkId: ids.birimSevkId,
      sourceRecord: record,
      sourceGrid: grid,
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
    state.sourceOperationCount = meaningful.length;
    state.genderProcessed = 0;
    state.genderResolved = false;
    state.operations = meaningful;
    return {
      gridId: selected.grid?.id || "-",
      count: meaningful.length,
      sourceCount: meaningful.length,
      score: Math.round(selected.score)
    };
  }

  function genderFromPatientPayload(payload) {
    const root = payload?.data || payload || {};
    const sevk = root.hastaBirimSevk || root.birimSevk || root;
    const hasta = sevk?.hastaGelis?.hasta || root?.hastaGelis?.hasta || root?.hasta || {};
    return genderText(hasta?.kimlik?.cinsiyet?.adi || hasta?.kimlik?.cinsiyet || hasta?.cinsiyet?.adi || hasta?.cinsiyet || "");
  }

  async function genderFromService(operation) {
    if (!operation.birimSevkId) return "";
    try {
      const payload = await apiJson(`/Tibbi/HastaBirimSevk/getSevkUyariInfo/${encodeURIComponent(operation.birimSevkId)}`);
      return genderFromPatientPayload(payload);
    } catch { return ""; }
  }

  function currentPatientGender() {
    try {
      const fields = Ext.ComponentQuery.query("textfield, displayfield, field");
      for (const field of fields) {
        const label = trLower(`${field.fieldLabel || ""} ${field.name || ""} ${field.itemId || ""}`);
        const value = clean(field.getValue?.() ?? field.value ?? field.getRawValue?.() ?? "");
        if ((label.includes("cinsiyet") || isFemaleGender(value) || isMaleGender(value)) && (isFemaleGender(value) || isMaleGender(value))) return value;
      }
    } catch {}
    return "";
  }

  function currentOperationNo() {
    try {
      const fields = Ext.ComponentQuery.query("textfield, displayfield, field");
      for (const field of fields) {
        const label = trLower(`${field.fieldLabel || ""} ${field.name || ""} ${field.itemId || ""}`);
        if (!label.includes("işlem no") && !label.includes("islem no") && !label.includes("islemno")) continue;
        return clean(field.getValue?.() ?? field.value ?? field.getRawValue?.() ?? "");
      }
    } catch {}
    return "";
  }

  async function genderFromSelectedRow(operation) {
    const grid = operation.sourceGrid;
    const record = operation.sourceRecord;
    if (!grid || !record) return "";
    try {
      grid.getSelectionModel?.().select(record, false, false);
      grid.getView?.().focusRow?.(record);
    } catch { return ""; }
    for (let attempt = 0; attempt < 24; attempt++) {
      await sleep(50);
      const value = currentPatientGender();
      const shownOperationNo = currentOperationNo();
      const correctRow = !operation.operationNo || !shownOperationNo || shownOperationNo.includes(operation.operationNo);
      if (correctRow && (isFemaleGender(value) || isMaleGender(value))) return value;
    }
    return "";
  }

  async function resolveOperationsByGender(targetGender) {
    const all = [...state.operations];
    const matches = value => targetGender === "female" ? isFemaleGender(value) : isMaleGender(value);
    const matched = all.filter(x => matches(x.gender));
    let unknown = all.filter(x => !isFemaleGender(x.gender) && !isMaleGender(x.gender));
    state.genderProcessed = all.length - unknown.length;

    const serviceQueue = unknown.filter(x => x.birimSevkId);
    await Promise.all(Array.from({ length: 8 }, async () => {
      while (serviceQueue.length && !state.stopped) {
        const operation = serviceQueue.shift();
        if (!operation) return;
        operation.gender = await genderFromService(operation);
        state.genderProcessed++;
        if (matches(operation.gender)) matched.push(operation);
        render();
      }
    }));

    unknown = unknown.filter(x => !isFemaleGender(x.gender) && !isMaleGender(x.gender));
    for (const operation of unknown) {
      if (state.stopped) break;
      operation.gender = await genderFromSelectedRow(operation);
      state.genderProcessed++;
      if (matches(operation.gender)) matched.push(operation);
      render();
    }
    state.operations = matched;
    return matched;
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

  function contactFromObject(data) {
    return {
      identityNo: clean(findValue(data, [
        "kimlikNo", "KIMLIK_NO", "tcKimlikNo", "TC_KIMLIK_NO", "tckn",
        "hasta.kimlik.kimlikNo", "hasta.kimlik.tcKimlikNo", "hastaGelis.hasta.kimlik.kimlikNo"
      ])),
      phone: clean(findValue(data, [
        "telefon", "TELEFON", "telefonNo", "TELEFON_NO", "cepTelefonu", "CEP_TELEFONU", "cepTelefon",
        "gsm", "GSM", "mobilTelefon", "hasta.telefon", "hasta.cepTelefonu", "hasta.kimlik.telefon",
        "hasta.kimlik.cepTelefonu", "hastaGelis.hasta.telefon", "hastaGelis.hasta.cepTelefonu"
      ]))
    };
  }

  function isGynecology(consult) {
    return /kadın hastalıkları ve doğum|kadın doğum|jinekoloji|jinekolojik|obstetri|perinatoloji|perinatology/i.test(`${consult.consultationUnit} ${consult.requestReason} ${consult.answer}`);
  }

  async function scanOperation(operation) {
    if (!operation.hastaGelisId) throw new Error("hastaGelisId bulunamadı");
    const payload = await apiJson(`/Poliklinik/Poliklinik/getHastaGelisKonsultasyonList/${encodeURIComponent(operation.hastaGelisId)}/1`);
    const rows = Array.isArray(payload?.data) ? payload.data : [];
    const gynecologyRows = rows.filter(raw => isGynecology(consultationFromRaw(raw)));
    if (!gynecologyRows.length) return [];

    let identityNo = operation.identityNo;
    let phone = operation.phone;
    for (const raw of gynecologyRows) {
      const contact = contactFromObject(raw);
      identityNo ||= contact.identityNo;
      phone ||= contact.phone;
    }

    const fallbackSevkId = operation.birimSevkId || clean(gynecologyRows[0]?.birimSevk?.id || "");
    if ((!identityNo || !phone) && fallbackSevkId) {
      try {
        const patientPayload = await apiJson(`/Tibbi/HastaBirimSevk/getSevkUyariInfo/${encodeURIComponent(fallbackSevkId)}`);
        const contact = contactFromObject(patientPayload);
        identityNo ||= contact.identityNo;
        phone ||= contact.phone;
      } catch {}
    }

    return gynecologyRows.map(consultationFromRaw).map(consult => ({ ...operation, identityNo, phone, ...consult }));
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
    state.genderFilter = document.getElementById("fkd-gender")?.value || "all";
    if (state.genderFilter !== "all" && !state.genderResolved) {
      state.running = true;
      state.stopped = false;
      const genderLabel = state.genderFilter === "female" ? "kadın" : "erkek";
      setMessage(`Cinsiyetler HBYS hasta bilgi alanından okunuyor; ${genderLabel} hastalar ayrılıyor…`);
      render();
      await resolveOperationsByGender(state.genderFilter);
      state.running = false;
      state.genderResolved = true;
      if (!state.operations.length) {
        setMessage(`Seçilen ${genderLabel} cinsiyetinde ameliyat kaydı bulunamadı.`, true);
        render();
        return;
      }
    } else {
      state.genderResolved = true;
    }
    const selectionLabel = state.genderFilter === "female" ? "kadın" : state.genderFilter === "male" ? "erkek" : "kadın ve erkek tüm";
    setMessage(`${state.operations.length} ${selectionLabel} ameliyat kaydının konsültasyonları taranıyor; kadın doğum, jinekoloji, obstetri ve perinatoloji sonuçlara alınacak…`);
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

  function hasPregnancyMention(item) {
    return trLower(`${item?.requestReason || ""} ${item?.answer || ""}`).includes("gebe");
  }

  function operationDateNumber(value) {
    const match = clean(value).match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})(?:\D+(\d{1,2}):(\d{2}))?/);
    if (!match) return 0;
    return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]), Number(match[4] || 0), Number(match[5] || 0)).getTime();
  }

  function sortedResults() {
    return uniqueResults().sort((a, b) =>
      Number(hasPregnancyMention(b)) - Number(hasPregnancyMention(a)) ||
      operationDateNumber(b.operationDate) - operationDateNumber(a.operationDate)
    );
  }

  function csvEscape(value) {
    return `"${String(value ?? "").replace(/"/g, '""')}"`;
  }

  function downloadCsv() {
    const headers = [
      "Ameliyat No", "Ameliyat Tarihi", "Ameliyat", "Hasta", "TC Kimlik No", "Telefon", "Cinsiyet", "Doğum Tarihi",
      "Hasta Geliş ID", "Konsültasyon Tarihi", "Kadın Doğum Birimi", "İsteyen Birim", "İsteyen Doktor",
      "Gebe İfadesi", "İstem Nedeni", "Konsültasyon Yanıtı", "Durum"
    ];
    const rows = sortedResults().map(x => [
      x.operationNo, x.operationDate, x.operationName, x.patientName, x.identityNo, x.phone, x.gender, x.birthDate,
      x.hastaGelisId, x.consultationDate, x.consultationUnit, x.requestingUnit, x.requestingDoctor,
      hasPregnancyMention(x) ? "EVET" : "HAYIR", x.requestReason, x.answer, x.status
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
    const rows = sortedResults();
    tbody.innerHTML = rows.map(x => `
      <tr class="${hasPregnancyMention(x) ? "fkd-pregnant" : ""}">
        <td>${hasPregnancyMention(x) ? '<span class="fkd-badge">GEBE</span>' : "-"}</td>
        <td>${escapeHtml(x.patientName || "-")}</td>
        <td>${escapeHtml(x.identityNo || "-")}</td>
        <td>${escapeHtml(x.phone || "-")}</td>
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
    const results = uniqueResults();
    const pregnancyCount = results.filter(hasPregnancyMention).length;
    if (stats) stats.textContent = `Toplam ameliyat: ${state.sourceOperationCount} | İşlenen: ${state.processed} | Kadın doğum konsültasyonu: ${results.length} | “Gebe” geçen: ${pregnancyCount} | Hata: ${state.errors.length}`;
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
        #fonet-kd-app select{border:1px solid #94a3b8;border-radius:7px;padding:8px 28px 8px 9px;margin:0 10px 0 5px;background:#fff}
        #fonet-kd-app button.alt{background:#475569} #fonet-kd-app button.stop{background:#b91c1c} #fonet-kd-app button:disabled{opacity:.4;cursor:not-allowed}
        #fonet-kd-app .progress{height:9px;background:#dbe5ec;border-radius:6px;overflow:hidden} #fkd-progress-bar{height:100%;width:0;background:#16a34a;transition:width .2s}
        #fonet-kd-app .table-wrap{overflow:auto;flex:1;background:#fff;border:1px solid #cbd5e1;border-radius:8px}
        #fonet-kd-app table{border-collapse:collapse;width:100%;font-size:12px} #fonet-kd-app th{position:sticky;top:0;background:#e2e8f0;text-align:left;padding:8px;border-bottom:1px solid #94a3b8}
        #fonet-kd-app td{padding:7px;border-bottom:1px solid #e2e8f0;vertical-align:top;max-width:260px} #fkd-message{min-height:18px}
        #fonet-kd-app tr.fkd-pregnant td{background:#fff7d6} #fonet-kd-app .fkd-badge{display:inline-block;background:#d97706;color:#fff;border-radius:999px;padding:3px 7px;font-weight:800}
      </style>
      <header><h1>FONET Kadın Doğum Konsültasyon Tarayıcı <small>v${VERSION}</small></h1><button id="fkd-close" class="alt">Kapat</button></header>
      <div class="body">
        <div>
          <label for="fkd-gender"><strong>Cinsiyet:</strong></label><select id="fkd-gender"><option value="all">Tümü</option><option value="female">Kadın</option><option value="male">Erkek</option></select>
          <button id="fkd-find">Listeyi Bul</button><button id="fkd-start" disabled>Taramayı Başlat</button>
          <button id="fkd-pause" class="alt" disabled>Duraklat</button><button id="fkd-stop" class="stop" disabled>Durdur</button>
          <button id="fkd-csv" class="alt">CSV İndir</button>
        </div>
        <div id="fkd-message">Önce HBYS ameliyat sorgusunu çalıştırın, sonra Listeyi Bul'a basın.</div>
        <strong id="fkd-stats">Toplam ameliyat: 0 | İşlenen: 0 | Kadın doğum konsültasyonu: 0 | “Gebe” geçen: 0 | Hata: 0</strong>
        <div class="progress"><div id="fkd-progress-bar"></div></div>
        <div class="table-wrap"><table><thead><tr><th>Grup</th><th>Hasta</th><th>TC Kimlik No</th><th>Telefon</th><th>Ameliyat</th><th>Konsültasyon</th><th>Birim</th><th>İstem nedeni</th><th>Yanıt</th></tr></thead><tbody id="fkd-results"></tbody></table></div>
      </div>`;
    document.documentElement.appendChild(panel);
    state.panel = panel;

    document.getElementById("fkd-find").onclick = () => {
      try {
        const found = collectOperations();
        setMessage(`${found.sourceCount} ameliyat kaydı bulundu. Cinsiyet seçimini yapıp taramayı başlatın; yalnızca kadın doğum konsültasyonları listelenecek (tablo: ${found.gridId}).`);
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
