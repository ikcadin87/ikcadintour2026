// ============================================================
// LINKTREE ACARA - Google Apps Script (Code.gs)
// ============================================================

const SPREADSHEET_ID     = "1SM3eXzQ0NyKKofkvVntshS-Y0wFvK-PbWpn5o6PiQRQ";
const BUS_SPREADSHEET_ID = "1Qp7ayGYnfbI604AVTXpqnCw8qDMWAF5ihtfDkLaTBWc";

var CACHE_KEY = "ikcadin_all_data_v1";
var CACHE_TTL  = 300; // detik (5 menit)

var CLAIM_SHEET     = "Claim";
var ACTIVE_MEAL_KEY = "ACTIVE_MEAL";

function doGet(e) {
  var param        = (e && e.parameter) ? e.parameter : {};
  var forceRefresh = (param.refresh === "1");
  var wantsData    = (param.type === "data" || forceRefresh);
  var wantsClaim   = (param.type === "claimList");

  // ── Mode API: claim list (nama + dept saja, inisial tidak di-expose) ──
  if (wantsClaim) {
    var cl  = getClaimList();
    var am  = PropertiesService.getScriptProperties().getProperty(ACTIVE_MEAL_KEY) || "none";
    cl.active_meal = am;
    return ContentService.createTextOutput(JSON.stringify(cl)).setMimeType(ContentService.MimeType.JSON);
  }

  // ── Mode API: daftar PIC kunci kamar ──
  if (param.type === "keyList") {
    return ContentService.createTextOutput(JSON.stringify(getKeyList())).setMimeType(ContentService.MimeType.JSON);
  }

  // ── Mode API: kembalikan JSON ──────────────────────────────────
  if (wantsData) {
    var cache = CacheService.getScriptCache();
    if (!forceRefresh) {
      var cached = cache.get(CACHE_KEY);
      if (cached) {
        return ContentService
          .createTextOutput(cached)
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    var json = JSON.stringify(getAllData());
    try { cache.put(CACHE_KEY, json, CACHE_TTL); } catch(e2) {}
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── Mode Web App: sajikan halaman HTML ─────────────────────────
  var page = param.page || "index";
  var file = (page === "admin") ? "admin" : "index";
  return HtmlService
    .createHtmlOutputFromFile(file)
    .setTitle(file === "admin" ? "Admin Panel — IKCADIN TOUR 2026" : "IKCADIN TOUR 2026")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=5.0");
}

function getAllData() {
  function safe(fn) {
    try { return fn(); } catch(e) { return { data: [], _error: e.message }; }
  }
  return {
    config:  safe(() => getConfig()),
    rules:   safe(() => getSheetData("Rules")),
    rundown: safe(() => getSheetData("Rundown")),
    kamar:   safe(() => getKamarData()),
    games:   safe(() => getGamesDataHorizontal()),
    bus:     safe(() => getBusDataVertical()),
    kontak:  safe(() => getSheetData("Kontak")),
    custom:  safe(() => getCustomSections()),
    mobil:   safe(() => getSheetData("Mobil Pribadi")),
    denah:   safe(() => getDenahKamar()),
    panitia: safe(() => getPanitiaData()),
    beranda:      safe(() => getSheetData("Hadiah")),
    pengumuman:   safe(() => getSheetData("Pengumuman")),
    _ok: true,
    _ts: new Date().toISOString()
  };
}

// Tes apakah return value bisa di-serialize (jalankan dari editor)
function debugSerialize() {
  try {
    const result = getAllData();
    const json = JSON.stringify(result);
    Logger.log("Serialize OK — ukuran data: " + json.length + " bytes");
    // Cek tiap key untuk temukan nilai bermasalah
    for (const key of Object.keys(result)) {
      try { JSON.stringify(result[key]); }
      catch(e) { Logger.log("GAGAL serialize key: " + key + " → " + e.message); }
    }
    return { ok: true, size: json.length };
  } catch(e) {
    Logger.log("GAGAL serialize: " + e.message);
    return { ok: false, error: e.message };
  }
}

// Fungsi debug — jalankan manual dari editor untuk cek tiap bagian
function debugGetAllData() {
  const result = getAllData();
  const log = {};
  for (const key of Object.keys(result)) {
    const val = result[key];
    log[key] = val?._error ? "ERROR: " + val._error : "OK";
  }
  Logger.log(JSON.stringify(log, null, 2));
  return log;
}

function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { data: [] };
    const raw = sheet.getDataRange().getValues();
    if (raw.length < 2) return { data: [] };
    const headers = raw[0].map(h => String(h).trim());
    const data = raw.slice(1).filter(r => r.some(c => c !== "")).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ""; });
      return obj;
    });
    return { data };
  } catch (e) { return { data: [] }; }
}

function getConfig() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Config");
    if (!sheet) return { data: {} };
    const data = sheet.getDataRange().getValues();
    const config = {};
    data.forEach(row => {
      if (!row[0]) return;
      const key = row[0].toString().toLowerCase().replace(/\s/g, "_");
      const val = row[1];
      // Konversi Date object ke string agar bisa di-serialize via google.script.run
      config[key] = (val instanceof Date) ? Utilities.formatDate(val, 'Asia/Jakarta', "yyyy-MM-dd") : String(val == null ? "" : val);
    });
    return { data: config };
  } catch (e) { return { data: {} }; }
}

// Tambahkan fungsi-fungsi lainnya (getBusDataVertical, getKamarData, dll) di bawah sini
// Pastikan di awal setiap fungsi ada baris: const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

function getGamesDataHorizontal() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Kelompok Games");
    if (!sheet) return { data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return { data: [] };

    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(10, data.length); i++) {
       if (data[i].filter(cell => String(cell).toUpperCase().includes("KELOMPOK")).length > 1) { headerRowIdx = i; break; }
    }
    if (headerRowIdx === -1) return { data: [] };

    const headers = data[headerRowIdx];
    const groups = [];

    for (let col = 0; col < headers.length; col++) {
      const groupName = String(headers[col]).trim();
      if (groupName.toUpperCase().includes("KELOMPOK")) {
        const members = [];
        let stats = { single: 0, double: 0, keluarga: 0, total_main: 0, total_panitia: 0 };
        for (let row = headerRowIdx + 1; row < data.length; row++) {
          let namaRaw = String(data[row][col] || "").trim();
          if (namaRaw !== "") {
            let nama = namaRaw; let dept = String(data[row][col + 1] || "").trim();
            let status = String(data[row][col + 2] || "").trim(); let keterangan = String(data[row][col + 3] || "").trim();
            let isPanitia = (keterangan.toLowerCase().includes("panitia") || nama.toLowerCase().includes("panitia") || dept.toLowerCase().includes("panitia"));
            if (isPanitia) { nama = nama.replace(/[-_]?\s*\(?panitia\)?/gi, "").trim(); stats.total_panitia++; } else { stats.total_main++; }
            members.push({ nama, dept, status, isPanitia });
            const s = status.toLowerCase();
            if (s.includes("single")) stats.single++; else if (s.includes("double") || s.includes("pav")) stats.double++; else if (s.includes("keluarga")) stats.keluarga++;
          }
        }
        groups.push({ nama: groupName, stats: stats, members: members });
      }
    }
    return { format: "horizontal", kelompok: groups };
  } catch(e) { return { data: [] }; }
}

function getKamarData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Kamar");
    if (!sheet) return { data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { data: [] };

    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(20, data.length); i++) {
      if (data[i].some(cell => String(cell).toLowerCase().includes("room number"))) { headerRowIdx = i; break; }
    }
    if (headerRowIdx === -1) return { data: [] };

    const headers = data[headerRowIdx].map(h => String(h).toLowerCase().trim());
    const colType = headers.findIndex(h => h.includes("room type"));
    const colRoomNum = headers.findIndex(h => h.includes("room number"));
    const colKode = colRoomNum + 1; const colPax = headers.findIndex(h => h.includes("pax"));
    const colName = headers.findIndex(h => h.includes("name") || h.includes("nama"));

    const kamarList = []; let currentRoomType = ""; let currentRoomNum = "";  
    for (let i = headerRowIdx + 1; i < data.length; i++) {
      const row = data[i]; const name = String(row[colName] || "").trim();
      if (!name) continue;
      if (colType !== -1 && row[colType]) currentRoomType = String(row[colType]).trim();
      if (row[colRoomNum]) currentRoomNum = String(row[colRoomNum]).trim();
      // Room Type sengaja tidak dikirim ke client (disembunyikan dari API publik)
      kamarList.push({ "Room Number": currentRoomNum, "Kode Room": String(row[colKode] || "").trim(), "PAX": String(row[colPax] || "").trim(), "Name": name });
    }
    return { data: kamarList };
  } catch (e) { return { data: [] }; }
}

function getPanitiaData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Panitia");
    if (!sheet) return { data: {} };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { data: {} };

    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    const groups = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i]; const nama = String(row[headers.indexOf("nama")] || "").trim();
      if (!nama) continue;
      const jabatan = String(row[headers.indexOf("jabatan")] || "Lainnya").trim();
      const role = String(row[headers.indexOf("role")] || "").trim();
      const memberObj = { nama, role, foto: String(row[headers.indexOf("foto_url")] || "").trim(), hp: String(row[headers.indexOf("no_hp")] || "").trim(), keterangan: String(row[headers.indexOf("keterangan")] || "").trim(), isKoordinator: role.toLowerCase().includes("koordinator") || role.toLowerCase().includes("ketua") };
      if (!groups[jabatan]) groups[jabatan] = { nama_jabatan: jabatan, koordinator: [], anggota: [] };
      if (memberObj.isKoordinator) groups[jabatan].koordinator.push(memberObj); else groups[jabatan].anggota.push(memberObj);
    }
    return { data: Object.values(groups) };
  } catch (e) { return { data: {} }; }
}

function getDenahKamar() {
  try {
    // Bangun peta nomor kamar → daftar nama dari sheet Kamar
    const kamarResult = getKamarData();
    const rooms = {};
    let currentRoom = '';
    (kamarResult.data || []).forEach(function(r) {
      const n = String(r['Room Number'] || '').trim();
      if (n) currentRoom = n;
      const name = String(r['Name'] || '').trim();
      if (currentRoom && name) {
        if (!rooms[currentRoom]) rooms[currentRoom] = [];
        rooms[currentRoom].push(name);
      }
    });
    return { rooms: rooms };
  } catch(e) {
    return { rooms: {}, _error: e.message };
  }
}

function getCustomSections() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Custom Sections");
    if (!sheet) return { data: [] };
    const raw = sheet.getDataRange().getValues();
    if (raw.length < 2) return { data: [] };
    const rows = raw.slice(1).filter(row => row.some(cell => cell !== ""));
    const sections = {};
    rows.forEach(row => {
      const section_name = row[0]?.toString() || ""; const emoji = row[1]?.toString() || "📌";
      const judul = row[2]?.toString() || ""; const konten = row[3]?.toString() || "";
      if (!section_name) return;
      if (!sections[section_name]) sections[section_name] = { emoji, nama: section_name, items: [] };
      if (judul || konten) sections[section_name].items.push({ judul, konten });
    });
    return { data: Object.values(sections) };
  } catch (e) { return { data: [] }; }
}

// ──────────────────────────────────────────────────────────────────────────
// ADMIN API — doPost() dengan verifikasi password via Script Properties
// Password default: ikcadin2026 — SEGERA ganti di GAS > Project Settings > Script Properties
// Tambahkan key: ADMIN_PASSWORD, value: password pilihan kamu
// ──────────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;

    // Ping — hanya verifikasi token
    if (action === 'ping') {
      var valid = verifyToken(payload.token);
      return jsonOut({ ok: valid, error: valid ? null : 'Unauthorized' });
    }

    // ── User-facing (tidak butuh token) ──────────────────────────
    if (action === 'claim_verify') {
      return jsonOut(verifyClaimUser(payload.nama, payload.inisial));
    }
    if (action === 'claim_get_active') {
      var am = PropertiesService.getScriptProperties().getProperty(ACTIVE_MEAL_KEY) || 'none';
      return jsonOut({ ok: true, active_meal: am });
    }

    // ── Admin actions (butuh token) ───────────────────────────────
    if (!verifyToken(payload.token)) {
      return jsonOut({ ok: false, error: 'Unauthorized' });
    }

    var result;
    switch (action) {
      case 'get_admin_data':    result = getAdminSheetData(payload.sheet);                               break;
      case 'create':            result = adminCreateRow(payload.sheet, payload.data);                    break;
      case 'update':            result = adminUpdateRow(payload.sheet, payload.rowIndex, payload.data);  break;
      case 'delete':            result = adminDeleteRow(payload.sheet, payload.rowIndex);                break;
      case 'claim_scan':        result = scanClaimBarcode(payload.nama, payload.inisial, payload.meal_type); break;
      case 'claim_confirm':     result = confirmClaim(payload.nama, payload.inisial, payload.meal_type);     break;
      case 'claim_set_active':
        PropertiesService.getScriptProperties().setProperty(ACTIVE_MEAL_KEY, payload.meal_type || 'none');
        result = { ok: true, active_meal: payload.meal_type || 'none' };
        break;
      default: result = { ok: false, error: 'Unknown action: ' + action };
    }

    if (['create','update','delete','claim_confirm'].indexOf(action) !== -1) {
      try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e2) {}
    }

    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function verifyToken(token) {
  var storedPw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!storedPw) storedPw = 'ikcadin2026'; // default — ganti via Script Properties!
  return token === storedPw;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAdminSheetData(sheetName) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: 'Sheet tidak ditemukan: ' + sheetName };
    var raw = sheet.getDataRange().getValues();
    if (raw.length < 1) return { ok: true, headers: [], rows: [] };
    var headers = raw[0].map(function(h) { return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < raw.length; i++) {
      var row = raw[i];
      if (row.every(function(c) { return c === ''; })) continue;
      rows.push({
        rowIndex: i + 1, // 1-based (baris 1 = header)
        data: row.map(function(c) {
          if (c instanceof Date) return Utilities.formatDate(c, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm');
          return (c !== undefined && c !== null) ? String(c) : '';
        })
      });
    }
    return { ok: true, headers: headers, rows: rows };
  } catch (err) { return { ok: false, error: err.message }; }
}

function adminCreateRow(sheetName, data) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: 'Sheet tidak ditemukan' };
    sheet.appendRow(data);
    return { ok: true };
  } catch (err) { return { ok: false, error: err.message }; }
}

function adminUpdateRow(sheetName, rowIndex, data) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: 'Sheet tidak ditemukan' };
    sheet.getRange(rowIndex, 1, 1, data.length).setValues([data]);
    return { ok: true };
  } catch (err) { return { ok: false, error: err.message }; }
}

function adminDeleteRow(sheetName, rowIndex) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: 'Sheet tidak ditemukan' };
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) { return { ok: false, error: err.message }; }
}

// ══════════════════════════════════════════════════════════════════════════
// CLAIM SYSTEM
// Sheet "Claim" di spreadsheet utama:
//   A=Nama  B=Inisial(3 huruf)  C=Department  D=Pax
//   E=Snack_Claim  F=Makan_Siang_Claim  G=Makan_Malam_Claim
// ══════════════════════════════════════════════════════════════════════════

// Kolom claim (1-based untuk getRange)
var CLAIM_COL = { snack: 5, makan_siang: 6, makan_malam: 7 };

function getClaimList() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: false, data: [], error: 'Sheet Claim tidak ada' };
    var raw = sheet.getDataRange().getValues();
    if (raw.length < 2) return { ok: true, data: [] };
    var list = [];
    for (var i = 1; i < raw.length; i++) {
      var nama = String(raw[i][0] || '').trim();
      if (!nama) continue;
      list.push({ nama: nama, dept: String(raw[i][2] || '').trim() });
    }
    return { ok: true, data: list };
  } catch(err) { return { ok: false, data: [], error: err.message }; }
}

function verifyClaimUser(nama, inisial) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: false, error: 'Sheet Claim tidak ditemukan' };
    var raw = sheet.getDataRange().getValues();
    var normNama = (nama || '').trim().toUpperCase();
    var normIni  = (inisial || '').trim().toUpperCase();
    for (var i = 1; i < raw.length; i++) {
      var rNama = String(raw[i][0] || '').trim().toUpperCase();
      var rIni  = String(raw[i][1] || '').trim().toUpperCase();
      if (rNama === normNama && rIni === normIni) {
        var tz = 'Asia/Jakarta';
        var am = PropertiesService.getScriptProperties().getProperty(ACTIVE_MEAL_KEY) || 'none';
        return {
          ok: true,
          nama:              String(raw[i][0]).trim(),
          dept:              String(raw[i][2] || '').trim(),
          pax:               parseInt(raw[i][3]) || 1,
          snack_claim:       String(raw[i][4] || '').trim(),
          makan_siang_claim: String(raw[i][5] || '').trim(),
          makan_malam_claim: String(raw[i][6] || '').trim(),
          active_meal:       am
        };
      }
    }
    return { ok: false, error: 'Nama atau inisial tidak cocok' };
  } catch(err) { return { ok: false, error: err.message }; }
}

function scanClaimBarcode(nama, inisial, mealType) {
  // Admin scan: cek status tanpa konfirmasi
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: false, error: 'Sheet Claim tidak ditemukan' };
    var col = CLAIM_COL[mealType];
    if (!col) return { ok: false, error: 'Tipe meal tidak valid: ' + mealType };
    var raw = sheet.getDataRange().getValues();
    var normNama = (nama || '').trim().toUpperCase();
    var normIni  = (inisial || '').trim().toUpperCase();
    for (var i = 1; i < raw.length; i++) {
      var rNama = String(raw[i][0] || '').trim().toUpperCase();
      var rIni  = String(raw[i][1] || '').trim().toUpperCase();
      if (rNama === normNama && rIni === normIni) {
        var existing = String(raw[i][col - 1] || '').trim();
        return {
          ok:               true,
          nama:             String(raw[i][0]).trim(),
          dept:             String(raw[i][2] || '').trim(),
          pax:              parseInt(raw[i][3]) || 1,
          already_claimed:  !!existing,
          claim_timestamp:  existing
        };
      }
    }
    return { ok: false, error: 'QR tidak valid atau peserta tidak terdaftar' };
  } catch(err) { return { ok: false, error: err.message }; }
}

function confirmClaim(nama, inisial, mealType) {
  // Admin konfirmasi: tulis timestamp ke sheet
  // LockService → cegah double-claim saat banyak panitia scan bersamaan
  var lock = LockService.getScriptLock();
  try {
    // Tunggu giliran maksimal 10 detik bila ada panitia lain sedang menulis
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, error: 'Sistem sibuk, coba lagi sebentar.' };
  }
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: false, error: 'Sheet Claim tidak ditemukan' };
    var col = CLAIM_COL[mealType];
    if (!col) return { ok: false, error: 'Tipe meal tidak valid: ' + mealType };
    var raw = sheet.getDataRange().getValues();
    var normNama = (nama || '').trim().toUpperCase();
    var normIni  = (inisial || '').trim().toUpperCase();
    var tz = 'Asia/Jakarta';
    var ts = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
    for (var i = 1; i < raw.length; i++) {
      var rNama = String(raw[i][0] || '').trim().toUpperCase();
      var rIni  = String(raw[i][1] || '').trim().toUpperCase();
      if (rNama === normNama && rIni === normIni) {
        var existing = String(raw[i][col - 1] || '').trim();
        if (existing) {
          return { ok: false, already: true, timestamp: existing,
                   nama: String(raw[i][0]).trim(), dept: String(raw[i][2]||'').trim(), pax: parseInt(raw[i][3])||1 };
        }
        sheet.getRange(i + 1, col).setNumberFormat('@').setValue(ts);
        SpreadsheetApp.flush(); // pastikan tulisan tersimpan sebelum lock dilepas
        return { ok: true, timestamp: ts,
                 nama: String(raw[i][0]).trim(), dept: String(raw[i][2]||'').trim(), pax: parseInt(raw[i][3])||1 };
      }
    }
    return { ok: false, error: 'Peserta tidak ditemukan' };
  } catch(err) {
    return { ok: false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── DEBUG: jalankan dari GAS Editor untuk lihat struktur spreadsheet bus ──
function debugBusSheet() {
  const ss    = SpreadsheetApp.openById(BUS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("List Bus");
  if (!sheet) { Logger.log("Sheet 'List Bus' TIDAK DITEMUKAN"); return; }
  const data  = sheet.getDataRange().getValues();
  Logger.log("Total baris: " + data.length);
  Logger.log("=== 20 BARIS PERTAMA (kolom A–K) ===");
  for (let r = 0; r < Math.min(20, data.length); r++) {
    const cols = [];
    for (let c = 0; c < 11; c++) {
      cols.push(String(c === 0 ? "A" : String.fromCharCode(65 + c)) + "=" + JSON.stringify(String(data[r][c] || "")));
    }
    Logger.log("Baris " + (r + 1) + ": " + cols.join(" | "));
  }
}

// FUNGSI KHUSUS LAINNYA (Setiap fungsi membuka koneksinya sendiri)
function getBusDataVertical() {
  try {
    const ss = SpreadsheetApp.openById(BUS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("List Bus");
    if (!sheet) return { data: [] };
    const data = sheet.getDataRange().getValues();
    const result = [];
    let currentBus = null;

    for (let r = 0; r < data.length; r++) {
      let val        = String(data[r][1] || "").trim();
      let dept       = String(data[r][2] || "").trim();
      let jmlDewasa  = parseInt(data[r][6]) || 0;
      let jmlAnak    = parseInt(data[r][7]) || 0;
      let kolI       = String(data[r][8] || "").trim().toUpperCase(); // kolom I

      if (val.toUpperCase().startsWith("BUS")) {
        // Baris header bus — PIC belum diketahui, akan diisi saat scan anggota
        currentBus = { nama: val, Anggota: [], Kapasitas: 0, pic: "" };
        result.push(currentBus);
      } else if (currentBus && val !== "" && kolI !== "BUS/PRIBADI" && !val.toUpperCase().includes("DEPARTMENT") && !val.toUpperCase().includes("NAMA")) {
        // Jika kolom I = "PIC BUS", tandai sebagai PIC bus ini
        if (kolI === "PIC BUS") {
          currentBus.pic = val;
        }
        // Tetap masukkan ke daftar anggota
        currentBus.Anggota.push({ nama: val, department: dept, dewasa: jmlDewasa, anak: jmlAnak });
        currentBus.Kapasitas += (jmlDewasa + jmlAnak);
      }
    }
    return { data: result };
  } catch (e) { return { data: [] }; }
}

// ============================================================
// doPost — handles all POST requests
// ============================================================
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;

    // ── Ping: cek token saja ──────────────────────────────────
    if (action === 'ping') {
      var valid = verifyToken(payload.token);
      return jsonOut({ ok: valid, error: valid ? null : 'Unauthorized' });
    }

    // ── User-facing: tidak perlu token ────────────────────────
    if (action === 'claim_verify') {
      return jsonOut(verifyClaimUser(payload.nama, payload.inisial));
    }
    if (action === 'key_verify') {
      return jsonOut(verifyKeyUser(payload.nama, payload.inisial));
    }
    if (action === 'claim_get_active') {
      var am = PropertiesService.getScriptProperties().getProperty(ACTIVE_MEAL_KEY) || 'none';
      return jsonOut({ ok: true, active_meal: am });
    }

    // ── Admin-only: perlu token ───────────────────────────────
    if (!verifyToken(payload.token)) {
      return jsonOut({ ok: false, error: 'Unauthorized' });
    }

    var result;
    switch (action) {
      case 'claim_scan':
        result = scanClaimBarcode(payload.nama, payload.inisial, payload.meal_type);
        break;
      case 'claim_confirm':
        result = confirmClaim(payload.nama, payload.inisial, payload.meal_type);
        break;
      case 'key_scan':
        result = verifyKeyUser(payload.nama, payload.inisial);
        break;
      case 'key_lookup':
        result = lookupKeyByName(payload.nama);
        break;
      case 'key_confirm':
        result = confirmKey(payload.nama, payload.inisial);
        break;
      case 'key_rekap':
        result = getKeyRekap();
        break;
      case 'claim_set_active':
        PropertiesService.getScriptProperties().setProperty(ACTIVE_MEAL_KEY, payload.meal_type || 'none');
        result = { ok: true, active_meal: payload.meal_type || 'none' };
        break;
      case 'updateRow': {
        var ss2    = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet2 = ss2.getSheetByName(payload.sheet);
        if (!sheet2) { result = { ok: false, error: 'Sheet tidak ditemukan' }; break; }
        var rows2  = sheet2.getDataRange().getValues();
        var idx    = -1;
        for (var i = 1; i < rows2.length; i++) {
          if (String(rows2[i][0]).trim() === String(payload.key).trim()) { idx = i + 1; break; }
        }
        if (idx === -1) { result = { ok: false, error: 'Baris tidak ditemukan' }; break; }
        var headers2 = rows2[0].map(function(h){ return String(h).trim(); });
        Object.keys(payload.data).forEach(function(col) {
          var ci = headers2.indexOf(col);
          if (ci >= 0) sheet2.getRange(idx, ci + 1).setValue(payload.data[col]);
        });
        try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(_){}
        result = { ok: true };
        break;
      }
      case 'appendRow': {
        var ss3    = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet3 = ss3.getSheetByName(payload.sheet);
        if (!sheet3) { result = { ok: false, error: 'Sheet tidak ditemukan' }; break; }
        var headers3 = sheet3.getRange(1, 1, 1, sheet3.getLastColumn()).getValues()[0]
                        .map(function(h){ return String(h).trim(); });
        var rowData  = headers3.map(function(h){ return payload.data[h] !== undefined ? payload.data[h] : ''; });
        sheet3.appendRow(rowData);
        try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(_){}
        result = { ok: true };
        break;
      }
      case 'deleteRow': {
        var ss4    = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet4 = ss4.getSheetByName(payload.sheet);
        if (!sheet4) { result = { ok: false, error: 'Sheet tidak ditemukan' }; break; }
        var rows4  = sheet4.getDataRange().getValues();
        var idx4   = -1;
        for (var j = 1; j < rows4.length; j++) {
          if (String(rows4[j][0]).trim() === String(payload.key).trim()) { idx4 = j + 1; break; }
        }
        if (idx4 === -1) { result = { ok: false, error: 'Baris tidak ditemukan' }; break; }
        sheet4.deleteRow(idx4);
        try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(_){}
        result = { ok: true };
        break;
      }
      // ── Admin sheet CRUD (dipakai admin.html) ─────────────────
      case 'get_admin_data': {
        var ssG    = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetG = ssG.getSheetByName(payload.sheet);
        if (!sheetG) { result = { ok: false, error: 'Sheet tidak ditemukan: ' + payload.sheet }; break; }
        var allG   = sheetG.getDataRange().getValues();
        if (allG.length < 1) { result = { ok: true, headers: [], rows: [] }; break; }
        var hdrsG  = allG[0].map(function(h){ return String(h).trim(); });
        var rowsG  = [];
        for (var gi = 1; gi < allG.length; gi++) {
          if (allG[gi].every(function(c){ return c === '' || c === null || c === undefined; })) continue;
          rowsG.push({ rowIndex: gi + 1, data: allG[gi].map(function(c){ return String(c); }) });
        }
        result = { ok: true, headers: hdrsG, rows: rowsG };
        break;
      }
      case 'create': {
        var ssC    = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetC = ssC.getSheetByName(payload.sheet);
        if (!sheetC) { result = { ok: false, error: 'Sheet tidak ditemukan' }; break; }
        sheetC.appendRow(payload.data);
        try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(_){}
        result = { ok: true };
        break;
      }
      case 'update': {
        var ssU    = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetU = ssU.getSheetByName(payload.sheet);
        if (!sheetU) { result = { ok: false, error: 'Sheet tidak ditemukan' }; break; }
        var lastColU = sheetU.getLastColumn();
        var rangeU   = sheetU.getRange(payload.rowIndex, 1, 1, Math.max(lastColU, payload.data.length));
        rangeU.setValues([payload.data.concat(
          new Array(Math.max(0, lastColU - payload.data.length)).fill('')
        ).slice(0, Math.max(lastColU, payload.data.length))]);
        try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(_){}
        result = { ok: true };
        break;
      }
      case 'delete': {
        var ssDel    = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetDel = ssDel.getSheetByName(payload.sheet);
        if (!sheetDel) { result = { ok: false, error: 'Sheet tidak ditemukan' }; break; }
        sheetDel.deleteRow(payload.rowIndex);
        try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(_){}
        result = { ok: true };
        break;
      }
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyToken(token) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!stored) stored = 'ikcadin2026';
  return token === stored;
}

// ── Claim helpers ─────────────────────────────────────────────
function getClaimList() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: true, data: [] };
    var rows  = sheet.getDataRange().getValues();
    if (rows.length < 2) return { ok: true, data: [] };
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var nama = String(rows[i][0] || '').trim();
      if (!nama) continue;
      out.push({
        nama:              nama,
        inisial:           String(rows[i][1] || '').trim(),
        dept:              String(rows[i][2] || '').trim(),
        pax:               rows[i][3] || 1,
        snack_claim:       String(rows[i][4] || '').trim(),
        makan_siang_claim: String(rows[i][5] || '').trim(),
        makan_malam_claim: String(rows[i][6] || '').trim()
      });
    }
    return { ok: true, data: out };
  } catch (err) {
    return { ok: false, error: err.message, data: [] };
  }
}

function verifyClaimUser(nama, inisial) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: false, error: 'Sheet Claim tidak ditemukan' };
    var rows  = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rowNama = String(rows[i][0] || '').trim().toLowerCase();
      var rowIni  = String(rows[i][1] || '').trim().toLowerCase();
      if (rowNama === nama.trim().toLowerCase() && rowIni === inisial.trim().toLowerCase()) {
        var am = PropertiesService.getScriptProperties().getProperty(ACTIVE_MEAL_KEY) || 'none';
        return {
          ok:                true,
          nama:              String(rows[i][0]).trim(),
          inisial:           String(rows[i][1]).trim(),
          dept:              String(rows[i][2] || '').trim(),
          pax:               rows[i][3] || 1,
          snack_claim:       String(rows[i][4] || '').trim(),
          makan_siang_claim: String(rows[i][5] || '').trim(),
          makan_malam_claim: String(rows[i][6] || '').trim(),
          active_meal:       am
        };
      }
    }
    return { ok: false, error: 'Nama atau inisial tidak cocok' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Admin: preview scan (cek status sebelum konfirmasi)
function scanClaimBarcode(nama, inisial, mealType) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: false, error: 'Sheet Claim tidak ditemukan' };
    var rows  = sheet.getDataRange().getValues();
    var colMap = { snack: 4, makan_siang: 5, makan_malam: 6 }; // 0-based
    var col = colMap[mealType];
    if (col === undefined) return { ok: false, error: 'Tipe makan tidak valid' };
    var am = PropertiesService.getScriptProperties().getProperty(ACTIVE_MEAL_KEY) || 'none';
    for (var i = 1; i < rows.length; i++) {
      var rowNama = String(rows[i][0] || '').trim().toLowerCase();
      var rowIni  = String(rows[i][1] || '').trim().toLowerCase();
      if (rowNama === nama.trim().toLowerCase() && rowIni === inisial.trim().toLowerCase()) {
        var alreadyClaimed = String(rows[i][col] || '').trim();
        return {
          ok:           true,
          nama:         String(rows[i][0]).trim(),
          dept:         String(rows[i][2] || '').trim(),
          pax:          rows[i][3] || 1,
          meal_type:    mealType,
          active_meal:  am,
          already_claimed: !!alreadyClaimed,
          claimed_at:   alreadyClaimed
        };
      }
    }
    return { ok: false, error: 'Peserta tidak ditemukan' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Admin: tulis timestamp claim ke sheet
function confirmClaim(nama, inisial, mealType) {
  // LockService → cegah double-claim saat banyak panitia scan bersamaan
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return { ok: false, error: 'Sistem sibuk, coba lagi sebentar.' }; }
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CLAIM_SHEET);
    if (!sheet) return { ok: false, error: 'Sheet Claim tidak ditemukan' };
    var rows  = sheet.getDataRange().getValues();
    var colMap = { snack: 5, makan_siang: 6, makan_malam: 7 }; // 1-based (spreadsheet column)
    var col = colMap[mealType];
    if (!col) return { ok: false, error: 'Tipe makan tidak valid' };
    for (var i = 1; i < rows.length; i++) {
      var rowNama = String(rows[i][0] || '').trim().toLowerCase();
      var rowIni  = String(rows[i][1] || '').trim().toLowerCase();
      if (rowNama === nama.trim().toLowerCase() && rowIni === inisial.trim().toLowerCase()) {
        if (String(rows[i][col - 1] || '').trim()) {
          return { ok: false, error: 'Sudah pernah claim ' + mealType, claimed_at: String(rows[i][col-1]).trim() };
        }
        var ts = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
        sheet.getRange(i + 1, col).setNumberFormat('@').setValue(ts);
        SpreadsheetApp.flush();
        return { ok: true, nama: String(rows[i][0]).trim(), meal_type: mealType, claimed_at: ts };
      }
    }
    return { ok: false, error: 'Peserta tidak ditemukan' };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
// ══════════════════════════════════════════════════════════════════
// CLAIM KUNCI KAMAR  (key handover ke PIC Kamar)
// Kolom sheet "Kamar": Room Type | Room Number | Kode Room | Pax | Name | STATUS | Inisial | Kunci
// ══════════════════════════════════════════════════════════════════
function parseKamarSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Kamar");
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var hr = -1;
  for (var i = 0; i < Math.min(20, data.length); i++) {
    if (data[i].some(function(c){ return String(c).toLowerCase().indexOf("room number") !== -1; })) { hr = i; break; }
  }
  if (hr === -1) return null;
  var H = data[hr].map(function(h){ return String(h).toLowerCase().trim(); });
  function find(kw){ for (var j=0;j<H.length;j++){ if (H[j].indexOf(kw)!==-1) return j; } return -1; }
  var colName = find("name"); if (colName === -1) colName = find("nama");
  var colIni = find("inisial"), colKunci = find("kunci");
  var colStatus = find("status");
  // Fallback: kalau header STATUS tidak ada, cari kolom yang berisi "PIC"
  if (colStatus === -1) {
    for (var c = 0; c < (data[hr] ? data[hr].length : 0); c++) {
      if (c === colName || c === colIni || c === colKunci) continue;
      var found = false;
      for (var rr = hr + 1; rr < data.length; rr++) {
        if (String(data[rr][c] || "").toUpperCase().indexOf("PIC") !== -1) { found = true; break; }
      }
      if (found) { colStatus = c; break; }
    }
  }
  return {
    sheet: sheet, data: data, hr: hr,
    colType: find("room type"), colNum: find("room number"),
    colKode: find("room number") + 1, colPax: find("pax"),
    colName: colName, colStatus: colStatus,
    colIni: colIni, colKunci: colKunci
  };
}

function buildKamarModel(P) {
  var rooms = {}, order = [], curType = "", curNum = "";
  for (var i = P.hr + 1; i < P.data.length; i++) {
    var row = P.data[i];
    if (P.colType !== -1 && row[P.colType] !== "" && row[P.colType] != null) curType = String(row[P.colType]).trim();
    if (row[P.colNum] !== "" && row[P.colNum] != null) curNum = String(row[P.colNum]).trim();
    var name = String(row[P.colName] || "").trim();
    if (!name || !curNum) continue;
    if (!rooms[curNum]) { rooms[curNum] = { num:curNum, type:curType, occupants:[], picName:"", picIni:"", kunci:"", picRow:-1 }; order.push(curNum); }
    rooms[curNum].occupants.push({ kode:String(row[P.colKode]||"").trim(), name:name, pax:String(row[P.colPax]||"").trim() });
    var status = P.colStatus !== -1 ? String(row[P.colStatus]||"").toUpperCase() : "";
    if (status.indexOf("PIC") !== -1) {
      rooms[curNum].picName = name;
      rooms[curNum].picIni  = P.colIni   !== -1 ? String(row[P.colIni]||"").trim()   : "";
      rooms[curNum].kunci   = P.colKunci !== -1 ? String(row[P.colKunci]||"").trim() : "";
      rooms[curNum].picRow  = i; // 0-based; sheet row = i+1
    }
  }
  var pics = {};
  order.forEach(function(num){
    var r = rooms[num];
    if (!r.picName) return;
    var key = r.picName.toLowerCase();
    if (!pics[key]) pics[key] = { nama:r.picName, inisial:r.picIni, rooms:[] };
    pics[key].rooms.push(num);
    if (!pics[key].inisial && r.picIni) pics[key].inisial = r.picIni;
  });
  return { rooms: rooms, order: order, pics: pics };
}

// Daftar PIC untuk dropdown peserta (inisial tidak diexpose)
function getKeyList() {
  try {
    var P = parseKamarSheet();
    if (!P) return { ok: true, data: [] };
    var M = buildKamarModel(P);
    var out = [];
    Object.keys(M.pics).forEach(function(k){
      var p = M.pics[k];
      out.push({ nama: p.nama, rooms: p.rooms });
    });
    out.sort(function(a,b){ return a.nama.localeCompare(b.nama); });
    return { ok: true, data: out };
  } catch (e) {
    return { ok: false, error: String(e), data: [] };
  }
}

// Verifikasi PIC + detail kamar (peserta & preview admin)
function verifyKeyUser(nama, inisial, skipIni) {
  var P = parseKamarSheet();
  if (!P) return { ok:false, error:'Sheet Kamar tidak ditemukan' };
  var M = buildKamarModel(P);
  var p = M.pics[String(nama||'').trim().toLowerCase()];
  if (!p) return { ok:false, error:'Nama PIC tidak ditemukan' };
  if (!skipIni) {
    var ini = String(inisial||'').trim().toUpperCase();
    if (String(p.inisial||'').trim().toUpperCase() !== ini) return { ok:false, error:'Inisial tidak cocok' };
  }
  var rooms = p.rooms.map(function(num){
    var r = M.rooms[num];
    return {
      room: num, type: r.type, claimed_at: r.kunci,
      occupants: r.occupants.map(function(o){ return { kode:o.kode, name:o.name, pax:o.pax }; })
    };
  });
  var allClaimed = rooms.length > 0 && rooms.every(function(r){ return !!r.claimed_at; });
  return { ok:true, nama:p.nama, inisial:p.inisial, rooms:rooms, all_claimed:allClaimed };
}

// Admin: lookup PIC tanpa inisial (panitia sudah terverifikasi via token)
function lookupKeyByName(nama) {
  return verifyKeyUser(nama, '', true);
}

// Admin: tandai kunci diserahkan (semua kamar PIC sekaligus)
// inisial opsional — kalau kosong, lewati pengecekan (admin sudah dipercaya via token)
function confirmKey(nama, inisial) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return { ok:false, error:'Sistem sibuk, coba lagi sebentar.' }; }
  try {
    var P = parseKamarSheet();
    if (!P) return { ok:false, error:'Sheet Kamar tidak ditemukan' };
    if (P.colKunci === -1) return { ok:false, error:'Kolom "Kunci" belum ada di sheet Kamar' };
    var M = buildKamarModel(P);
    var p = M.pics[String(nama||'').trim().toLowerCase()];
    if (!p) return { ok:false, error:'PIC tidak ditemukan' };
    // inisial hanya dicek kalau diisi (alur barcode). Pencarian nama oleh admin: lewati.
    if (String(inisial||'').trim() &&
        String(p.inisial||'').trim().toUpperCase() !== String(inisial||'').trim().toUpperCase())
      return { ok:false, error:'Inisial tidak cocok' };
    var ts = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
    var wrote = [], already = [];
    p.rooms.forEach(function(num){
      var r = M.rooms[num];
      if (r.picRow < 0) return;
      if (String(r.kunci||'').trim()) { already.push(num); return; }
      P.sheet.getRange(r.picRow + 1, P.colKunci + 1).setNumberFormat('@').setValue(ts);
      wrote.push(num);
    });
    SpreadsheetApp.flush();
    return { ok:true, nama:p.nama, claimed_at:ts, rooms:p.rooms, wrote:wrote, already:already,
             already_all: (wrote.length === 0) };
  } catch (err) {
    return { ok:false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Admin: rekap & data semua PIC kunci kamar
function getKeyRekap() {
  try {
    var P = parseKamarSheet();
    if (!P) return { ok:true, data:[], totals:{ pics:0, rooms:0, takenRooms:0, takenPics:0 } };
    var M = buildKamarModel(P);
    var data = [], totalRooms = 0, takenRooms = 0, takenPics = 0;
    Object.keys(M.pics).forEach(function(k){
      var p = M.pics[k];
      var rooms = p.rooms.map(function(num){
        var r = M.rooms[num];
        return { room:num, type:r.type, claimed_at:r.kunci };
      });
      var taken = rooms.filter(function(r){ return !!r.claimed_at; }).length;
      var allc = rooms.length > 0 && taken === rooms.length;
      totalRooms += rooms.length; takenRooms += taken; if (allc) takenPics++;
      data.push({ nama:p.nama, inisial:p.inisial, rooms:rooms, roomCount:rooms.length, takenCount:taken, all_claimed:allc });
    });
    data.sort(function(a,b){ return a.nama.localeCompare(b.nama); });
    return { ok:true, data:data, totals:{ pics:data.length, rooms:totalRooms, takenRooms:takenRooms, takenPics:takenPics } };
  } catch (e) {
    return { ok:false, error:String(e), data:[], totals:{ pics:0, rooms:0, takenRooms:0, takenPics:0 } };
  }
}
