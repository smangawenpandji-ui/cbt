// ID Google Sheets
const SPREADSHEET_ID = "1bdNIooE1uHFcnSuOZnVBu2YmLZua53BpXnxaS3NiUOQ";

// Header sheet yang punya kolom tambahan (tipe soal & koreksi esai)
const SOAL_HEADERS  = ["id", "kode_tes", "pertanyaan", "opsi_a", "opsi_b", "opsi_c", "opsi_d", "opsi_e", "jawaban_benar", "poin", "tipe_soal"];
const SESI_HEADERS = ["username", "nama", "kelompok", "status", "kode_tes", "waktu_login", "last_update", "session_id", "ping_ms"];
const PROGRESS_HEADERS = ["username", "kode_tes", "jawaban_json", "deadline_ms", "pelanggaran", "updated"];
const HASIL_HEADERS = ["id", "username", "nama", "kode_tes", "nilai", "waktu_selesai",
                       "jawaban_json", "skor_otomatis", "poin_total", "poin_esai", "status_koreksi", "skor_esai", "skor_esai_json"];

// Pastikan kolom header lengkap (aman untuk sheet lama yang sudah berisi data)
function ensureHeaders(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < headers.length) {
    sheet.getRange(1, lastCol + 1, 1, headers.length - lastCol)
         .setValues([headers.slice(lastCol)]);
  }
}

let _ss = null;
const _shCache = {};

function getSheet(sheetName) {
  if (_shCache[sheetName]) return _shCache[sheetName];
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = _ss;
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === "mapel") {
      sheet.appendRow(["kode_tes", "nama_mapel", "waktu"]);
    } else if (sheetName === "siswa") {
      sheet.appendRow(["username", "nama", "kelompok", "nik"]);
    } else if (sheetName === "soal") {
      sheet.appendRow(SOAL_HEADERS);
    } else if (sheetName === "admin") {
      sheet.appendRow(["id_proktor", "password", "nama_sekolah"]);
      sheet.appendRow(["U03220006-DK4U", "12345", "SMA NEGERI 1 NGAWEN"]);
    } else if (sheetName === "status_tes") {
      sheet.appendRow(["id", "kode_tes", "token", "status", "kelompok", "tanggal", "jam_mulai", "jam_selesai"]);
    } else if (sheetName === "reset_login") {
      sheet.appendRow(["id", "username", "nama", "waktu_request", "status"]);
    } else if (sheetName === "hasil_tes") {
      sheet.appendRow(HASIL_HEADERS);
    } else if (sheetName === "sesi_siswa") {
      sheet.appendRow(SESI_HEADERS);
    } else if (sheetName === "progress_ujian") {
      sheet.appendRow(PROGRESS_HEADERS);
    }
  }
  if (sheetName === "soal") ensureHeaders(sheet, SOAL_HEADERS);
  if (sheetName === "hasil_tes") ensureHeaders(sheet, HASIL_HEADERS);
  if (sheetName === "sesi_siswa") ensureHeaders(sheet, SESI_HEADERS);
  _shCache[sheetName] = sheet;
  return sheet;
}

// ===== CACHE SERVER (CacheService) =====
const CACHE_TTL = 21600; // 6 jam (maksimum CacheService)
// Aksi tulis -> grup cache yang harus dibuang
const INVALIDATE = {
  addSiswa: ["siswa"], editSiswa: ["siswa"], deleteSiswa: ["siswa"], importSiswaBatch: ["siswa"],
  addMapel: ["mapel"], editMapel: ["mapel", "soal"], deleteMapel: ["mapel"], importMapelBatch: ["mapel"],
  uploadSoal: ["soal"], editSoal: ["soal"], deleteSoal: ["soal"], importSoalBatch: ["soal"],
  simpanJawabanSiswa: ["hasil"], deleteHasil: ["hasil"], deleteHasilByMapel: ["hasil"], simpanKoreksiEsai: ["hasil"]
};

function cacheVer_(group) {
  const c = CacheService.getScriptCache();
  let v = c.get("ver_" + group);
  if (!v) { v = String(Date.now()); c.put("ver_" + group, v, CACHE_TTL); }
  return v;
}
function bumpCache_(groups) {
  const c = CacheService.getScriptCache();
  groups.forEach(gr => c.put("ver_" + gr, String(Date.now()), CACHE_TTL));
}
function cacheGet_(group, key) {
  const c = CacheService.getScriptCache();
  const base = "d_" + group + "_" + cacheVer_(group) + "_" + key;
  const meta = c.get(base + "_n");
  if (!meta) return null;
  const n = parseInt(meta), keys = [];
  for (let i = 0; i < n; i++) keys.push(base + "_" + i);
  const parts = c.getAll(keys);
  let s = "";
  for (let i = 0; i < n; i++) { const p = parts[base + "_" + i]; if (p === undefined) return null; s += p; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
function cachePut_(group, key, obj) {
  const c = CacheService.getScriptCache();
  const base = "d_" + group + "_" + cacheVer_(group) + "_" + key;
  const s = JSON.stringify(obj), SZ = 40000, n = Math.ceil(s.length / SZ);
  if (n > 40) return; // terlalu besar, tidak di-cache
  const o = {};
  for (let i = 0; i < n; i++) o[base + "_" + i] = s.substr(i * SZ, SZ);
  o[base + "_n"] = String(n);
  try { c.putAll(o, CACHE_TTL); } catch (e) {}
}
function cached_(group, key, fn) {
  const hit = cacheGet_(group, key);
  if (hit) return hit;
  const r = fn();
  if (r && r.status === "success") cachePut_(group, key, r);
  return r;
}

function responseJSON(res) {
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET REQUEST HANDLER
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "getMapel";
  let result = {};

  try {
    if (action === "getSiswa") {
      result = cached_("siswa", "all", getSiswaData);
    } else if (action === "getMapel") {
      result = cached_("mapel", "all", getMapelData);
    } else if (action === "getAllSoal") {
      result = cached_("soal", "all", getAllSoalData);
    } else if (action === "getSoal") {
      const kodeTes = e ? e.parameter.kode_tes : "";
      result = cached_("soal", "k_" + kodeTes, () => getSoalByKode(kodeTes));
    } else if (action === "getStatusTes") {
      result = getStatusTesData();
    } else if (action === "generateToken") {
      result = generateNewToken();
    } else if (action === "getResetRequest") {
      result = getResetRequestData();
    } else if (action === "getSesiSiswa") {
      result = getSesiSiswaData();
    } else if (action === "getHasilUjian") {
      result = cached_("hasil", "all", getHasilUjianData);
    } else if (action === "getDetailHasil") {
      result = getDetailHasil(e.parameter.id);
    } else {
      result = { status: "error", message: "Action GET tidak valid" };
    }
  } catch (err) {
    result = { status: "error", message: err.toString() };
  }

  return responseJSON(result);
}

// POST REQUEST HANDLER
function doPost(e) {
  let result = {};
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;

    if (action === "loginAdmin") {
      result = handleLoginAdmin(payload);
    } else if (action === "loginSiswa") {
      result = handleLoginSiswa(payload);
    } else if (action === "konfirmasiToken") {
      result = handleKonfirmasiToken(payload);
    } else if (action === "simpanJawabanSiswa") {
      result = handleSimpanJawabanSiswa(payload);
    } else if (action === "requestResetSiswa") {
      result = handleRequestResetSiswa(payload);
    } else if (action === "addSiswa") {
      result = addSiswa(payload);
    } else if (action === "editSiswa") {
      result = editSiswa(payload);
    } else if (action === "deleteSiswa") {
      result = deleteSiswa(payload);
    } else if (action === "importSiswaBatch") {
      result = importSiswaBatch(payload);
    } else if (action === "addMapel") {
      result = addMapel(payload);
    } else if (action === "editMapel") {
      result = editMapel(payload);
    } else if (action === "deleteMapel") {
      result = deleteMapel(payload);
    } else if (action === "importMapelBatch") {
      result = importMapelBatch(payload);
    } else if (action === "uploadSoal") {
      result = uploadSoalBatch(payload);
    } else if (action === "editSoal") {
      result = editSoal(payload);
    } else if (action === "deleteSoal") {
      result = deleteSoal(payload);
    } else if (action === "importSoalBatch") {
      result = importSoalBatch(payload);
    } else if (action === "deleteHasil") {
      result = deleteHasil(payload);
    } else if (action === "deleteHasilByMapel") {
      result = deleteHasilByMapel(payload);
    } else if (action === "simpanKoreksiEsai") {
      result = simpanKoreksiEsai(payload);
    } else if (action === "simpanStatusTes") {
      result = simpanStatusTes(payload);
    } else if (action === "toggleStatusTes") {
      result = toggleStatusTes(payload);
    } else if (action === "editStatusTes") {
      result = editStatusTes(payload);
    } else if (action === "deleteStatusTes") {
      result = deleteStatusTes(payload);
    } else if (action === "simpanProgress") {
      result = simpanProgress(payload);
    } else if (action === "resetSesiSiswa") {
      result = resetSesiSiswa(payload);
    } else if (action === "approveReset") {
      result = approveResetLogin(payload);
    } else if (action === "rejectReset") {
      result = rejectResetLogin(payload);
    } else {
      result = { status: "error", message: "Action POST tidak dikenal" };
    }
    if (result && result.status === "success" && INVALIDATE[action]) bumpCache_(INVALIDATE[action]);
  } catch (err) {
    result = { status: "error", message: err.toString() };
  }

  return responseJSON(result);
}

// ===== SESI SISWA (LOGIN / MENGERJAKAN / SELESAI / RESET) =====
function getSesi_(username) {
  const sheet = getSheet("sesi_siswa");
  const n = sheet.getLastRow() - 1;
  if (n < 1) return null;
  const u = username.toString().trim().toLowerCase();
  const col = sheet.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    if (col[i][0].toString().trim().toLowerCase() === u) {
      const r = sheet.getRange(i + 2, 1, 1, 9).getValues()[0];
      return { row: i + 2, username: r[0], nama: r[1], kelompok: r[2], status: r[3].toString(),
               kode_tes: r[4], waktu_login: r[5], last_update: r[6], session_id: (r[7] || "").toString(), ping_ms: Number(r[8]) || 0 };
    }
  }
  return null;
}

// Buat/ubah sesi satu siswa. f: {nama, kelompok, status, kode_tes, session_id, login}
function upsertSesi_(username, f) {
  const sheet = getSheet("sesi_siswa");
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const cur = getSesi_(username);
  let status, sid;
  if (cur) {
    status = f.status !== undefined ? f.status : cur.status;
    sid = f.session_id !== undefined ? f.session_id : cur.session_id;
    sheet.getRange(cur.row, 1, 1, 9).setValues([[
      cur.username,
      f.nama !== undefined ? f.nama : cur.nama,
      f.kelompok !== undefined ? f.kelompok : cur.kelompok,
      status,
      f.kode_tes !== undefined ? f.kode_tes : cur.kode_tes,
      f.login ? now : cur.waktu_login,
      now, sid, Date.now()
    ]]);
  } else {
    status = f.status || "LOGIN"; sid = f.session_id || "";
    sheet.appendRow([username, f.nama || "", f.kelompok || "", status, f.kode_tes || "", now, now, sid, Date.now()]);
  }
  // salin ke cache agar pengecekan sesi saat simpan otomatis tidak menyentuh sheet
  try { CacheService.getScriptCache().put("s_" + username.toString().trim().toLowerCase(), JSON.stringify({ status: status, sid: sid }), CACHE_TTL); } catch (e) {}
}

function getSesiSiswaData() {
  const sheet = getSheet("sesi_siswa");
  const n = sheet.getLastRow() - 1;
  if (n < 1) return { status: "success", data: [] };
  const rows = sheet.getRange(2, 1, n, 9).getValues();
  const keys = rows.filter(r => r[3] === "MENGERJAKAN").map(r => "p_" + r[0].toString().trim().toLowerCase());
  const pc = keys.length ? CacheService.getScriptCache().getAll(keys) : {};
  const adaProg = {};
  const ps = getSheet("progress_ujian"), pn = ps.getLastRow() - 1;
  if (pn >= 1) ps.getRange(2, 1, pn, 1).getValues().forEach(r => adaProg[r[0].toString().trim().toLowerCase()] = true);
  const now = Date.now();
  return { status: "success", data: rows.map(r => {
    const u = r[0].toString().trim().toLowerCase();
    let p = null;
    try { p = pc["p_" + u] ? JSON.parse(pc["p_" + u]) : null; } catch (e) {}
    let terjawab = null;
    if (p) { try { terjawab = Object.keys(JSON.parse(p.j || "{}")).length; } catch (e) { terjawab = 0; } }
    const ping = Math.max(Number(r[8]) || 0, (p && p.t) || 0);
    return {
      username: r[0], nama: r[1], kelompok: r[2], status: r[3], kode_tes: r[4],
      waktu_login: r[5] ? r[5].toString() : "",
      idle_detik: ping ? Math.max(Math.round((now - ping) / 1000), 0) : null,
      terjawab: terjawab,
      ada_progres: !!(adaProg[u] || p)
    };
  }) };
}

// ===== PROGRES UJIAN (jawaban sementara di server) =====
function getProgress_(username, kodeTes) {
  const sheet = getSheet("progress_ujian");
  const n = sheet.getLastRow() - 1;
  if (n < 1) return null;
  const u = username.toString().trim().toLowerCase(), k = kodeTes.toString().split(" (")[0].trim().toLowerCase();
  const keys = sheet.getRange(2, 1, n, 2).getValues();
  for (let i = 0; i < n; i++) {
    if (keys[i][0].toString().trim().toLowerCase() === u && keys[i][1].toString().split(" (")[0].trim().toLowerCase() === k) {
      const r = sheet.getRange(i + 2, 3, 1, 3).getValues()[0];
      return { row: i + 2, jawaban: r[0] || "{}", deadline_ms: Number(r[1]) || 0, pelanggaran: Number(r[2]) || 0 };
    }
  }
  return null;
}

function hapusProgress_(username) {
  const u = username.toString().trim().toLowerCase();
  try { CacheService.getScriptCache().remove("p_" + u); } catch (e) {}
  const sheet = getSheet("progress_ujian");
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  const col = sheet.getRange(2, 1, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) {
    if (col[i][0].toString().trim().toLowerCase() === u) sheet.deleteRow(i + 2);
  }
}

// Mulai/lanjutkan progres: deadline dihitung dari awal pengerjaan (waktu mati tetap berjalan walau terputus)
function mulaiAtauLanjutProgress_(username, kodeTes, durasiMenit) {
  const u = username.toString().trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const kodeAsli = kodeTes.toString().split(" (")[0];
  let p = null;
  try { p = JSON.parse(cache.get("p_" + u) || "null"); } catch (e) {}
  if (p && p.k.toString().split(" (")[0] !== kodeAsli) p = null;
  if (!p) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sp = getProgress_(username, kodeTes);
      if (sp) {
        p = { k: kodeTes, j: sp.jawaban, v: sp.pelanggaran, d: sp.deadline_ms };
      } else {
        const d = Date.now() + durasiMenit * 60000;
        getSheet("progress_ujian").appendRow([username, kodeTes, "{}", d, 0, Date.now()]);
        p = { k: kodeTes, j: "{}", v: 0, d: d };
      }
      try { cache.put("p_" + u, JSON.stringify(p), CACHE_TTL); } catch (e) {}
    } finally { lock.releaseLock(); }
  }
  let jawaban = {};
  try { jawaban = JSON.parse(p.j || "{}"); } catch (e) {}
  return {
    sisa_detik: Math.max(Math.floor((p.d - Date.now()) / 1000), 0),
    jawaban: jawaban,
    pelanggaran: p.v || 0,
    dilanjutkan: Object.keys(jawaban).length > 0 || (p.v || 0) > 0
  };
}

// Jalur cepat: cek sesi & simpan ke CacheService (tanpa menyentuh sheet).
// Sheet hanya ditulis bila payload.persist = true (berkala / saat tab ditutup).
function simpanProgress(payload) {
  const u = (payload.username || "").toString().trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  let ses = null;
  try { ses = JSON.parse(cache.get("s_" + u) || "null"); } catch (e) {}
  if (!ses) { const s = getSesi_(u); if (s) ses = { status: s.status, sid: s.session_id }; }
  if (!ses || ses.status !== "MENGERJAKAN" || ses.sid !== (payload.session_id || "")) {
    return { status: "error", code: "SESI_INVALID", message: "Sesi Anda tidak lagi aktif (kemungkinan direset Proktor atau login di perangkat lain)." };
  }
  let p = null;
  try { p = JSON.parse(cache.get("p_" + u) || "null"); } catch (e) {}
  if (!p) {
    const sp = getProgress_(u, payload.kode_tes);
    p = sp ? { k: payload.kode_tes, j: sp.jawaban, v: sp.pelanggaran, d: sp.deadline_ms } : { k: payload.kode_tes, j: "{}", v: 0, d: 0 };
  }
  if (payload.jawaban_json !== undefined) { p.j = payload.jawaban_json; p.v = parseInt(payload.pelanggaran) || 0; }
  p.t = Date.now();
  let cacheOk = true;
  try { cache.put("p_" + u, JSON.stringify(p), CACHE_TTL); } catch (e) { cacheOk = false; } // mis. terlalu besar
  if (payload.persist || !cacheOk) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sp = getProgress_(u, payload.kode_tes);
      if (sp) getSheet("progress_ujian").getRange(sp.row, 3, 1, 4).setValues([[p.j, sp.deadline_ms, p.v, Date.now()]]);
      upsertSesi_(payload.username, {});
    } finally { lock.releaseLock(); }
  }
  return { status: "success" };
}

// Proktor membuka kunci sesi siswa dari dashboard (hapus_progress = mulai ulang dari awal)
function resetSesiSiswa(payload) {
  const sesi = getSesi_(payload.username || "");
  if (!sesi) return { status: "error", message: "Sesi siswa tidak ditemukan" };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    upsertSesi_(payload.username, { status: "RESET", session_id: "" });
    if (payload.hapus_progress) hapusProgress_(payload.username);
  } finally { lock.releaseLock(); }
  return { status: "success", message: payload.hapus_progress
    ? "Sesi direset. Siswa akan mengulang ujian dari awal."
    : "Sesi direset. Siswa bisa login dan melanjutkan ujian." };
}

// FITUR INTERAKSI SISWA
function handleLoginSiswa(payload) {
  const sisRes = cached_("siswa", "all", getSiswaData);
  const data = ((sisRes && sisRes.data) || []).map(s => [s.username, s.nama, s.kelompok, s.nik]);
  if (data.length === 0) return { status: "error", message: "Data siswa masih kosong" };

  for (let i = 0; i < data.length; i++) {
    const user = data[i][0].toString().trim();
    const nik = data[i][3].toString().trim();

    if (user.toLowerCase() === payload.username.toString().trim().toLowerCase() && 
        nik.toLowerCase() === payload.password.toString().trim().toLowerCase()) {
      // Cegah login ganda: sesi aktif di perangkat lain harus di-reset Proktor
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const sesi = getSesi_(data[i][0]);
        const aktif = sesi && (sesi.status === "LOGIN" || sesi.status === "MENGERJAKAN");
        if (aktif && sesi.session_id !== (payload.session_id || "")) {
          return { status: "error", message: "Akun ini sedang aktif / sudah login di perangkat lain. Silakan klik 'Ajukan Reset Login ke Proktor' lalu minta Proktor menyetujui." };
        }
        upsertSesi_(data[i][0], { nama: data[i][1], kelompok: data[i][2], status: aktif ? sesi.status : "LOGIN",
                                  kode_tes: aktif ? sesi.kode_tes : "", session_id: payload.session_id || "", login: !aktif });
      } finally { lock.releaseLock(); }
      return {
        status: "success",
        data: {
          username: data[i][0],
          nama: data[i][1],
          kelompok: data[i][2],
          nik: data[i][3]
        }
      };
    }
  }
  return { status: "error", message: "Username atau NIK / Password Salah!" };
}

function handleKonfirmasiToken(payload) {
  const sheetTes = getSheet("status_tes");
  const dataTes = sheetTes.getDataRange().getValues();
  if (dataTes.length <= 1) return { status: "error", message: "Tidak ada ujian yang sedang aktif saat ini" };
  dataTes.shift();

  for (let i = 0; i < dataTes.length; i++) {
    const item = dataTes[i];
    const token = item[2].toString().trim().toUpperCase();
    const status = item[3].toString().trim().toUpperCase();
    const kelompok = item[4].toString().trim();

    if (token === payload.token.toString().trim().toUpperCase() && status === "AKTIF") {
      if (kelompok === payload.kelompok.toString().trim()) {
        // Cegah pengerjaan ganda: tolak jika siswa sudah punya hasil untuk mapel ini.
        // Proktor bisa membuka kesempatan lagi dengan menghapus hasilnya di menu Hasil Ujian.
        if (payload.username && sudahMengerjakan_(payload.username, item[1])) {
          return { status: "error", message: "Anda sudah mengerjakan ujian ini. Hubungi Proktor jika perlu mengulang." };
        }
        // Validasi jadwal (tanggal & jam) sesuai zona waktu script
        const cek = cekJadwalTes_(item);
        if (cek) return { status: "error", message: cek };
        const info = getInfoMapel_(item[1]);
        if (payload.username) upsertSesi_(payload.username, { status: "MENGERJAKAN", kode_tes: item[1] });
        const prog = payload.username ? mulaiAtauLanjutProgress_(payload.username, item[1], info.waktu) : null;
        return {
          status: "success",
          kode_tes: item[1],
          nama_mapel: info.nama_mapel,
          waktu: info.waktu,
          sisa_detik: prog ? prog.sisa_detik : info.waktu * 60,
          progress: prog && prog.dilanjutkan ? { jawaban: prog.jawaban, pelanggaran: prog.pelanggaran } : null
        };
      } else {
        return { status: "error", message: "Token ini bukan untuk Kelompok/Sesi Anda!" };
      }
    }
  }
  return { status: "error", message: "Token tidak valid atau Ujian belum diaktifkan oleh Proktor!" };
}


// ===== HELPER JADWAL TES =====
function toDateStr_(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return val.toString().replace(/^'/, "").split("T")[0].trim();
}
function toTimeStr_(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  const m = val.toString().replace(/^'/, "").match(/(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2, "0") + ":" + m[2] : "";
}
// Tes dianggap SELESAI bila status bukan AKTIF atau jam selesai sudah lewat
function isTesSelesai_(row) {
  if (row[3].toString().trim().toUpperCase() !== "AKTIF") return true;
  const tgl = toDateStr_(row[5]), jamSel = toTimeStr_(row[7]);
  if (!tgl) return false;
  const tz = Session.getScriptTimeZone(), now = new Date();
  const today = Utilities.formatDate(now, tz, "yyyy-MM-dd"), hm = Utilities.formatDate(now, tz, "HH:mm");
  if (today > tgl) return true;
  return today === tgl && jamSel !== "" && hm > jamSel;
}
// Kembalikan pesan error bila di luar jadwal, atau "" bila boleh mulai
function cekJadwalTes_(row) {
  const tgl = toDateStr_(row[5]), jm = toTimeStr_(row[6]), js = toTimeStr_(row[7]);
  if (!tgl) return "";
  const tz = Session.getScriptTimeZone(), now = new Date();
  const today = Utilities.formatDate(now, tz, "yyyy-MM-dd"), hm = Utilities.formatDate(now, tz, "HH:mm");
  if (today < tgl) return "Tes ini dijadwalkan pada tanggal " + tgl + ".";
  if (today > tgl) return "Jadwal tes ini sudah berakhir (" + tgl + ").";
  if (jm && hm < jm) return "Tes belum dimulai. Jadwal mulai pukul " + jm + ".";
  if (js && hm > js) return "Waktu tes sudah berakhir (selesai pukul " + js + ").";
  return "";
}
function getInfoMapel_(kodeTes) {
  const list = (cached_("mapel", "all", getMapelData).data) || [];
  const kode = kodeTes.toString().split(" (")[0].trim().toLowerCase();
  for (let i = 0; i < list.length; i++) {
    if (list[i].kode_tes.toString().trim().toLowerCase() === kode) {
      return { nama_mapel: list[i].nama_mapel, waktu: parseInt(list[i].waktu) || 60 };
    }
  }
  return { nama_mapel: kodeTes, waktu: 60 };
}

// Cek apakah username sudah punya baris hasil_tes untuk kode tes tertentu
// (kode di hasil_tes bisa berbentuk "KODE (KETERANGAN)", jadi dipotong dulu)
function sudahMengerjakan_(username, kodeTes) {
  const sheet = getSheet("hasil_tes");
  const n = sheet.getLastRow() - 1;
  if (n < 1) return false;
  const data = sheet.getRange(2, 2, n, 3).getValues(); // hanya kolom username, nama, kode_tes
  const user = username.toString().trim().toLowerCase();
  const kode = kodeTes.toString().split(" (")[0].trim().toLowerCase();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === user &&
        data[i][2].toString().split(" (")[0].trim().toLowerCase() === kode) return true;
  }
  return false;
}

function handleSimpanJawabanSiswa(payload) {
  const sheet = getSheet("hasil_tes");
  const now = new Date();
  const timeStr = now.toLocaleDateString('id-ID') + ' ' + now.toLocaleTimeString('id-ID');

  const poinEsai = parseFloat(payload.poin_esai) || 0;
  const status = poinEsai > 0 ? "MENUNGGU_KOREKSI" : "FINAL";

  sheet.appendRow([
    "HASIL-" + Date.now(),
    payload.username,
    payload.nama,
    payload.kode_tes,
    payload.nilai,                       // nilai sementara (esai dihitung 0 sampai dikoreksi)
    timeStr,
    payload.jawaban_json || "{}",        // seluruh jawaban siswa (untuk koreksi esai)
    parseFloat(payload.skor_otomatis) || 0,
    parseFloat(payload.poin_total) || 0,
    poinEsai,
    status,
    0,
    "{}"
  ]);
  if (payload.username) {
    upsertSesi_(payload.username, { status: "SELESAI", kode_tes: payload.kode_tes });
    hapusProgress_(payload.username);
  }
  return { status: "success", message: "Jawaban berhasil dikirim!" };
}

function handleRequestResetSiswa(payload) {
  const sheet = getSheet("reset_login");
  const now = new Date();
  const timeStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
  sheet.appendRow(["REQ-" + Date.now(), payload.username, payload.nama, timeStr, "PENDING"]);
  return { status: "success", message: "Permohonan reset login telah dikirim ke Proktor. Silakan beritahu Proktor!" };
}

// FITUR HASIL UJIAN
function getHasilUjianData() {
  const sheet = getSheet("hasil_tes");
  const n = sheet.getLastRow() - 1;
  if (n < 1) return { status: "success", data: [] };
  const a = sheet.getRange(2, 1, n, 6).getValues();   // tanpa kolom jawaban_json yang besar
  const b = sheet.getRange(2, 10, n, 3).getValues();
  return { status: "success", data: a.map((r, i) => ({
    id: r[0], username: r[1], nama: r[2], kode_tes: r[3], nilai: r[4], waktu_selesai: r[5],
    ada_esai: (parseFloat(b[i][0]) || 0) > 0,
    status_koreksi: b[i][1] || "FINAL",
    skor_esai: b[i][2] || 0
  })) };
}

// Hapus satu hasil ujian berdasarkan id (kolom A sheet hasil_tes)
function deleteHasil(payload) {
  const sheet = getSheet("hasil_tes");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === (payload.id || "").toString()) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "Hasil ujian siswa berhasil dihapus" };
    }
  }
  return { status: "error", message: "Data hasil ujian tidak ditemukan" };
}

// Hapus semua hasil ujian satu mata pelajaran (berdasarkan kode_tes, kolom D)
function deleteHasilByMapel(payload) {
  const kode = (payload.kode_tes || "").toString().split(" (")[0].trim();
  if (!kode) return { status: "error", message: "Kode tes tidak valid" };

  const sheet = getSheet("hasil_tes");
  const data = sheet.getDataRange().getValues();
  let jumlah = 0;
  // dari baris paling bawah agar nomor baris tidak bergeser
  for (let i = data.length - 1; i >= 1; i--) {
    const kodeRow = data[i][3].toString().split(" (")[0].trim();
    if (kodeRow === kode) {
      sheet.deleteRow(i + 1);
      jumlah++;
    }
  }
  if (jumlah === 0) return { status: "error", message: "Tidak ada hasil ujian untuk mapel tersebut" };
  return { status: "success", message: jumlah + " hasil ujian mapel " + kode + " berhasil dihapus" };
}

// Detail jawaban esai satu siswa (untuk halaman koreksi proktor/guru)
function getDetailHasil(id) {
  const sheet = getSheet("hasil_tes");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === (id || "").toString()) {
      const row = data[i];
      let jawaban = {};
      let skorTersimpan = {};
      try { jawaban = JSON.parse(row[6] || "{}"); } catch (e) {}
      try { skorTersimpan = JSON.parse(row[12] || "{}"); } catch (e) {}

      const kodeAsli = row[3].toString().split(" (")[0].trim();
      const soalRes = getSoalByKode(kodeAsli);
      const esai = (soalRes.data || [])
        .filter(sl => sl.tipe_soal === "esai")
        .map(sl => ({
          id: sl.id,
          pertanyaan: sl.pertanyaan,
          pedoman: sl.jawaban_benar,
          poin: parseFloat(sl.poin) || 1,
          jawaban: jawaban[sl.id] || "",
          skor: skorTersimpan[sl.id] !== undefined ? skorTersimpan[sl.id] : ""
        }));

      return {
        status: "success",
        data: {
          id: row[0],
          username: row[1],
          nama: row[2],
          kode_tes: row[3],
          status_koreksi: row[10] || "FINAL",
          esai: esai
        }
      };
    }
  }
  return { status: "error", message: "Data hasil ujian tidak ditemukan" };
}

// Simpan nilai esai dari guru lalu hitung ulang nilai akhir
function simpanKoreksiEsai(payload) {
  const sheet = getSheet("hasil_tes");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      const row = data[i];
      const kodeAsli = row[3].toString().split(" (")[0].trim();
      const soalRes = getSoalByKode(kodeAsli);

      // batas skor tiap soal = poin soal tersebut
      const maxPoin = {};
      (soalRes.data || []).forEach(sl => { maxPoin[sl.id] = parseFloat(sl.poin) || 1; });

      const skorBersih = {};
      let totalEsai = 0;
      const input = payload.skor || {};
      Object.keys(input).forEach(sid => {
        if (maxPoin[sid] === undefined) return;
        let v = parseFloat(input[sid]);
        if (isNaN(v) || v < 0) v = 0;
        if (v > maxPoin[sid]) v = maxPoin[sid];
        skorBersih[sid] = v;
        totalEsai += v;
      });

      const skorOtomatis = parseFloat(row[7]) || 0;
      const poinTotal = parseFloat(row[8]) || 0;
      const nilai = poinTotal > 0 ? Math.round(((skorOtomatis + totalEsai) / poinTotal) * 100) : 0;

      sheet.getRange(i + 1, 5).setValue(nilai);                       // nilai akhir
      sheet.getRange(i + 1, 11).setValue("FINAL");                    // status_koreksi
      sheet.getRange(i + 1, 12).setValue(totalEsai);                  // skor_esai
      sheet.getRange(i + 1, 13).setValue(JSON.stringify(skorBersih)); // skor per soal

      return { status: "success", message: "Koreksi esai tersimpan. Nilai akhir: " + nilai, nilai: nilai };
    }
  }
  return { status: "error", message: "Data hasil ujian tidak ditemukan" };
}

// FITUR REQUEST RESET LOGIN
function getResetRequestData() {
  const sheet = getSheet("reset_login");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "success", data: [] };
  data.shift();

  const list = data.map(row => ({
    id: row[0],
    username: row[1],
    nama: row[2],
    waktu_request: row[3],
    status: row[4]
  }));
  return { status: "success", data: list };
}

function approveResetLogin(payload) {
  const sheet = getSheet("reset_login");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      sheet.getRange(i + 1, 5).setValue("DISETUJUI");
      // Buka kunci sesi agar siswa bisa login lagi
      const sesi = getSesi_(data[i][1]);
      if (sesi) upsertSesi_(data[i][1], { status: "RESET", session_id: "" });
      return { status: "success", message: "Reset Login siswa telah disetujui!" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

function rejectResetLogin(payload) {
  const sheet = getSheet("reset_login");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      sheet.getRange(i + 1, 5).setValue("DITOLAK");
      return { status: "success", message: "Permohonan reset login ditolak." };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

// FITUR STATUS TES & PROKTOR
function getStatusTesData() {
  const sheet = getSheet("status_tes");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "success", data: [] };
  data.shift();

  return {
    status: "success",
    data: data.map(row => ({
      id: row[0],
      kode_tes: row[1],
      token: row[2],
      status: row[3],
      kelompok: row[4],
      tanggal: toDateStr_(row[5]) || "-",
      jam_mulai: toTimeStr_(row[6]) || "-",
      jam_selesai: toTimeStr_(row[7]) || "-",
      selesai: isTesSelesai_(row)
    }))
  };
}

function generateNewToken() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let token = "";
  for (let i = 0; i < 6; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return { status: "success", token: token };
}

function simpanStatusTes(payload) {
  const sheet = getSheet("status_tes");
  
  sheet.appendRow([
    "TES-" + Date.now(),
    payload.kode_tes,
    payload.token,
    "AKTIF",
    payload.kelompok || "1",
    "'" + (payload.tanggal || ""),
    "'" + (payload.jam_mulai || ""),
    "'" + (payload.jam_selesai || "")
  ]);
  return { status: "success", message: "Tes Berhasil Dirilis!" };
}

function toggleStatusTes(payload) {
  const sheet = getSheet("status_tes");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      sheet.getRange(i + 1, 4).setValue(payload.status);
      return { status: "success", message: "Status Tes Diperbarui" };
    }
  }
  return { status: "error", message: "Data tes tidak ditemukan" };
}

// Edit jadwal/token tes - hanya boleh bila tes sudah selesai
function editStatusTes(payload) {
  const sheet = getSheet("status_tes");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      if (!isTesSelesai_(data[i])) return { status: "error", message: "Tes masih berlangsung. Hentikan tes terlebih dahulu sebelum mengedit." };
      const r = i + 1;
      sheet.getRange(r, 2, 1, 2).setValues([[payload.kode_tes, (payload.token || "").toString().toUpperCase()]]);
      sheet.getRange(r, 5).setValue(payload.kelompok || "1");
      sheet.getRange(r, 6, 1, 3).setNumberFormat("@").setValues([[payload.tanggal || "", payload.jam_mulai || "", payload.jam_selesai || ""]]);
      sheet.getRange(r, 4).setValue(payload.aktifkan ? "AKTIF" : "NONAKTIF");
      return { status: "success", message: payload.aktifkan ? "Tes diperbarui dan diaktifkan kembali." : "Data tes berhasil diperbarui." };
    }
  }
  return { status: "error", message: "Data tes tidak ditemukan" };
}

// Hapus riwayat tes - hanya boleh bila tes sudah selesai
function deleteStatusTes(payload) {
  const sheet = getSheet("status_tes");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      if (!isTesSelesai_(data[i])) return { status: "error", message: "Tes masih berlangsung. Hentikan tes terlebih dahulu sebelum menghapus." };
      sheet.deleteRow(i + 1);
      return { status: "success", message: "Data tes berhasil dihapus" };
    }
  }
  return { status: "error", message: "Data tes tidak ditemukan" };
}

function handleLoginAdmin(payload) {
  const sheet = getSheet("admin");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    if (payload.idProktor === "U03220006-DK4U" && payload.password === "12345") return { status: "success", message: "Login Berhasil" };
  } else {
    data.shift();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString() === payload.idProktor.toString() && data[i][1].toString() === payload.password.toString()) return { status: "success", message: "Login Berhasil" };
    }
  }
  return { status: "error", message: "ID Proktor atau Password Salah" };
}

function getMapelData() {
  const sheet = getSheet("mapel");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "success", data: [] };
  data.shift();
  return { status: "success", data: data.map(row => ({ kode_tes: row[0], nama_mapel: row[1], waktu: row[2] })) };
}

function addMapel(payload) { getSheet("mapel").appendRow([payload.kode_tes, payload.nama_mapel, payload.waktu]); return { status: "success", message: "Mata pelajaran berhasil ditambahkan" }; }
function editMapel(payload) {
  const sheet = getSheet("mapel");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.originalKodeTes.toString()) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[payload.kode_tes, payload.nama_mapel, payload.waktu]]);
      let n = 0;
      if (payload.kode_tes.toString() !== payload.originalKodeTes.toString()) {
        n += gantiKodeTes_("soal", 2, payload.originalKodeTes, payload.kode_tes);
        n += gantiKodeTes_("status_tes", 2, payload.originalKodeTes, payload.kode_tes);
      }
      return { status: "success", message: "Mata pelajaran diperbarui" + (n ? " (" + n + " soal/jadwal ikut diperbarui)" : "") };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}
// Ganti kode tes pada sheet lain agar data tetap sinkron
function gantiKodeTes_(sheetName, colNo, lama, baru) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][colNo - 1].toString() === lama.toString()) { sheet.getRange(i + 1, colNo).setValue(baru); n++; }
  }
  return n;
}
function deleteMapel(payload) {
  const sheet = getSheet("mapel");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.kode_tes.toString()) { sheet.deleteRow(i + 1); return { status: "success", message: "Mata pelajaran dihapus" }; }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}
function importMapelBatch(payloadList) {
  const sheet = getSheet("mapel");
  payloadList.forEach(item => sheet.appendRow([item.kode_tes || item.KODE_TES || "", item.nama_mapel || item.NAMA_MAPEL || "", item.waktu || item.WAKTU || 60]));
  return { status: "success", message: `${payloadList.length} Mapel berhasil diimport!` };
}

function getSiswaData() {
  const sheet = getSheet("siswa");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "success", data: [] };
  data.shift();
  return { status: "success", data: data.map(row => ({ username: row[0], nama: row[1], kelompok: row[2], nik: row[3] })) };
}

function addSiswa(payload) { getSheet("siswa").appendRow([payload.username, payload.nama, payload.kelompok, payload.nik]); return { status: "success", message: "Siswa berhasil disimpan" }; }
function editSiswa(payload) {
  const sheet = getSheet("siswa");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.originalUsername.toString()) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[payload.username, payload.nama, payload.kelompok, payload.nik]]);
      return { status: "success", message: "Data siswa diperbarui" };
    }
  }
  return { status: "error", message: "Username tidak ditemukan" };
}
function deleteSiswa(payload) {
  const sheet = getSheet("siswa");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.username.toString()) { sheet.deleteRow(i + 1); return { status: "success", message: "Data siswa dihapus" }; }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}
function importSiswaBatch(payloadList) {
  const sheet = getSheet("siswa");
  payloadList.forEach(item => sheet.appendRow([item.username || item.USERNAME || "", item.nama || item.NAMA || "", item.kelompok || item.KELOMPOK || 1, item.nik || item.NIK || ""]));
  return { status: "success", message: `${payloadList.length} Siswa berhasil diimport!` };
}

// FITUR KELOLA SOAL (PILIHAN GANDA, BENAR/SALAH, JAWABAN SINGKAT, ESAI)
function normalizeTipeSoal(t) {
  const x = (t || "").toString().trim().toLowerCase().replace(/[\s\-\/]+/g, "_");
  if (!x) return "pilihan_ganda";
  if (x === "pg" || x.indexOf("ganda") > -1 || x.indexOf("pilihan") > -1) return "pilihan_ganda";
  if (x === "bs" || x.indexOf("benar") > -1 || x.indexOf("salah") > -1) return "benar_salah";
  if (x.indexOf("isian") > -1 || x.indexOf("singkat") > -1) return "isian_singkat";
  if (x.indexOf("esai") > -1 || x.indexOf("essay") > -1 || x.indexOf("uraian") > -1) return "esai";
  return "pilihan_ganda";
}

function normalizeBenarSalah(v) {
  const t = (v || "").toString().trim().toUpperCase();
  if (["B", "BENAR", "TRUE", "T", "YA", "1"].indexOf(t) > -1) return "BENAR";
  if (["S", "SALAH", "FALSE", "F", "TIDAK", "0"].indexOf(t) > -1) return "SALAH";
  return "BENAR";
}

function strVal(v) {
  return (v === undefined || v === null) ? "" : v.toString().trim();
}

// Susun 1 baris sheet "soal" sesuai tipe
function buildSoalRow(id, kodeTes, item) {
  const tipe = normalizeTipeSoal(item.tipe_soal || item.TIPE_SOAL);
  const isPG = tipe === "pilihan_ganda";
  const raw = strVal(item.jawaban_benar !== undefined ? item.jawaban_benar : item.JAWABAN_BENAR);

  let jawaban;
  if (tipe === "pilihan_ganda") jawaban = (raw || "A").toUpperCase();
  else if (tipe === "benar_salah") jawaban = normalizeBenarSalah(raw);
  else jawaban = raw; // isian singkat & esai: teks apa adanya

  return {
    tipe: tipe,
    row: [
      id,
      kodeTes,
      strVal(item.pertanyaan !== undefined ? item.pertanyaan : item.PERTANYAAN) || "",
      isPG ? strVal(item.opsi_a || item.OPSI_A) : "",
      isPG ? strVal(item.opsi_b || item.OPSI_B) : "",
      isPG ? strVal(item.opsi_c || item.OPSI_C) : "",
      isPG ? strVal(item.opsi_d || item.OPSI_D) : "",
      isPG ? strVal(item.opsi_e || item.OPSI_E) : "",
      jawaban,
      item.poin || item.POIN || 1,
      tipe
    ]
  };
}

// Kunci isian singkat dipaksa berformat teks agar "1/2" atau "007" tidak diubah Sheets
function setKunciSebagaiTeks(sheet, rowNum, tipe, value) {
  if (tipe === "isian_singkat") {
    sheet.getRange(rowNum, 9).setNumberFormat("@").setValue(value);
  }
}

function getAllSoalData() {
  const sheet = getSheet("soal");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "success", data: [] };
  data.shift();
  return {
    status: "success",
    data: data.map(row => {
      const tipe = normalizeTipeSoal(row[10]); // soal lama tanpa tipe = pilihan ganda
      const kunci = (row[8] === undefined || row[8] === null) ? "" : row[8].toString();
      return {
        id: row[0],
        kode_tes: row[1],
        tipe_soal: tipe,
        pertanyaan: row[2],
        opsi_a: row[3],
        opsi_b: row[4],
        opsi_c: row[5],
        opsi_d: row[6],
        opsi_e: row[7] || "",
        jawaban_benar: tipe === "pilihan_ganda" ? (kunci || "A") : kunci,
        poin: row[9] || 1
      };
    })
  };
}

function getSoalByKode(kodeTes) {
  const res = cached_("soal", "all", getAllSoalData);
  if (res.status === "success") {
    return { status: "success", data: res.data.filter(item => item.kode_tes.toString() === kodeTes.toString()) };
  }
  return res;
}

function uploadSoalBatch(payload) {
  const sheet = getSheet("soal");
  let listSoal = Array.isArray(payload.soal) ? payload.soal : (payload.pertanyaan ? [payload] : []);
  if (listSoal.length === 0) return { status: "error", message: "Data soal kosong" };

  listSoal.forEach((item, index) => {
    const built = buildSoalRow("SOAL-" + Date.now() + "-" + index, payload.kode_tes || item.kode_tes || "", item);
    sheet.appendRow(built.row);
    setKunciSebagaiTeks(sheet, sheet.getLastRow(), built.tipe, built.row[8]);
  });
  return { status: "success", message: "Soal berhasil disimpan!" };
}

function editSoal(payload) {
  const sheet = getSheet("soal");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) {
      const built = buildSoalRow(payload.id, payload.kode_tes, payload);
      sheet.getRange(i + 1, 1, 1, SOAL_HEADERS.length).setValues([built.row]);
      setKunciSebagaiTeks(sheet, i + 1, built.tipe, built.row[8]);
      return { status: "success", message: "Soal diperbarui" };
    }
  }
  return { status: "error", message: "ID Soal tidak ditemukan" };
}

function deleteSoal(payload) {
  const sheet = getSheet("soal");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === payload.id.toString()) { sheet.deleteRow(i + 1); return { status: "success", message: "Soal dihapus" }; }
  }
  return { status: "error", message: "Soal tidak ditemukan" };
}

function importSoalBatch(payloadList) {
  const sheet = getSheet("soal");
  payloadList.forEach((item, index) => {
    const built = buildSoalRow("SOAL-" + Date.now() + "-" + index, item.kode_tes || item.KODE_TES || "", item);
    sheet.appendRow(built.row);
    setKunciSebagaiTeks(sheet, sheet.getLastRow(), built.tipe, built.row[8]);
  });
  return { status: "success", message: `${payloadList.length} Soal berhasil diimport!` };
}
