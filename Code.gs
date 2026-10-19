// ID Google Sheets
const SPREADSHEET_ID = "1bdNIooE1uHFcnSuOZnVBu2YmLZua53BpXnxaS3NiUOQ";

// Header sheet yang punya kolom tambahan (tipe soal & koreksi esai)
const SOAL_HEADERS  = ["id", "kode_tes", "pertanyaan", "opsi_a", "opsi_b", "opsi_c", "opsi_d", "opsi_e", "jawaban_benar", "poin", "tipe_soal"];
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

function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    }
  }
  if (sheetName === "soal") ensureHeaders(sheet, SOAL_HEADERS);
  if (sheetName === "hasil_tes") ensureHeaders(sheet, HASIL_HEADERS);
  return sheet;
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
      result = getSiswaData();
    } else if (action === "getMapel") {
      result = getMapelData();
    } else if (action === "getAllSoal") {
      result = getAllSoalData();
    } else if (action === "getSoal") {
      const kodeTes = e ? e.parameter.kode_tes : "";
      result = getSoalByKode(kodeTes);
    } else if (action === "getStatusTes") {
      result = getStatusTesData();
    } else if (action === "generateToken") {
      result = generateNewToken();
    } else if (action === "getResetRequest") {
      result = getResetRequestData();
    } else if (action === "getHasilUjian") {
      result = getHasilUjianData();
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
    } else if (action === "approveReset") {
      result = approveResetLogin(payload);
    } else if (action === "rejectReset") {
      result = rejectResetLogin(payload);
    } else {
      result = { status: "error", message: "Action POST tidak dikenal" };
    }
  } catch (err) {
    result = { status: "error", message: err.toString() };
  }

  return responseJSON(result);
}

// FITUR INTERAKSI SISWA
function handleLoginSiswa(payload) {
  const sheet = getSheet("siswa");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "error", message: "Data siswa masih kosong" };
  data.shift();

  for (let i = 0; i < data.length; i++) {
    const user = data[i][0].toString().trim();
    const nik = data[i][3].toString().trim();

    if (user.toLowerCase() === payload.username.toString().trim().toLowerCase() && 
        nik.toLowerCase() === payload.password.toString().trim().toLowerCase()) {
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
        return {
          status: "success",
          kode_tes: item[1]
        };
      } else {
        return { status: "error", message: "Token ini bukan untuk Kelompok/Sesi Anda!" };
      }
    }
  }
  return { status: "error", message: "Token tidak valid atau Ujian belum diaktifkan oleh Proktor!" };
}

// Cek apakah username sudah punya baris hasil_tes untuk kode tes tertentu
// (kode di hasil_tes bisa berbentuk "KODE (KETERANGAN)", jadi dipotong dulu)
function sudahMengerjakan_(username, kodeTes) {
  const data = getSheet("hasil_tes").getDataRange().getValues();
  const user = username.toString().trim().toLowerCase();
  const kode = kodeTes.toString().split(" (")[0].trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim().toLowerCase() === user &&
        data[i][3].toString().split(" (")[0].trim().toLowerCase() === kode) {
      return true;
    }
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
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "success", data: [] };
  data.shift();

  const list = data.map(row => ({
    id: row[0],
    username: row[1],
    nama: row[2],
    kode_tes: row[3],
    nilai: row[4],
    waktu_selesai: row[5],
    ada_esai: (parseFloat(row[9]) || 0) > 0,
    status_koreksi: row[10] || "FINAL",
    skor_esai: row[11] || 0
  }));
  return { status: "success", data: list };
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

  const helperFormatDate = (val) => {
    if (!val) return "-";
    if (val instanceof Date) {
      const y = val.getFullYear();
      const m = String(val.getMonth() + 1).padStart(2, '0');
      const d = String(val.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return val.toString().split('T')[0];
  };

  const helperFormatTime = (val) => {
    if (!val) return "-";
    if (val instanceof Date) {
      const h = String(val.getHours()).padStart(2, '0');
      const m = String(val.getMinutes()).padStart(2, '0');
      return `${h}:${m}`;
    }
    return val.toString().split('T')[0];
  };

  return {
    status: "success",
    data: data.map(row => ({
      id: row[0],
      kode_tes: row[1],
      token: row[2],
      status: row[3],
      kelompok: row[4],
      tanggal: helperFormatDate(row[5]),
      jam_mulai: helperFormatTime(row[6]),
      jam_selesai: helperFormatTime(row[7])
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
      return { status: "success", message: "Mata pelajaran diperbarui" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
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
  const res = getAllSoalData();
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
