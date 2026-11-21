// ====================================================
// KODE.GS - RAPOT DIGITAL ACS 2026 - v12
// ====================================================
// PERUBAHAN UTAMA v12 (dibanding snapshot 23 Agu 2026):
// - Semua kredensial hardcode DIHAPUS. Token WA dan kredensial
//   admin hanya dibaca dari Script Properties.
// - Login memakai token sesi server-side (CacheService, 6 jam)
//   dan hash SHA-256 untuk semua password.
// - Semua fungsi backend (kecuali doGet, prosesLogin,
//   hitungHashSha256, dan helper internal) menerima token
//   sebagai argumen pertama dan memvalidasi peran.
// - Modul Presensi, SPP, Registrasi Atlet Baru, dan
//   Kalkulator Pacing dibangun kembali (kontrak v11).
//
// CARA MENYIAPKAN KREDENSIAL (dilakukan pemilik, sekali saja,
// di editor Apps Script > Project Settings > Script Properties):
//   1. Jalankan fungsi hitungHashSha256("password-admin-yang-dipilih")
//      dari editor, lalu salin hasil hex yang tampil di log.
//   2. Isi Script Properties:
//        ADMIN_USERNAME       = username pelatih/admin
//        ADMIN_PASSWORD_HASH  = hex SHA-256 dari langkah 1
//        FONNTE_TOKEN         = token WA (bila memakai Fonnte)
//   3. Akun wali dibuat lewat simpanRegistrasiAtletBaru atau
//      ditambah manual di sheet USER_ACCOUNT dengan kolom
//      PASSWORD_HASH berisi hex SHA-256 (bukan password asli).
//   4. Jangan pernah menulis token/password asli di dalam kode,
//      komentar, maupun contoh di file ini.
// ====================================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Rapot Digital ACS 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ====================================================
// HELPER & UTILITIES
// ====================================================

function getColumnIndex(headers, possibleNames) {
  if (!headers || !headers.length) return -1;
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").toUpperCase().trim();
    for (var j = 0; j < possibleNames.length; j++) {
      if (h === possibleNames[j].toUpperCase().trim()) return i;
    }
  }
  return -1;
}

// Normalisasi nama untuk pencocokan PERSIS: huruf besar semua,
// spasi beruntun jadi satu spasi, tanpa spasi di ujung.
// Bandingkan selalu dengan === pada hasil fungsi ini.
function normalisasiNama(s) {
  return String(s || "").toUpperCase().replace(/\s+/g, " ").trim();
}

function standarkanNomorLomba(str) {
  if (!str) return "";
  var clean = String(str).toUpperCase().replace(/\s+/g, ' ').trim();
  var jarakMatch = clean.match(/(25|50|100|200|400|800|1500)/);
  if (!jarakMatch) return clean;
  var jarak = jarakMatch[1] + "m";

  var gaya = "FR";
  if (clean.indexOf("BR") !== -1 || clean.indexOf("DADA") !== -1) gaya = "BR";
  else if (clean.indexOf("BK") !== -1 || clean.indexOf("PUNGGUNG") !== -1) gaya = "BK";
  else if (clean.indexOf("BU") !== -1 || clean.indexOf("KUPU") !== -1) gaya = "BU";
  else if (clean.indexOf("IM") !== -1 || clean.indexOf("GANTI") !== -1) gaya = "IM";

  return jarak + " " + gaya;
}

function parseKeDetik(teksWaktu) {
  if (!teksWaktu) return 0;
  var str = String(teksWaktu).trim();
  if (str.indexOf("GMT") !== -1 || str.indexOf("1941") !== -1 || /[a-zA-Z]{3,}/.test(str)) return 0;

  var bersih = str.replace(',', '.').replace(/[^0-9\.:]/g, '').trim();
  if (!bersih) return 0;

  if (bersih.indexOf(':') !== -1) {
    var p = bersih.split(':');
    var menit = parseFloat(p[0]) || 0;
    var detik = parseFloat(p[1]) || 0;
    return (menit * 60) + detik;
  }

  var d = parseFloat(bersih) || 0;
  return (d < 2 || d > 7200) ? 0 : d;
}

function formatTanggalStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return String(val);
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(val, tz, "dd/MM/yyyy");
  }
  var str = String(val).trim();
  if (!str) return "";
  // Coba parse sebagai tanggal; bila valid, format dd/MM/yyyy.
  // Bila tidak valid, kembalikan string asli (jangan pernah
  // mengembalikan tanggal hardcode).
  var d = new Date(str);
  if (!isNaN(d.getTime())) {
    var tz2 = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(d, tz2, "dd/MM/yyyy");
  }
  if (str.indexOf("T") !== -1) str = str.split("T")[0];
  if (str.indexOf("-") !== -1) {
    var p = str.split("-");
    if (p.length === 3 && p[0].length === 4) return p[2] + "/" + p[1] + "/" + p[0];
  }
  return str;
}

// Format detik ke "m:ss.cc", contoh 36.26 -> "0:36.26".
function formatWaktuDetik(detik) {
  var d = Number(detik) || 0;
  var menit = Math.floor(d / 60);
  var sisa = d - menit * 60;
  var sisaStr = sisa.toFixed(2);
  if (Number(sisaStr) < 10) sisaStr = "0" + sisaStr;
  return menit + ":" + sisaStr;
}

// Batas bawah wajar (detik) per jarak, dan batas atas umum.
var BATAS_WAKTU_MIN = { "25": 8, "50": 15, "100": 30, "200": 70, "400": 160, "800": 360, "1500": 700 };
var BATAS_WAKTU_MAX = 7200;

// Mengembalikan string error bila waktu di luar batas, atau null bila wajar.
// Pesan selalu menyebut nomor lomba dan batasnya.
function validasiWaktu(nomorLomba, detik) {
  var nomor = standarkanNomorLomba(nomorLomba);
  var m = nomor.match(/(25|50|100|200|400|800|1500)/);
  var batasMin = m ? BATAS_WAKTU_MIN[m[1]] : 0;
  if (!(detik > 0)) return "Waktu tidak valid untuk nomor " + nomor + ".";
  if (detik > BATAS_WAKTU_MAX) return "Waktu " + detik + " detik untuk nomor " + nomor + " melebihi batas atas " + BATAS_WAKTU_MAX + " detik.";
  if (batasMin && detik < batasMin) return "Waktu " + detik + " detik untuk nomor " + nomor + " di bawah batas bawah wajar " + batasMin + " detik.";
  return null;
}

function namaBulanIndonesia(bulanIndex) {
  var nama = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return nama[bulanIndex] || "";
}

// Ubah berbagai bentuk "bulan tahun" menjadi {bulanIndex, tahun, nama}.
// Menerima: "November 2026", "Nov 2026", "11/2026", "2026-11", "11-2026".
// Mengembalikan null bila tidak bisa diparse.
function parseBulanTahun(str) {
  if (!str) return null;
  var s = String(str).trim();
  var namaPanjang = ["januari", "februari", "maret", "april", "mei", "juni", "juli", "agustus", "september", "oktober", "november", "desember"];
  var namaSingkat = ["jan", "feb", "mar", "apr", "mei", "jun", "jul", "agu", "sep", "okt", "nov", "des"];
  var lower = s.toLowerCase();
  var tahunMatch = s.match(/(19|20)\d{2}/);
  var tahun = tahunMatch ? parseInt(tahunMatch[0], 10) : null;
  var bulanIndex = -1;
  for (var i = 0; i < 12; i++) {
    if (lower.indexOf(namaPanjang[i]) !== -1 || lower.indexOf(namaSingkat[i]) !== -1) { bulanIndex = i; break; }
  }
  if (bulanIndex === -1) {
    var angka = s.match(/(\d{1,2})[\/\-](\d{4})|(\d{4})[\/\-](\d{1,2})/);
    if (angka) {
      if (angka[1] && angka[2]) { bulanIndex = parseInt(angka[1], 10) - 1; tahun = parseInt(angka[2], 10); }
      else if (angka[3] && angka[4]) { bulanIndex = parseInt(angka[4], 10) - 1; tahun = parseInt(angka[3], 10); }
    } else {
      var hanyaAngka = s.match(/^(\d{1,2})$/);
      if (hanyaAngka) bulanIndex = parseInt(hanyaAngka[1], 10) - 1;
    }
  }
  if (bulanIndex < 0 || bulanIndex > 11) return null;
  return { bulanIndex: bulanIndex, tahun: tahun, nama: namaBulanIndonesia(bulanIndex) };
}

// Format "bulan tahun" apa pun menjadi "Nama Bulan Indonesia YYYY"
// (atau hanya nama bulan bila tahun tidak diketahui).
function formatBulanTahunIndonesia(str) {
  var p = parseBulanTahun(str);
  if (!p) return String(str || "");
  return p.tahun ? (p.nama + " " + p.tahun) : p.nama;
}

// "yyyy-MM-dd" -> "dd/MM/yyyy". Bila sudah dd/MM/yyyy, dikembalikan apa adanya.
function tanggalYmdKeDmy(ymd) {
  var s = String(ymd || "").trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    var p = s.substring(0, 10).split("-");
    return p[2] + "/" + p[1] + "/" + p[0];
  }
  return formatTanggalStr(s);
}

// Apakah dua nilai tanggal merujuk hari yang sama (dibandingkan dd/MM/yyyy)?
function tanggalSama(a, b) {
  return formatTanggalStr(a) === formatTanggalStr(b);
}

// Jalankan fn di dalam ScriptLock. Bila lock gagal, kembalikan error sibuk.
function denganLock(fn) {
  var lock = LockService.getScriptLock();
  var ok = false;
  try {
    ok = lock.tryLock(10000);
  } catch (e) {
    ok = false;
  }
  if (!ok) return { error: "Sistem sibuk, coba lagi sebentar." };
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function hapusCacheDaftarAtlet(atletKey) {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove("DAFTAR_ATLET_pelatih_");
    if (atletKey) cache.remove("DAFTAR_ATLET_wali_" + normalisasiNama(atletKey));
  } catch (e) {}
}

// ====================================================
// KEAMANAN - HASH, SESI, DAN PERAN
// ====================================================

// Mengembalikan hex SHA-256 dari teks.
// CARA PAKAI (pemilik, sekali saja, di editor Apps Script):
// jalankan fungsi ini dengan password yang dipilih, salin hasil
// hex dari log, lalu tempelkan sebagai nilai Script Properties
// ADMIN_PASSWORD_HASH (untuk pelatih) atau kolom PASSWORD_HASH
// di sheet USER_ACCOUNT (untuk wali). Jangan simpan password
// asli di mana pun di kode atau di Sheet.
function hitungHashSha256(teks) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(teks || ""), Utilities.Charset.UTF_8);
  var hex = "";
  for (var i = 0; i < bytes.length; i++) {
    var v = (bytes[i] < 0 ? bytes[i] + 256 : bytes[i]).toString(16);
    if (v.length === 1) v = "0" + v;
    hex += v;
  }
  return hex;
}

function adalahHashSha256(s) {
  return /^[0-9a-f]{64}$/i.test(String(s || "").trim());
}

// Buat password acak 8 karakter (huruf + angka). Hanya dipakai
// saat registrasi atlet baru; password aslinya tidak disimpan.
function buatPasswordAcak() {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789";
  var hasil = "";
  var uuid = Utilities.getUuid().replace(/-/g, "");
  for (var i = 0; i < 8; i++) {
    var idx = Math.floor(Math.random() * chars.length);
    // Campur sumber acak agar tidak bergantung pada satu sumber saja.
    if (i < uuid.length) idx = (idx + parseInt(uuid.charAt(i), 16)) % chars.length;
    hasil += chars.charAt(idx);
  }
  return hasil;
}

function buatTokenSesi() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
}

// Internal: baca sesi dari cache. Mengembalikan objek sesi atau null.
function validasiToken(token) {
  if (!token) return null;
  try {
    var raw = CacheService.getScriptCache().get("SESS_" + String(token));
    if (!raw) return null;
    var sesi = JSON.parse(raw);
    if (!sesi || !sesi.role) return null;
    return sesi;
  } catch (e) {
    return null;
  }
}

// Internal: validasi token + peran. Mengembalikan objek sesi bila
// lolos; melempar Error("SESI_HABIS") atau Error("AKSES_DITOLAK")
// bila tidak. Pemanggil publik menangkap error ini dan mengubahnya
// menjadi {error: "..."}.
function butuhPeran(token, peranYangDiizinkan) {
  var sesi = validasiToken(token);
  if (!sesi) throw new Error("SESI_HABIS");
  var daftar = Array.isArray(peranYangDiizinkan) ? peranYangDiizinkan : [peranYangDiizinkan];
  if (daftar.indexOf(sesi.role) === -1) throw new Error("AKSES_DITOLAK");
  return sesi;
}

// Internal: apakah sesi wali ini boleh mengakses data atlet bernama
// `nama` (dibandingkan persis setelah normalisasi, terhadap
// panggilan maupun nama lengkap yang dikenal)? Pelatih selalu boleh.
function bolehAksesAtlet(sesi, nama, namaLengkap) {
  if (!sesi) return false;
  if (sesi.role === "pelatih") return true;
  var kunci = normalisasiNama(sesi.atletKey);
  if (!kunci) return false;
  if (normalisasiNama(nama) === kunci) return true;
  if (namaLengkap && normalisasiNama(namaLengkap) === kunci) return true;
  return false;
}

// Internal: ambil (atau buat bila belum ada) sheet USER_ACCOUNT
// dengan header USERNAME, PASSWORD_HASH, ROLE, ATLET_KEY,
// HARUS_GANTI_PASSWORD.
function getUserAccountSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("USER_ACCOUNT");
  if (!sheet) {
    sheet = ss.insertSheet("USER_ACCOUNT");
    sheet.appendRow(["USERNAME", "PASSWORD_HASH", "ROLE", "ATLET_KEY", "HARUS_GANTI_PASSWORD"]);
  }
  return sheet;
}

// ====================================================
// MODUL CORE FONNTE WA GATEWAY
// ====================================================

// Hanya membaca Script Properties. Bila kosong, kembalikan "".
function getFonnteToken() {
  var propToken = PropertiesService.getScriptProperties().getProperty("FONNTE_TOKEN");
  return propToken ? propToken : "";
}

function kirimPesanFonnte(targetPhone, messageText) {
  var token = getFonnteToken();
  if (!token) return "Token Fonnte belum diatur di Script Properties";
  try {
    var hpClean = String(targetPhone).replace(/[^0-9]/g, '');
    if (hpClean.indexOf("08") === 0) hpClean = "628" + hpClean.substring(2);

    if (!hpClean || hpClean.length < 8) {
      return "Nomor HP tujuan tidak valid.";
    }

    var options = {
      'method': 'post',
      'headers': { 'Authorization': token },
      'payload': {
        'target': hpClean,
        'message': messageText
      },
      'muteHttpExceptions': true
    };
    var response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
    var resObj = JSON.parse(response.getContentText());
    return resObj.status ? "WA Terkirim!" : ("Gagal Kirim WA: " + (resObj.reason || resObj.detail || "Error Fonnte"));
  } catch (err) {
    return "Error Fonnte: " + err.toString();
  }
}

function kirimNotifikasiWaTagihanFonnte(token, payload) {
  try {
    butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var targetHp = payload.noHp || payload.noHpWali;
    if (!targetHp) return { error: "Nomor WhatsApp tujuan tidak ditemukan / belum diisi." };

    var namaAtlet = payload.namaLengkap || payload.namaAtlet || "Atlet";
    var namaEvent = payload.namaEvent || "Kejuaraan Renang";
    var totalTagihan = payload.totalTagihan || ("Rp " + (parseInt(payload.totalBiaya) || 0).toLocaleString('id-ID'));
    var statusBayar = payload.statusBayar || "Belum Lunas";
    var nomorDiikuti = payload.nomorDiikuti || "-";

    var pesanWA = "📋 *REKAP PENDAFTARAN LOMBA ACS BANJARNEGARA*\n" +
                  "===================================\n\n" +
                  "Yth. Wali/Orang Tua dari *" + namaAtlet.toUpperCase() + "*,\n\n" +
                  "Berikut rincian pendaftaran & tagihan kejuaraan:\n" +
                  "🏆 *Event:* " + namaEvent + "\n" +
                  "🏊‍♂️ *Nomor Lomba:* " + nomorDiikuti + "\n" +
                  "💰 *Total Tagihan:* " + totalTagihan + "\n" +
                  "📌 *Status Pembayaran:* *" + statusBayar.toUpperCase() + "*\n\n";

    if (String(statusBayar).toLowerCase() !== "lunas") {
      pesanWA += "Mohon konfirmasi jika pembayaran telah dilakukan. Terima kasih! 💦\n\n";
    } else {
      pesanWA += "Terima kasih, pembayaran telah kami terima LUNAS. 💦\n\n";
    }

    pesanWA += "_Pesan otomatis oleh System E-Rapor ACS Banjarnegara_";

    var statusKirim = kirimPesanFonnte(targetHp, pesanWA);

    if (statusKirim.indexOf("WA Terkirim") !== -1) {
      return { success: "Pesan WhatsApp tagihan berhasil terkirim via Fonnte ke " + targetHp };
    } else {
      return { error: statusKirim };
    }
  } catch (err) {
    return { error: "Exception WA Server: " + err.toString() };
  }
}

// ====================================================
// AUTHENTIKASI & USER MANAGEMENT
// ====================================================

function prosesLogin(username, password) {
  try {
    var uRaw = String(username || "").trim();
    var pRaw = String(password || "");
    if (!uRaw || !pRaw) return { error: "Username atau Password salah!" };

    var props = PropertiesService.getScriptProperties();
    var adminUser = props.getProperty("ADMIN_USERNAME") || "";
    var adminHash = props.getProperty("ADMIN_PASSWORD_HASH") || "";

    // Login pelatih/admin: username dibandingkan case-sensitive
    // apa adanya; password dicocokkan sebagai hash SHA-256,
    // case-sensitive (tanpa perubahan huruf).
    if (adminUser && uRaw === adminUser && adminHash) {
      if (hitungHashSha256(pRaw) === String(adminHash).trim().toLowerCase()) {
        var tokenAdmin = buatTokenSesi();
        var sesiAdmin = { role: "pelatih", username: adminUser, atletKey: "" };
        CacheService.getScriptCache().put("SESS_" + tokenAdmin, JSON.stringify(sesiAdmin), 21600);
        return { role: "pelatih", username: adminUser, atletKey: "", token: tokenAdmin, mustChangePassword: false };
      }
      return { error: "Username atau Password salah!" };
    }

    // Login wali: HANYA lewat sheet USER_ACCOUNT. Tidak ada lagi
    // login memakai nama panggilan / nama lengkap / password umum.
    var sheet = getUserAccountSheet();
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return { error: "Username atau Password salah!" };
    var headers = data[0];
    var idxUser = getColumnIndex(headers, ["USERNAME"]);
    var idxHash = getColumnIndex(headers, ["PASSWORD_HASH"]);
    var idxRole = getColumnIndex(headers, ["ROLE"]);
    var idxAtlet = getColumnIndex(headers, ["ATLET_KEY"]);
    var idxGanti = getColumnIndex(headers, ["HARUS_GANTI_PASSWORD"]);
    if (idxUser === -1) idxUser = 0;
    if (idxHash === -1) idxHash = 1;
    if (idxRole === -1) idxRole = 2;
    if (idxAtlet === -1) idxAtlet = 3;
    if (idxGanti === -1) idxGanti = 4;

    var uCari = uRaw.toUpperCase();
    for (var i = 1; i < data.length; i++) {
      var userTersimpan = String(data[i][idxUser] || "").trim();
      if (!userTersimpan || userTersimpan.toUpperCase() !== uCari) continue;
      var tersimpan = String(data[i][idxHash] || "");
      var cocok = false;
      if (adalahHashSha256(tersimpan)) {
        cocok = (hitungHashSha256(pRaw) === tersimpan.trim().toLowerCase());
      } else {
        // Migrasi peninggalan lama: nilai tersimpan bukan hash.
        // Bandingkan sebagai teks biasa; bila cocok, simpan ulang
        // sebagai hash agar ke depannya hanya hash yang tersimpan.
        if (tersimpan !== "" && pRaw === tersimpan) {
          cocok = true;
          try {
            sheet.getRange(i + 1, idxHash + 1).setValue(hitungHashSha256(pRaw));
          } catch (eMig) {}
        }
      }
      if (!cocok) return { error: "Username atau Password salah!" };
      var roleWali = String(data[i][idxRole] || "wali").trim().toLowerCase() || "wali";
      var atletKey = String(data[i][idxAtlet] || userTersimpan).trim();
      var harusGanti = String(data[i][idxGanti] || "").trim().toUpperCase() === "YA";
      var tokenWali = buatTokenSesi();
      var sesiWali = { role: roleWali, username: userTersimpan, atletKey: atletKey };
      CacheService.getScriptCache().put("SESS_" + tokenWali, JSON.stringify(sesiWali), 21600);
      return { role: roleWali, username: userTersimpan, atletKey: atletKey, token: tokenWali, mustChangePassword: harusGanti };
    }

    return { error: "Username atau Password salah!" };
  } catch (err) {
    return { error: err.toString() };
  }
}

function ubahPassword(token, passwordLama, passwordBaru) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  if (!passwordBaru || String(passwordBaru).length < 6) {
    return { error: "Password baru minimal 6 karakter." };
  }
  if (!passwordLama) return { error: "Password lama wajib diisi." };
  return denganLock(function() {
    try {
      if (sesi.role === "pelatih") {
        var props = PropertiesService.getScriptProperties();
        var hashLama = props.getProperty("ADMIN_PASSWORD_HASH") || "";
        if (!hashLama || hitungHashSha256(String(passwordLama)) !== String(hashLama).trim().toLowerCase()) {
          return { error: "Password lama salah." };
        }
        props.setProperty("ADMIN_PASSWORD_HASH", hitungHashSha256(String(passwordBaru)));
        return { success: "Password berhasil diubah." };
      }
      // Wali: verifikasi terhadap baris USER_ACCOUNT-nya sendiri.
      var sheet = getUserAccountSheet();
      var data = sheet.getDataRange().getDisplayValues();
      var headers = data[0] || [];
      var idxUser = getColumnIndex(headers, ["USERNAME"]);
      var idxHash = getColumnIndex(headers, ["PASSWORD_HASH"]);
      var idxGanti = getColumnIndex(headers, ["HARUS_GANTI_PASSWORD"]);
      if (idxUser === -1) idxUser = 0;
      if (idxHash === -1) idxHash = 1;
      if (idxGanti === -1) idxGanti = 4;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idxUser] || "").trim().toUpperCase() === String(sesi.username || "").trim().toUpperCase()) {
          var tersimpan = String(data[i][idxHash] || "");
          var cocok = false;
          if (adalahHashSha256(tersimpan)) {
            cocok = (hitungHashSha256(String(passwordLama)) === tersimpan.trim().toLowerCase());
          } else {
            cocok = (String(passwordLama) === tersimpan);
          }
          if (!cocok) return { error: "Password lama salah." };
          sheet.getRange(i + 1, idxHash + 1).setValue(hitungHashSha256(String(passwordBaru)));
          sheet.getRange(i + 1, idxGanti + 1).setValue("TIDAK");
          return { success: "Password berhasil diubah." };
        }
      }
      return { error: "Akun wali tidak ditemukan." };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// Reset password wali (lupa password) — HANYA pelatih.
// Password lama langsung tidak berlaku. Password baru dibuat acak
// (8 karakter, seperti saat registrasi), hanya hash-nya yang
// disimpan; password aslinya HANYA tampil sekali di hasil fungsi
// ini dan tidak disimpan di mana pun. Wali wajib menggantinya
// saat login pertama (HARUS_GANTI_PASSWORD = "YA").
// Bila atlet belum punya baris di USER_ACCOUNT (atlet lama),
// baris dibuatkan dulu setelah dipastikan atletnya ada di
// PROFIL_ATLET.
function resetPasswordWali(token, usernameAtlet) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  var kunci = normalisasiNama(usernameAtlet);
  if (!kunci) return { error: "Nama panggilan atlet wajib diisi." };
  return denganLock(function() {
    try {
      var passwordBaru = buatPasswordAcak();
      var hashBaru = hitungHashSha256(passwordBaru);
      var sheetAkun = getUserAccountSheet();
      var data = sheetAkun.getDataRange().getDisplayValues();
      var headers = data[0] || [];
      var idxUser = getColumnIndex(headers, ["USERNAME"]);
      var idxHash = getColumnIndex(headers, ["PASSWORD_HASH"]);
      var idxRole = getColumnIndex(headers, ["ROLE"]);
      var idxAtlet = getColumnIndex(headers, ["ATLET_KEY"]);
      var idxGanti = getColumnIndex(headers, ["HARUS_GANTI_PASSWORD"]);
      if (idxUser === -1) idxUser = 0;
      if (idxHash === -1) idxHash = 1;
      if (idxRole === -1) idxRole = 2;
      if (idxAtlet === -1) idxAtlet = 3;
      if (idxGanti === -1) idxGanti = 4;
      for (var i = 1; i < data.length; i++) {
        if (normalisasiNama(data[i][idxUser]) === kunci) {
          sheetAkun.getRange(i + 1, idxHash + 1).setValue(hashBaru);
          sheetAkun.getRange(i + 1, idxGanti + 1).setValue("YA");
          return { success: "Password wali untuk " + kunci + " berhasil direset.", username: kunci, passwordBaru: passwordBaru };
        }
      }
      // Belum punya akun: pastikan atletnya ada di PROFIL_ATLET,
      // lalu buatkan baris akun wali baru.
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
      if (!sheetProfil) return { error: "Atlet tidak ditemukan di data PROFIL_ATLET" };
      var dataProfil = sheetProfil.getDataRange().getDisplayValues();
      if (dataProfil.length < 2) return { error: "Atlet tidak ditemukan di data PROFIL_ATLET" };
      var idxPanggilan = getColumnIndex(dataProfil[0], ["NAMA_PANGGILAN", "PANGGILAN", "KEY"]);
      if (idxPanggilan === -1) idxPanggilan = 0;
      var ditemukan = false;
      for (var p = 1; p < dataProfil.length; p++) {
        if (normalisasiNama(dataProfil[p][idxPanggilan]) === kunci) { ditemukan = true; break; }
      }
      if (!ditemukan) return { error: "Atlet tidak ditemukan di data PROFIL_ATLET" };
      sheetAkun.appendRow([kunci, hashBaru, "wali", kunci, "YA"]);
      return { success: "Akun wali untuk " + kunci + " berhasil dibuat (atlet lama belum punya akun).", username: kunci, passwordBaru: passwordBaru };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// ====================================================
// MODUL ATLET & PROFIL
// ====================================================

function getNamaAtletList(token, paksaSegar) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = "DAFTAR_ATLET_" + sesi.role + "_" + normalisasiNama(sesi.atletKey || "");
    // Bila paksaSegar true, lewati pembacaan cache dan baca langsung
    // dari Sheet; hasilnya tetap disimpan ke cache seperti biasa.
    if (!paksaSegar) {
      try {
        var cached = cache.get(cacheKey);
        if (cached) return JSON.parse(cached);
      } catch (eCache) {}
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    var sheetWaktu = ss.getSheetByName("LOG_WAKTU") || ss.getSheetByName("DATABASE_ACS");

    if (!sheetProfil) return [];
    var dataProfil = sheetProfil.getDataRange().getValues();
    if (dataProfil.length < 2) return [];

    var headersP = dataProfil[0];
    var idxPanggilan = getColumnIndex(headersP, ["NAMA_PANGGILAN", "PANGGILAN", "KEY"]);
    var idxLengkap = getColumnIndex(headersP, ["NAMA_LENGKAP", "NAMA LENGKAP", "NAMA SISWA", "NAMA"]);
    var idxGender = getColumnIndex(headersP, ["GENDER", "JENIS_KELAMIN", "JK"]);
    var idxTglLahir = getColumnIndex(headersP, ["TGL_LAHIR", "TANGGAL_LAHIR", "LAHIR", "TANGGAL LAHIR"]);
    var idxFoto = getColumnIndex(headersP, ["FOTO_BASE64", "FOTO", "AVATAR"]);
    var idxNoHp = getColumnIndex(headersP, ["NOMOR HP", "NOMOR_HP", "NO_HP", "NO HP", "NO_WA", "HP_WALI"]);

    if (idxPanggilan === -1) idxPanggilan = 0;
    if (idxLengkap === -1) idxLengkap = 1;
    if (idxGender === -1) idxGender = 2;
    if (idxTglLahir === -1) idxTglLahir = 3;
    if (idxFoto === -1) idxFoto = 6;
    if (idxNoHp === -1) idxNoHp = 7;

    var listAtlet = [];
    var lookupExact = {};
    var tz = ss.getSpreadsheetTimeZone();

    for (var i = 1; i < dataProfil.length; i++) {
      var row = dataProfil[i];
      var panggilan = String(row[idxPanggilan] || "").trim().toUpperCase();
      var lengkap = String(row[idxLengkap] || "").trim();
      var gender = String(row[idxGender] || "").trim();

      var rawTgl = row[idxTglLahir];
      var tglStr = (rawTgl instanceof Date) ? Utilities.formatDate(rawTgl, tz, "dd/MM/yyyy") : formatTanggalStr(rawTgl);

      // Wali hanya menerima anaknya sendiri. Pelatih menerima semua,
      // tetap tanpa foto base64 (foto hanya diambil lewat rapot
      // per atlet) tapi dengan noHp.
      var isAnakSendiri = (sesi.role === "wali") && (normalisasiNama(panggilan) === normalisasiNama(sesi.atletKey) || normalisasiNama(lengkap) === normalisasiNama(sesi.atletKey));
      if (sesi.role === "wali" && !isAnakSendiri) continue;

      var foto = isAnakSendiri ? normalisasiUrlFoto(String(row[idxFoto] || "").trim()) : "";
      var noHp = (idxNoHp !== -1) ? String(row[idxNoHp] || "").replace(/'/g, '').trim() : "";

      if (panggilan !== "") {
        var atletIdx = listAtlet.length;
        var atletObj = {
          panggilan: panggilan,
          lengkap: lengkap || panggilan,
          gender: gender,
          tglLahir: tglStr,
          foto: foto,
          noHp: noHp,
          allPbMap: {}
        };
        listAtlet.push(atletObj);

        var keyP = normalisasiNama(panggilan);
        var keyL = normalisasiNama(lengkap);
        lookupExact[keyP] = atletIdx;
        if (keyL !== "") lookupExact[keyL] = atletIdx;
      }
    }

    if (sheetWaktu) {
      var dataWaktu = sheetWaktu.getDataRange().getDisplayValues();
      if (dataWaktu.length > 1) {
        var headersW = dataWaktu[0];
        var idxWNama = getColumnIndex(headersW, ["NAMA_ATLET", "NAMA", "NAMA PANGGILAN", "ATLET"]);
        var idxWNomor = getColumnIndex(headersW, ["NOMOR_LOMBA", "NOMOR", "NOMOR LOMBA", "EVENT"]);
        var idxWWaktu = getColumnIndex(headersW, ["WAKTU", "CATATAN_WAKTU", "HASIL"]);

        if (idxWNama === -1) idxWNama = 0;
        if (idxWNomor === -1) idxWNomor = 1;
        if (idxWWaktu === -1) idxWWaktu = 3;

        for (var j = 1; j < dataWaktu.length; j++) {
          var rowW = dataWaktu[j];
          var rawNamaUpper = String(rowW[idxWNama] || "").trim().toUpperCase();
          var rawNomor = String(rowW[idxWNomor] || "").trim();
          var rawWaktu = String(rowW[idxWWaktu] || "").trim();

          if (rawNamaUpper !== "" && rawNomor !== "" && rawWaktu !== "") {
            var nomorKey = standarkanNomorLomba(rawNomor);
            var dtk = parseKeDetik(rawWaktu);
            if (dtk > 0 && lookupExact.hasOwnProperty(normalisasiNama(rawNamaUpper))) {
              var targetIdx = lookupExact[normalisasiNama(rawNamaUpper)];
              var targetPbMap = listAtlet[targetIdx].allPbMap;
              if (!targetPbMap[nomorKey] || dtk < targetPbMap[nomorKey].detik) {
                targetPbMap[nomorKey] = { waktuStr: rawWaktu, detik: dtk };
              }
            }
          }
        }
      }
    }

    try { cache.put(cacheKey, JSON.stringify(listAtlet), 300); } catch (ePut) {}
    return listAtlet;
  } catch (err) {
    return { error: err.toString() };
  }
}

// Peta foto atlet untuk pelatih: {PANGGILAN: "data:image/..."}.
// Dipisah dari getNamaAtletList agar foto base64 (besar) tidak masuk
// cache 100KB dan tidak memperlambat daftar atlet. Wali tidak perlu
// memanggil ini — fotonya sudah ada di getNamaAtletList (anak sendiri).
// Ubah isi kolom FOTO menjadi URL gambar yang bisa ditampilkan <img>.
// Menerima: link share Drive (file/d/ID/view, open?id=ID, uc?id=ID) ->
// thumbnail Drive; data-URL base64 atau URL gambar langsung diteruskan
// apa adanya. Syarat: file Drive harus dibagikan "Siapa saja yang
// memiliki link" — kalau tidak, thumbnail tidak akan tampil.
function normalisasiUrlFoto(nilai) {
  var s = String(nilai || "").trim();
  if (!s) return "";
  if (s.indexOf("data:") === 0) return s;
  var m = s.match(/drive\.google\.com\/file\/d\/([A-Za-z0-9_-]+)/)
    || s.match(/[?&]id=([A-Za-z0-9_-]{10,})/);
  if (m) return "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w400";
  return s;
}

function getFotoAtletMap(token) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    if (!sheetProfil) return {};
    var dataProfil = sheetProfil.getDataRange().getValues();
    if (dataProfil.length < 2) return {};
    var headersP = dataProfil[0];
    var idxPanggilan = getColumnIndex(headersP, ["NAMA_PANGGILAN", "PANGGILAN", "KEY"]);
    var idxFoto = getColumnIndex(headersP, ["FOTO_BASE64", "FOTO", "AVATAR"]);
    if (idxPanggilan === -1) idxPanggilan = 0;
    if (idxFoto === -1) idxFoto = 6;
    var peta = {};
    for (var i = 1; i < dataProfil.length; i++) {
      var panggilan = String(dataProfil[i][idxPanggilan] || "").trim().toUpperCase();
      var foto = String(dataProfil[i][idxFoto] || "").trim();
      if (panggilan && foto) peta[panggilan] = normalisasiUrlFoto(foto);
    }
    return peta;
  } catch (err) {
    return { error: err.toString() };
  }
}

// Internal: baca sheet PRESENSI_LATIHAN dan hitung statistik untuk
// satu atlet (pencocokan nama persis). Mengembalikan
// {totalSesi, hadir, izin, alpha, persen, riwayat10}.
function hitungPresensiAtlet(namaPanggilan, namaLengkap) {
  var hasil = { totalSesi: 0, hadir: 0, izin: 0, alpha: 0, persen: 0, riwayat10: [] };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PRESENSI_LATIHAN");
  if (!sheet) return hasil;
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return hasil;
  var headers = data[0];
  var idxTgl = getColumnIndex(headers, ["TANGGAL"]);
  var idxNama = getColumnIndex(headers, ["NAMA_ATLET", "NAMA"]);
  var idxStatus = getColumnIndex(headers, ["STATUS"]);
  var idxCatatan = getColumnIndex(headers, ["CATATAN"]);
  if (idxTgl === -1) idxTgl = 1;
  if (idxNama === -1) idxNama = 2;
  if (idxStatus === -1) idxStatus = 3;
  if (idxCatatan === -1) idxCatatan = 4;
  var kunci = normalisasiNama(namaPanggilan);
  var kunciLengkap = normalisasiNama(namaLengkap);
  var riwayat = [];
  for (var i = 1; i < data.length; i++) {
    var nm = normalisasiNama(data[i][idxNama]);
    if (nm !== kunci && nm !== kunciLengkap) continue;
    var st = String(data[i][idxStatus] || "").trim().toLowerCase();
    hasil.totalSesi++;
    if (st === "hadir") hasil.hadir++;
    else if (st === "izin") hasil.izin++;
    else if (st === "alpha") hasil.alpha++;
    riwayat.push({ rowIndex: i + 1, tanggal: formatTanggalStr(data[i][idxTgl]), namaAtlet: String(data[i][idxNama] || ""), status: String(data[i][idxStatus] || ""), catatan: String(data[i][idxCatatan] || "") });
  }
  // Persen 0 bila belum ada sesi (bukan 100).
  hasil.persen = hasil.totalSesi > 0 ? Math.round(hasil.hadir / hasil.totalSesi * 100) : 0;
  riwayat.reverse();
  hasil.riwayat10 = riwayat.slice(0, 10);
  return hasil;
}

// Internal: baca sheet PEMBAYARAN_SPP untuk satu atlet (pencocokan
// persis pada nama panggilan maupun nama lengkap).
function hitungSppAtlet(namaPanggilan, namaLengkap) {
  var hasil = { histori: [], statusBulanIni: "BELUM ADA" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PEMBAYARAN_SPP");
  if (!sheet) return hasil;
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return hasil;
  var headers = data[0];
  var idxPanggilan = getColumnIndex(headers, ["NAMA PANGGILAN", "NAMA_PANGGILAN"]);
  var idxLengkap = getColumnIndex(headers, ["NAMA LENGKAP", "NAMA LENGKAP"]);
  var idxBulan = getColumnIndex(headers, ["BULAN TAHUN", "BULAN_TAHUN", "BULAN"]);
  var idxNominal = getColumnIndex(headers, ["NOMINAL"]);
  var idxStatus = getColumnIndex(headers, ["STATUS"]);
  var idxCatatan = getColumnIndex(headers, ["CATATAN"]);
  if (idxPanggilan === -1) idxPanggilan = 1;
  if (idxLengkap === -1) idxLengkap = 2;
  if (idxBulan === -1) idxBulan = 3;
  if (idxNominal === -1) idxNominal = 4;
  if (idxStatus === -1) idxStatus = 5;
  if (idxCatatan === -1) idxCatatan = 6;
  var kunci = normalisasiNama(namaPanggilan);
  var kunciLengkap = normalisasiNama(namaLengkap);
  var histori = [];
  var sekarang = new Date();
  var bulanIni = namaBulanIndonesia(sekarang.getMonth()) + " " + sekarang.getFullYear();
  for (var i = 1; i < data.length; i++) {
    var p = normalisasiNama(data[i][idxPanggilan]);
    var l = normalisasiNama(data[i][idxLengkap]);
    if (p !== kunci && l !== kunciLengkap && p !== kunciLengkap && l !== kunci) continue;
    histori.push({
      rowIndex: i + 1,
      namaPanggilan: String(data[i][idxPanggilan] || ""),
      namaLengkap: String(data[i][idxLengkap] || ""),
      bulanTahun: formatBulanTahunIndonesia(data[i][idxBulan]),
      nominal: data[i][idxNominal],
      status: String(data[i][idxStatus] || ""),
      catatan: String(data[i][idxCatatan] || "")
    });
  }
  histori.reverse();
  hasil.histori = histori;
  for (var j = 0; j < histori.length; j++) {
    if (normalisasiNama(histori[j].bulanTahun) === normalisasiNama(bulanIni)) {
      hasil.statusBulanIni = histori[j].status || "BELUM ADA";
      break;
    }
  }
  return hasil;
}

function getAthleteData(token, namaPanggilan) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    var sheetWaktu = ss.getSheetByName("LOG_WAKTU") || ss.getSheetByName("DATABASE_ACS");
    var sheetKarakter = ss.getSheetByName("EVALUASI_KARAKTER") || ss.getSheetByName("KARAKTER");

    var keyNama = String(namaPanggilan || "").trim().toUpperCase();
    var keyNorm = normalisasiNama(namaPanggilan);

    var profilObj = { namaPanggilan: keyNama, namaLengkap: keyNama, gender: "-", tglLahir: "-", nisn: "-", nik: "-", fotoBase64: "", noHp: "" };

    if (sheetProfil) {
      var pData = sheetProfil.getDataRange().getValues();
      if (pData.length > 1) {
        var hP = pData[0].map(function(c) { return String(c || "").trim().toUpperCase(); });
        var idxP = getColumnIndex(hP, ["NAMA_PANGGILAN", "PANGGILAN", "KEY"]);
        var idxL = getColumnIndex(hP, ["NAMA_LENGKAP", "NAMA LENGKAP", "NAMA SISWA", "NAMA"]);
        var idxG = getColumnIndex(hP, ["GENDER", "JENIS_KELAMIN", "JK"]);
        var idxT = getColumnIndex(hP, ["TGL_LAHIR", "TANGGAL_LAHIR", "LAHIR", "TANGGAL LAHIR"]);
        var idxNs = getColumnIndex(hP, ["NISN"]);
        var idxNk = getColumnIndex(hP, ["NIK"]);
        var idxFt = getColumnIndex(hP, ["FOTO_BASE64", "FOTO", "AVATAR"]);
        var idxHp = getColumnIndex(hP, ["NOMOR HP", "NOMOR_HP", "NO_HP", "NO HP"]);

        if (idxP === -1) idxP = 0;
        if (idxL === -1) idxL = 1;
        if (idxG === -1) idxG = 2;
        if (idxT === -1) idxT = 3;
        if (idxNs === -1) idxNs = 4;
        if (idxNk === -1) idxNk = 5;
        if (idxFt === -1) idxFt = 6;
        if (idxHp === -1) idxHp = 7;

        var tz = ss.getSpreadsheetTimeZone();

        for (var i = 1; i < pData.length; i++) {
          var row = pData[i];
          var pVal = normalisasiNama(row[idxP]);
          if (pVal === keyNorm) {
            // Wali hanya boleh membuka rapot anaknya sendiri.
            if (!bolehAksesAtlet(sesi, row[idxP], row[idxL])) return { error: "AKSES_DITOLAK" };
            var rawTgl = row[idxT];
            var strTglFormatted = (rawTgl instanceof Date) ? Utilities.formatDate(rawTgl, tz, "dd/MM/yyyy") : formatTanggalStr(rawTgl);

            profilObj = {
              namaPanggilan: keyNama,
              namaLengkap: String(row[idxL] || keyNama).trim(),
              gender: String(row[idxG] || "-").trim(),
              tglLahir: strTglFormatted,
              nisn: (idxNs !== -1 && row[idxNs] !== undefined && row[idxNs] !== "") ? String(row[idxNs]).replace(/'/g, '').trim() : "-",
              nik: (idxNk !== -1 && row[idxNk] !== undefined && row[idxNk] !== "") ? String(row[idxNk]).replace(/'/g, '').trim() : "-",
              fotoBase64: (idxFt !== -1 && row[idxFt]) ? normalisasiUrlFoto(String(row[idxFt]).trim()) : "",
              noHp: (idxHp !== -1 && row[idxHp]) ? String(row[idxHp]).replace(/'/g, '').trim() : ""
            };
            break;
          }
        }
      }
    }

    // Bila profil tidak ditemukan dan pemanggil adalah wali, tolak
    // (wali tidak boleh menebak-nebak nama atlet lain).
    if (sesi.role === "wali" && !bolehAksesAtlet(sesi, namaPanggilan, profilObj.namaLengkap)) {
      return { error: "AKSES_DITOLAK" };
    }

    var riwayat = [];
    var bestTimes = {};
    if (sheetWaktu) {
      var wData = sheetWaktu.getDataRange().getDisplayValues();
      if (wData.length > 1) {
        var hW = wData[0];
        var idxWN = getColumnIndex(hW, ["NAMA_ATLET", "NAMA", "ATLET"]);
        var idxWNo = getColumnIndex(hW, ["NOMOR_LOMBA", "NOMOR", "EVENT"]);
        var idxWT = getColumnIndex(hW, ["TANGGAL", "TGL"]);
        var idxWW = getColumnIndex(hW, ["WAKTU", "CATATAN_WAKTU"]);
        var idxWS = getColumnIndex(hW, ["STATUS", "KATEGORI"]);

        if (idxWN === -1) idxWN = 0; if (idxWNo === -1) idxWNo = 1; if (idxWT === -1) idxWT = 2; if (idxWW === -1) idxWW = 3; if (idxWS === -1) idxWS = 4;

        for (var j = 1; j < wData.length; j++) {
          var wNama = normalisasiNama(wData[j][idxWN]);
          if (wNama === keyNorm) {
            var nomorStd = standarkanNomorLomba(String(wData[j][idxWNo] || "").trim());
            var tgl = formatTanggalStr(wData[j][idxWT]);
            var wkt = String(wData[j][idxWW] || "").trim();
            var st = String(wData[j][idxWS] || "Latihan").trim();

            var dtk = parseKeDetik(wkt);
            if (dtk > 0) {
              riwayat.push({ rowIndex: (j + 1), nomorLomba: nomorStd, tanggal: tgl, waktu: wkt, status: st });
              if (!bestTimes[nomorStd] || dtk < parseKeDetik(bestTimes[nomorStd])) {
                bestTimes[nomorStd] = wkt;
              }
            }
          }
        }
      }
    }

    var karakterObj = { disiplin: 80, fokus: 80, dayaJuang: 80, catatanPelatih: "Belum ada catatan." };
    if (sheetKarakter) {
      var kData = sheetKarakter.getDataRange().getDisplayValues();
      for (var k = kData.length - 1; k >= 1; k--) {
        var kNama = normalisasiNama(kData[k][0]);
        if (kNama === keyNorm) {
          karakterObj = {
            disiplin: Number(kData[k][1]) || 80,
            fokus: Number(kData[k][2]) || 80,
            dayaJuang: Number(kData[k][3]) || 80,
            catatanPelatih: String(kData[k][4] || "Belum ada catatan.")
          };
          break;
        }
      }
    }

    var presensi = hitungPresensiAtlet(profilObj.namaPanggilan, profilObj.namaLengkap);
    var spp = hitungSppAtlet(profilObj.namaPanggilan, profilObj.namaLengkap);

    return { profil: profilObj, riwayat: riwayat, bestTimes: bestTimes, karakter: karakterObj, presensi: presensi, spp: spp };
  } catch (err) {
    return { error: err.toString() };
  }
}

function simpanWaktuLatihan(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var detik = parseKeDetik(paket.waktu);
      var errValidasi = validasiWaktu(paket.nomorLomba, detik);
      if (errValidasi) return { error: errValidasi };

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("LOG_WAKTU") || ss.getSheetByName("DATABASE_ACS");
      if (!sheet) {
        sheet = ss.insertSheet("LOG_WAKTU");
        sheet.appendRow(["NAMA_ATLET", "NOMOR_LOMBA", "TANGGAL", "WAKTU", "STATUS"]);
      }

      if (paket.rowIndex && paket.rowIndex > 1) {
        sheet.getRange(paket.rowIndex, 1, 1, 5).setValues([[
          "'" + paket.namaAtlet.toUpperCase(),
          "'" + paket.nomorLomba,
          "'" + paket.tanggal,
          "'" + paket.waktu,
          "'" + paket.status
        ]]);
      } else {
        sheet.appendRow([
          "'" + paket.namaAtlet.toUpperCase(),
          "'" + paket.nomorLomba,
          "'" + paket.tanggal,
          "'" + paket.waktu,
          "'" + paket.status
        ]);
      }

      var waStatus = "Tidak dikirim (Checklist dimatikan)";
      if ((paket.kirimWa === true || paket.kirimWaAuto === true) && paket.noHpWali) {
        var pesanWA = "🏊‍♂️ *E-RAPOR DIGITAL ACS BANJARNEGARA*\n\n" +
                      "Halo Wali dari *" + paket.namaAtlet.toUpperCase() + "*,\n" +
                      "Berikut update rekor waktu terbaru:\n\n" +
                      "📌 *Nomor:* " + paket.nomorLomba + "\n" +
                      "⏱️ *Waktu:* " + paket.waktu + " detik\n" +
                      "📅 *Tanggal:* " + paket.tanggal + "\n" +
                      "🏷️ *Status:* " + paket.status + "\n\n" +
                      "DI AIR KITA JAYA! 💦";
        waStatus = kirimPesanFonnte(paket.noHpWali, pesanWA);
      }

      return { success: "Berhasil menyimpan catatan waktu!", waStatus: waStatus };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function hapusWaktuLatihan(token, rowIndex) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("LOG_WAKTU") || ss.getSheetByName("DATABASE_ACS");
      if (!sheet) return { error: "Sheet LOG_WAKTU tidak ditemukan!" };

      if (rowIndex && rowIndex > 1) {
        sheet.deleteRow(rowIndex);
        return { success: "Berhasil menghapus data rekor waktu!" };
      }
      return { error: "Baris data tidak valid!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function simpanWaktuLatihanMasal(token, dataArray) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("LOG_WAKTU") || ss.getSheetByName("DATABASE_ACS");
      if (!sheet) {
        sheet = ss.insertSheet("LOG_WAKTU");
        sheet.appendRow(["NAMA_ATLET", "NOMOR_LOMBA", "TANGGAL", "WAKTU", "STATUS"]);
      }

      // Daftar nama atlet yang sah (pencocokan persis).
      var namaSah = {};
      var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
      if (sheetProfil) {
        var pData = sheetProfil.getDataRange().getDisplayValues();
        for (var p = 1; p < pData.length; p++) {
          var nm = normalisasiNama(pData[p][0]);
          if (nm) namaSah[nm] = true;
        }
      }

      // Kunci duplikat yang sudah ada: nama+nomor+tanggal+waktu.
      var kunciAda = {};
      var dataLama = sheet.getDataRange().getDisplayValues();
      if (dataLama.length > 1) {
        for (var l = 1; l < dataLama.length; l++) {
          var k = normalisasiNama(dataLama[l][0]) + "|" + standarkanNomorLomba(dataLama[l][1]) + "|" + formatTanggalStr(dataLama[l][2]) + "|" + String(dataLama[l][3] || "").trim();
          kunciAda[k] = true;
        }
      }

      var rowsToAppend = [];
      var rincianDitolak = [];
      var ditolak = 0;
      for (var i = 0; i < dataArray.length; i++) {
        var item = dataArray[i];
        var namaNorm = normalisasiNama(item.namaAtlet);
        var nomorStd = standarkanNomorLomba(item.nomorLomba);
        var detik = parseKeDetik(item.waktu);
        var alasan = null;
        if (!namaSah[namaNorm]) {
          alasan = "Nama atlet tidak ditemukan di PROFIL_ATLET: " + item.namaAtlet;
        } else {
          var errW = validasiWaktu(item.nomorLomba, detik);
          if (errW) alasan = errW;
          else {
            var kunciBaru = namaNorm + "|" + nomorStd + "|" + formatTanggalStr(item.tanggal) + "|" + String(item.waktu || "").trim();
            if (kunciAda[kunciBaru]) alasan = "Duplikat: " + item.namaAtlet + " " + nomorStd + " " + item.tanggal + " " + item.waktu + " sudah ada.";
            else kunciAda[kunciBaru] = true;
          }
        }
        if (alasan) {
          ditolak++;
          rincianDitolak.push({ baris: i + 1, namaAtlet: item.namaAtlet, nomorLomba: item.nomorLomba, alasan: alasan });
        } else {
          rowsToAppend.push([
            "'" + String(item.namaAtlet).toUpperCase(),
            "'" + item.nomorLomba,
            "'" + item.tanggal,
            "'" + item.waktu,
            "'" + item.status
          ]);
        }
      }

      if (rowsToAppend.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
      }

      return { success: true, disimpan: rowsToAppend.length, ditolak: ditolak, rincianDitolak: rincianDitolak };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function simpanProfilAtlet(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
      if (!sheet) {
        sheet = ss.insertSheet("PROFIL_ATLET");
        sheet.appendRow(["NAMA_PANGGILAN", "NAMA_LENGKAP", "GENDER", "TGL_LAHIR", "NISN", "NIK", "FOTO_BASE64", "NOMOR HP"]);
      }

      var keyNama = String(paket.namaPanggilan || "").toUpperCase().trim();
      var keyNorm = normalisasiNama(paket.namaPanggilan);
      var data = sheet.getDataRange().getDisplayValues();
      var foundIndex = -1;

      for (var i = 1; i < data.length; i++) {
        var itemNorm = normalisasiNama(data[i][0]);
        if (itemNorm === keyNorm) {
          foundIndex = i + 1;
          break;
        }
      }

      var rowContent = [
        "'" + keyNama,
        "'" + paket.namaLengkap,
        "'" + paket.gender,
        "'" + paket.tglLahir,
        "'" + (paket.nisn ? String(paket.nisn).replace(/'/g, '').trim() : ""),
        "'" + (paket.nik ? String(paket.nik).replace(/'/g, '').trim() : ""),
        paket.fotoBase64 || "",
        "'" + (paket.noHp ? String(paket.noHp).replace(/'/g, '').trim() : "")
      ];

      if (foundIndex > 0) {
        sheet.getRange(foundIndex, 1, 1, rowContent.length).setValues([rowContent]);
      } else {
        sheet.appendRow(rowContent);
      }

      hapusCacheDaftarAtlet(keyNama);
      return { success: "Profil atlet " + keyNama + " berhasil disimpan!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function simpanCatatanKarakter(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("EVALUASI_KARAKTER") || ss.getSheetByName("KARAKTER");
      if (!sheet) {
        sheet = ss.insertSheet("EVALUASI_KARAKTER");
        sheet.appendRow(["NAMA_ATLET", "DISIPLIN", "FOKUS", "DAYA_JUANG", "CATATAN_PELATIH"]);
      }
      // Update baris yang sudah ada untuk atlet ini; bila belum ada,
      // baru append. Pencocokan nama persis.
      var keyNorm = normalisasiNama(paket.namaAtlet);
      var data = sheet.getDataRange().getDisplayValues();
      var foundIndex = -1;
      for (var i = 1; i < data.length; i++) {
        if (normalisasiNama(data[i][0]) === keyNorm) { foundIndex = i + 1; break; }
      }
      var rowContent = ["'" + String(paket.namaAtlet || "").toUpperCase(), paket.disiplin, paket.fokus, paket.dayaJuang, "'" + paket.catatanPelatih];
      if (foundIndex > 0) {
        sheet.getRange(foundIndex, 1, 1, rowContent.length).setValues([rowContent]);
      } else {
        sheet.appendRow(rowContent);
      }
      return { success: "Evaluasi karakter berhasil disimpan!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// ====================================================
// MODUL BUKU ACARA & PENDAFTARAN LOMBA (REKAP A3)
// ====================================================

function simpanAnalisisBukuAcara(token, eventName, dataJadwal) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("BUKU_ACARA_JADWAL");

      if (!sheet) {
        sheet = ss.insertSheet("BUKU_ACARA_JADWAL");
        sheet.appendRow(["TIMESTAMP", "NAMA_EVENT", "NO_ACARA", "NAMA_ACARA", "HEAT_SERI", "LANE_LINTASAN", "NAMA_ATLET", "KLUB", "ENTRY_TIME"]);
        sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#0f1a30").setFontColor("#00d4b2");
      }

      var timestamp = new Date();
      var rowsToAppend = [];

      for (var i = 0; i < dataJadwal.length; i++) {
        var item = dataJadwal[i];
        rowsToAppend.push([
          timestamp,
          "'" + eventName,
          "'" + item.noAcara,
          "'" + item.namaAcara,
          "'" + item.heat,
          "'" + item.lane,
          "'" + item.namaAtlet.toUpperCase(),
          "'" + item.klub,
          "'" + item.entryTime
        ]);
      }

      if (rowsToAppend.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 9).setValues(rowsToAppend);
      }

      return { success: true, count: rowsToAppend.length };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function getJadwalBukuAcara(token, eventName, namaAtletFilter) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("BUKU_ACARA_JADWAL");
    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    var result = [];
    var targetEv = String(eventName || "").trim().toLowerCase();
    // Wali hanya boleh melihat jadwal anaknya sendiri: filter
    // dipaksa ke atletKey wali, apa pun yang diminta pemanggil.
    var targetAtlet = sesi.role === "wali" ? normalisasiNama(sesi.atletKey) : normalisasiNama(namaAtletFilter);

    for (var i = 1; i < data.length; i++) {
      var ev = String(data[i][1] || "").trim().toLowerCase();
      var atl = normalisasiNama(data[i][6]);

      var matchEv = (!targetEv || targetEv === "all" || ev === targetEv);
      var matchAtl = (!targetAtlet || atl === targetAtlet);

      if (matchEv && matchAtl) {
        result.push({
          timestamp: data[i][0],
          eventName: data[i][1],
          noAcara: data[i][2],
          namaAcara: data[i][3],
          heat: data[i][4],
          lane: data[i][5],
          namaAtlet: data[i][6],
          klub: data[i][7],
          entryTime: data[i][8]
        });
      }
    }
    return result;
  } catch (err) {
    return { error: err.toString() };
  }
}

function getDaftarEventBukuAcara(token) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("BUKU_ACARA_JADWAL");
    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    var listEvent = [];
    for (var i = 1; i < data.length; i++) {
      var rawEv = String(data[i][1] || "").replace(/^'/, '').trim();
      if (rawEv !== "") {
        var isExist = listEvent.some(function(item) { return item.toUpperCase() === rawEv.toUpperCase(); });
        if (!isExist) listEvent.push(rawEv);
      }
    }
    return listEvent;
  } catch (err) {
    return { error: err.toString() };
  }
}

function hapusBukuAcara(token, eventName) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("BUKU_ACARA_JADWAL");
      if (!sheet) return { error: "Sheet BUKU_ACARA_JADWAL tidak ditemukan!" };

      var data = sheet.getDataRange().getDisplayValues();
      var targetEv = String(eventName || "").trim().toLowerCase();

      for (var i = data.length - 1; i >= 1; i--) {
        var ev = String(data[i][1] || "").trim().toLowerCase();
        if (targetEv === "all" || ev === targetEv) {
          sheet.deleteRow(i + 1);
        }
      }
      return { success: "Buku Acara berhasil dihapus!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// Sheet EVENT_LOMBA menyimpan metadata event yang dibuat pelatih
// (belum ada pendaftar): NAMA_EVENT, TANGGAL, LOKASI, KATEGORI,
// BIAYA_PER_NOMOR, DIBUAT_OLEH, TIMESTAMP.
function tambahEventLomba(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var namaEvent = String((paket && paket.namaEvent) || "").trim();
      if (!namaEvent) return { error: "Nama event wajib diisi." };
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("EVENT_LOMBA");
      if (!sheet) {
        sheet = ss.insertSheet("EVENT_LOMBA");
        sheet.appendRow(["NAMA_EVENT", "TANGGAL", "LOKASI", "KATEGORI", "BIAYA_PER_NOMOR", "DIBUAT_OLEH", "TIMESTAMP"]);
      }
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0] || "").trim().toLowerCase() === namaEvent.toLowerCase()) {
          return { error: "Event \"" + namaEvent + "\" sudah ada." };
        }
      }
      // Cek juga di PENDAFTARAN_LOMBA (event lama yang sudah ada pendaftar).
      var sheetP = ss.getSheetByName("PENDAFTARAN_LOMBA");
      if (sheetP) {
        var dataP = sheetP.getDataRange().getDisplayValues();
        for (var j = 1; j < dataP.length; j++) {
          if (String(dataP[j][0] || "").trim().toLowerCase() === namaEvent.toLowerCase()) {
            return { error: "Event \"" + namaEvent + "\" sudah ada (di pendaftaran lama)." };
          }
        }
      }
      var biaya = parseFloat(paket.biayaPerNomor) || 0;
      sheet.appendRow([namaEvent, String(paket.tanggal || ""), String(paket.lokasi || ""), String(paket.kategori || ""), biaya, sesi.username || sesi.atletKey || "pelatih", new Date()]);
      return { success: "Event \"" + namaEvent + "\" berhasil ditambahkan.", eventName: namaEvent };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function getDaftarEventLomba(token) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var listEvent = [];
    function tambahKeDaftar(nama) {
      var rawEv = String(nama || "").replace(/^'/, '').trim();
      if (rawEv === "") return;
      var isExist = listEvent.some(function(item) { return item.toUpperCase() === rawEv.toUpperCase(); });
      if (!isExist) listEvent.push(rawEv);
    }
    var sheet = ss.getSheetByName("PENDAFTARAN_LOMBA");
    if (sheet) {
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) tambahKeDaftar(data[i][0]);
    }
    // + event yang dibuat pelatih tapi belum ada pendaftar (sheet EVENT_LOMBA).
    var sheetEv = ss.getSheetByName("EVENT_LOMBA");
    if (sheetEv) {
      var dataEv = sheetEv.getDataRange().getDisplayValues();
      for (var k = 1; k < dataEv.length; k++) tambahKeDaftar(dataEv[k][0]);
    }
    return listEvent;
  } catch (err) {
    return { error: err.toString() };
  }
}

function simpanPendaftaranLomba(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("PENDAFTARAN_LOMBA");
      if (!sheet) {
        sheet = ss.insertSheet("PENDAFTARAN_LOMBA");
        sheet.appendRow(["NAMA_EVENT", "NAMA_ATLET", "TGL_LAHIR", "PA_PI", "NOMOR_DIIKUTI", "TOTAL_NOMOR", "BIAYA_PER_NOMOR", "TOTAL_BIAYA", "STATUS_BAYAR", "NO_HP_WALI", "TGL_INPUT"]);
      }

      var data = sheet.getDataRange().getDisplayValues();
      var targetEvent = String(paket.namaEvent).trim();
      var targetNamaKey = normalisasiNama(paket.namaLengkap || paket.namaAtlet);

      var inputNomorArr = Array.isArray(paket.nomorDiikuti) ? paket.nomorDiikuti : [paket.nomorDiikuti];

      var existingRowIdx = -1;
      var existingNomorArr = [];

      for (var i = 1; i < data.length; i++) {
        var ev = String(data[i][0] || "").trim();
        var nm = normalisasiNama(data[i][1]);

        if (ev.toLowerCase() === targetEvent.toLowerCase() && nm === targetNamaKey) {
          existingRowIdx = i + 1;
          var rawNomor = String(data[i][4] || "").trim();
          if (rawNomor) {
            existingNomorArr = rawNomor.split(",").map(function(s) { return s.trim(); });
          }
          break;
        }
      }

      var currentStatusBayar = paket.statusBayar;
      if (!currentStatusBayar) {
        if (existingRowIdx > 0 && data[existingRowIdx - 1][8]) {
          currentStatusBayar = data[existingRowIdx - 1][8];
        } else {
          currentStatusBayar = "Belum Lunas";
        }
      }

      var nomorDuplikat = [];
      var nomorBaruValid = [];

      inputNomorArr.forEach(function(nomor) {
        var clean = String(nomor).trim();
        if (clean !== "") {
          var isAlreadyRegistered = existingNomorArr.some(function(ex) {
            return ex.toLowerCase() === clean.toLowerCase();
          });

          if (isAlreadyRegistered) {
            nomorDuplikat.push(clean);
          } else {
            nomorBaruValid.push(clean);
          }
        }
      });

      if (nomorDuplikat.length > 0 && nomorBaruValid.length === 0 && !paket.isEditMode) {
        return {
          error: "Pendaftaran Ditolak! Atlet " + targetNamaKey + " sudah terdaftar di nomor: " + nomorDuplikat.join(", ") + " pada event " + targetEvent + "."
        };
      }

      var finalNomorSet = paket.isEditMode ? inputNomorArr : existingNomorArr.concat(nomorBaruValid);
      var nomorStrFinal = finalNomorSet.join(", ");
      var hitungTotalNomor = finalNomorSet.length;
      var biayaSatuan = parseInt(paket.biayaPerNomor) || 0;
      var totalBiayaFinal = hitungTotalNomor * biayaSatuan;

      var rowContent = [
        "'" + targetEvent,
        "'" + (paket.namaLengkap || paket.namaAtlet),
        "'" + paket.tglLahir,
        "'" + paket.paPi,
        "'" + nomorStrFinal,
        hitungTotalNomor,
        biayaSatuan,
        totalBiayaFinal,
        "'" + currentStatusBayar,
        "'" + (paket.noHpWali || ""),
        "'" + paket.tanggalInput
      ];

      if (existingRowIdx > 0) {
        sheet.getRange(existingRowIdx, 1, 1, rowContent.length).setValues([rowContent]);
      } else {
        sheet.appendRow(rowContent);
      }

      var waStatus = "Tidak dikirim (Checklist dimatikan)";
      var isKirimWa = (paket.kirimWaAuto === true || paket.kirimWa === true);

      if (isKirimWa && paket.noHpWali) {
        var resWa = kirimNotifikasiWaTagihanFonnte(token, {
          noHp: paket.noHpWali,
          namaAtlet: (paket.namaLengkap || paket.namaAtlet),
          namaEvent: targetEvent,
          nomorDiikuti: nomorStrFinal,
          totalBiaya: totalBiayaFinal,
          statusBayar: currentStatusBayar
        });
        waStatus = resWa.success ? resWa.success : resWa.error;
      }

      var msgSuccess = "Pendaftaran berhasil disimpan!";
      if (nomorDuplikat.length > 0 && !paket.isEditMode) {
        msgSuccess += " (Catatan: Nomor " + nomorDuplikat.join(", ") + " diabaikan karena sudah pernah terdaftar).";
      }

      return { success: msgSuccess, waStatus: waStatus, eventNameSaved: targetEvent, biayaSaved: biayaSatuan };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function hapusPendaftaranAtlet(token, namaAtletKey, namaEvent) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("PENDAFTARAN_LOMBA");
      if (!sheet) return { error: "Sheet PENDAFTARAN_LOMBA tidak ditemukan." };

      var data = sheet.getDataRange().getValues();
      var targetEv = normalisasiNama(namaEvent);
      var targetAt = normalisasiNama(namaAtletKey);
      var deletedCount = 0;

      for (var i = data.length - 1; i >= 1; i--) {
        var ev = normalisasiNama(data[i][0]);
        var at = normalisasiNama(data[i][1]);

        // Pencocokan PERSIS (setelah normalisasi) — tanpa substring.
        var isEvMatch = (targetEv === "" || targetEv === "ALL" || ev === targetEv);
        var isAtMatch = (at === targetAt);

        if (isEvMatch && isAtMatch) {
          sheet.deleteRow(i + 1);
          deletedCount++;
        }
      }

      if (deletedCount > 0) {
        return { success: "Pendaftaran atlet " + namaAtletKey + " berhasil dihapus dari lomba." };
      } else {
        return { error: "Data pendaftaran atlet tidak ditemukan di spreadsheet." };
      }
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function getRekapPendaftaranLomba(token, filterEvent) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPendaftaran = ss.getSheetByName("PENDAFTARAN_LOMBA");
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");

    if (!sheetPendaftaran) return [];

    var mapNamaLengkap = {};
    var namaLengkapAnakWali = "";
    if (sheetProfil) {
      var pData = sheetProfil.getDataRange().getDisplayValues();
      if (pData.length > 1) {
        var hP = pData[0];
        var idxP = getColumnIndex(hP, ["NAMA_PANGGILAN", "PANGGILAN", "KEY"]);
        var idxL = getColumnIndex(hP, ["NAMA_LENGKAP", "NAMA LENGKAP", "NAMA"]);
        if (idxP === -1) idxP = 0; if (idxL === -1) idxL = 1;

        for (var p = 1; p < pData.length; p++) {
          var pKey = String(pData[p][idxP] || "").toUpperCase().trim();
          var pLkp = String(pData[p][idxL] || "").trim();
          if (pKey) mapNamaLengkap[pKey] = pLkp || pKey;
          if (sesi.role === "wali" && normalisasiNama(pData[p][idxP]) === normalisasiNama(sesi.atletKey)) {
            namaLengkapAnakWali = pLkp;
          }
        }
      }
    }

    var data = sheetPendaftaran.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    var consolidatedMap = {};

    for (var i = 1; i < data.length; i++) {
      var namaEvent = String(data[i][0] || "").trim();

      if (filterEvent && filterEvent !== "all" && namaEvent.toLowerCase() !== String(filterEvent).trim().toLowerCase()) {
        continue;
      }

      var rawNama = String(data[i][1] || "").trim();
      var rawNamaKey = rawNama.toUpperCase();
      var namaTampil = mapNamaLengkap[rawNamaKey] ? mapNamaLengkap[rawNamaKey] : rawNama;

      // Wali hanya melihat baris anaknya sendiri.
      if (sesi.role === "wali") {
        var cocokAnak = (normalisasiNama(rawNama) === normalisasiNama(sesi.atletKey)) ||
          (normalisasiNama(namaTampil) === normalisasiNama(sesi.atletKey)) ||
          (namaLengkapAnakWali && normalisasiNama(rawNama) === normalisasiNama(namaLengkapAnakWali));
        if (!cocokAnak) continue;
      }

      var comboKey = namaEvent.toUpperCase() + "_" + namaTampil.toUpperCase();
      var rawNomorStr = String(data[i][4] || "").trim();
      var arrNomor = rawNomorStr ? rawNomorStr.split(",").map(function(s){ return s.trim(); }) : [];

      if (!consolidatedMap[comboKey]) {
        consolidatedMap[comboKey] = {
          namaEvent: namaEvent,
          namaAtlet: namaTampil,
          tglLahir: data[i][2] || "-",
          paPi: data[i][3] || "-",
          nomorSet: arrNomor,
          biayaPerNomor: parseInt(data[i][6]) || 50000,
          statusBayar: data[i][8] || "Belum Lunas",
          noHpWali: data[i][9] || "-",
          tanggalInput: data[i][10] || "-"
        };
      } else {
        arrNomor.forEach(function(nomor) {
          if (nomor && consolidatedMap[comboKey].nomorSet.indexOf(nomor) === -1) {
            consolidatedMap[comboKey].nomorSet.push(nomor);
          }
        });
      }
    }

    var result = [];
    Object.keys(consolidatedMap).forEach(function(key) {
      var item = consolidatedMap[key];
      var totalNomor = item.nomorSet.length;
      var totalBiaya = totalNomor * item.biayaPerNomor;

      result.push({
        namaEvent: item.namaEvent,
        namaAtlet: item.namaAtlet,
        tglLahir: item.tglLahir,
        paPi: item.paPi,
        nomorDiikuti: item.nomorSet.join(", "),
        totalNomor: totalNomor,
        biayaPerNomor: item.biayaPerNomor,
        totalBiaya: totalBiaya,
        statusBayar: item.statusBayar,
        noHpWali: item.noHpWali,
        tanggalInput: item.tanggalInput
      });
    });

    return result;
  } catch (err) {
    return { error: err.toString() };
  }
}

function updateInformasiEventA3(token, oldName, newName, newBiaya) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("PENDAFTARAN_LOMBA");
      if (!sheet) return { error: "Sheet PENDAFTARAN_LOMBA tidak ditemukan!" };

      var data = sheet.getDataRange().getDisplayValues();
      var targetOld = String(oldName).trim().toLowerCase();

      for (var i = 1; i < data.length; i++) {
        var ev = String(data[i][0] || "").trim().toLowerCase();
        if (ev === targetOld) {
          sheet.getRange(i + 1, 1).setValue("'" + newName);
          if (newBiaya && newBiaya > 0) {
            var totalNomor = parseInt(data[i][5]) || 0;
            var newTotalBiaya = totalNomor * newBiaya;
            sheet.getRange(i + 1, 7).setValue(newBiaya);
            sheet.getRange(i + 1, 8).setValue(newTotalBiaya);
          }
        }
      }
      return { success: "Informasi Event berhasil diperbarui!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// ALIAS FUNGSI UNTUK PEMANGGILAN DARI INDEX.HTML
function editEventA3Header(token, oldName, newName, newBiaya) {
  return updateInformasiEventA3(token, oldName, newName, newBiaya);
}

function hapusEventA3(token, eventName) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("PENDAFTARAN_LOMBA");
      if (!sheet) return { error: "Sheet PENDAFTARAN_LOMBA tidak ditemukan!" };

      var data = sheet.getDataRange().getDisplayValues();
      var targetEv = String(eventName).trim().toLowerCase();

      for (var i = data.length - 1; i >= 1; i--) {
        var ev = String(data[i][0] || "").trim().toLowerCase();
        if (ev === targetEv) {
          sheet.deleteRow(i + 1);
        }
      }
      // Bersihkan juga metadata event (sheet EVENT_LOMBA) bila ada.
      var sheetEv = ss.getSheetByName("EVENT_LOMBA");
      if (sheetEv) {
        var dataEv = sheetEv.getDataRange().getDisplayValues();
        for (var k = dataEv.length - 1; k >= 1; k--) {
          if (String(dataEv[k][0] || "").trim().toLowerCase() === targetEv) sheetEv.deleteRow(k + 1);
        }
      }
      return { success: "Event " + eventName + " beserta pendaftarannya berhasil dihapus!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// ====================================================
// MODUL PRESENSI LATIHAN (sheet PRESENSI_LATIHAN)
// Kolom: TIMESTAMP, TANGGAL, NAMA_ATLET, STATUS, CATATAN
// ====================================================

function getPresensiSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PRESENSI_LATIHAN");
  if (!sheet) {
    sheet = ss.insertSheet("PRESENSI_LATIHAN");
    sheet.appendRow(["TIMESTAMP", "TANGGAL", "NAMA_ATLET", "STATUS", "CATATAN"]);
  }
  return sheet;
}

function simpanPresensiSatuAtlet(token, namaAtlet, status, catatan) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var sheet = getPresensiSheet();
      var hariIni = formatTanggalStr(new Date());
      var kunciNama = normalisasiNama(namaAtlet);
      if (!kunciNama) return { error: "Nama atlet wajib diisi." };
      var st = String(status || "Hadir").trim();

      var data = sheet.getDataRange().getDisplayValues();
      if (data.length > 1) {
        var headers = data[0];
        var idxTgl = getColumnIndex(headers, ["TANGGAL"]);
        var idxNama = getColumnIndex(headers, ["NAMA_ATLET", "NAMA"]);
        if (idxTgl === -1) idxTgl = 1;
        if (idxNama === -1) idxNama = 2;
        for (var i = 1; i < data.length; i++) {
          if (normalisasiNama(data[i][idxNama]) === kunciNama && formatTanggalStr(data[i][idxTgl]) === hariIni) {
            return { error: "Atlet sudah presensi hari ini." };
          }
        }
      }

      sheet.appendRow([new Date(), hariIni, "'" + String(namaAtlet).toUpperCase(), "'" + st, "'" + (catatan || "")]);
      return { success: "Presensi " + String(namaAtlet).toUpperCase() + " (" + st + ") berhasil disimpan!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function hapusPresensi(token, rowIndex) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var sheet = getPresensiSheet();
      var baris = Number(rowIndex);
      if (!(baris > 1) || baris % 1 !== 0 || baris > sheet.getLastRow()) {
        return { error: "Baris presensi tidak valid." };
      }
      sheet.deleteRow(baris);
      return { success: "Presensi berhasil dihapus." };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function getStatistikPresensiAtlet(token, namaPanggilan) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    // Cari nama lengkap atlet untuk pencocokan persis dua arah.
    var namaLengkap = "";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    if (sheetProfil) {
      var pData = sheetProfil.getDataRange().getDisplayValues();
      for (var p = 1; p < pData.length; p++) {
        if (normalisasiNama(pData[p][0]) === normalisasiNama(namaPanggilan)) { namaLengkap = String(pData[p][1] || ""); break; }
        if (normalisasiNama(pData[p][1]) === normalisasiNama(namaPanggilan)) { namaLengkap = String(pData[p][1] || ""); break; }
      }
    }
    if (!bolehAksesAtlet(sesi, namaPanggilan, namaLengkap)) return { error: "AKSES_DITOLAK" };
    var hasil = hitungPresensiAtlet(namaPanggilan, namaLengkap);
    return {
      namaAtlet: String(namaPanggilan || "").toUpperCase(),
      totalSesi: hasil.totalSesi,
      hadir: hasil.hadir,
      izin: hasil.izin,
      alpha: hasil.alpha,
      persen: hasil.persen,
      riwayat: hasil.riwayat10,
      riwayat10: hasil.riwayat10
    };
  } catch (err) {
    return { error: err.toString() };
  }
}

function getRekapPresensiAnalisis(token, bulan) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PRESENSI_LATIHAN");
    var target = parseBulanTahun(bulan);
    // Bila parameter bulan tidak bisa diparse, pakai bulan berjalan.
    var bulanTarget = target ? target.bulanIndex : new Date().getMonth();
    var tahunTarget = target && target.tahun ? target.tahun : new Date().getFullYear();

    // Mulai dari daftar PROFIL_ATLET agar atlet tanpa sesi pun
    // tampil dengan persen 0 (Perlu Perhatian).
    var agregasi = {};
    var urutan = [];
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    if (sheetProfil) {
      var pData = sheetProfil.getDataRange().getDisplayValues();
      for (var p = 1; p < pData.length; p++) {
        var nm = String(pData[p][0] || "").trim().toUpperCase();
        if (!nm) continue;
        var kunci = normalisasiNama(nm);
        if (!agregasi[kunci]) {
          agregasi[kunci] = { namaAtlet: nm, totalSesi: 0, hadir: 0, izin: 0, alpha: 0, persen: 0, kategori: "Perlu Perhatian" };
          urutan.push(kunci);
        }
      }
    }

    if (sheet) {
      var data = sheet.getDataRange().getDisplayValues();
      if (data.length > 1) {
        var headers = data[0];
        var idxTgl = getColumnIndex(headers, ["TANGGAL"]);
        var idxNama = getColumnIndex(headers, ["NAMA_ATLET", "NAMA"]);
        var idxStatus = getColumnIndex(headers, ["STATUS"]);
        if (idxTgl === -1) idxTgl = 1;
        if (idxNama === -1) idxNama = 2;
        if (idxStatus === -1) idxStatus = 3;
        for (var i = 1; i < data.length; i++) {
          var tglVal = data[i][idxTgl];
          var d = (tglVal instanceof Date) ? tglVal : new Date(String(tglVal || ""));
          // Fallback: parse dd/MM/yyyy secara manual bila Date gagal.
          if (isNaN(d.getTime()) && /^\d{2}\/\d{2}\/\d{4}$/.test(String(tglVal || "").trim())) {
            var pp = String(tglVal).split("/");
            d = new Date(parseInt(pp[2], 10), parseInt(pp[1], 10) - 1, parseInt(pp[0], 10));
          }
          if (isNaN(d.getTime())) continue;
          if (d.getMonth() !== bulanTarget || d.getFullYear() !== tahunTarget) continue;
          var kunciNama = normalisasiNama(data[i][idxNama]);
          if (!kunciNama) continue;
          if (!agregasi[kunciNama]) {
            agregasi[kunciNama] = { namaAtlet: String(data[i][idxNama] || "").trim().toUpperCase(), totalSesi: 0, hadir: 0, izin: 0, alpha: 0, persen: 0, kategori: "Perlu Perhatian" };
            urutan.push(kunciNama);
          }
          var st = String(data[i][idxStatus] || "").trim().toLowerCase();
          agregasi[kunciNama].totalSesi++;
          if (st === "hadir") agregasi[kunciNama].hadir++;
          else if (st === "izin") agregasi[kunciNama].izin++;
          else if (st === "alpha") agregasi[kunciNama].alpha++;
        }
      }
    }

    var hasil = [];
    for (var u = 0; u < urutan.length; u++) {
      var item = agregasi[urutan[u]];
      item.persen = item.totalSesi > 0 ? Math.round(item.hadir / item.totalSesi * 100) : 0;
      if (item.persen >= 70) item.kategori = "Sangat Disiplin";
      else if (item.persen >= 50) item.kategori = "Cukup Aktif";
      else item.kategori = "Perlu Perhatian";
      hasil.push(item);
    }
    hasil.sort(function(a, b) { return b.persen - a.persen; });
    return hasil;
  } catch (err) {
    return { error: err.toString() };
  }
}

function kirimIzinLatihanWali(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["wali"]);
  } catch (e) {
    return { error: e.message };
  }
  if (!paket || !paket.namaAtlet) return { error: "Nama atlet wajib diisi." };
  if (!bolehAksesAtlet(sesi, paket.namaAtlet, "")) return { error: "AKSES_DITOLAK" };
  return denganLock(function() {
    try {
      var sheet = getPresensiSheet();
      var tanggalStr = tanggalYmdKeDmy(paket.tanggal || new Date());
      var alasan = String(paket.alasan || "").trim();
      sheet.appendRow([new Date(), tanggalStr, "'" + String(paket.namaAtlet).toUpperCase(), "'Izin", "'Izin Wali: " + alasan]);
      return { success: "Surat izin latihan berhasil dikirim!" };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function getHistoryPresensiSpesifikAtlet(token, namaPanggilan) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var namaLengkap = "";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    if (sheetProfil) {
      var pData = sheetProfil.getDataRange().getDisplayValues();
      for (var p = 1; p < pData.length; p++) {
        if (normalisasiNama(pData[p][0]) === normalisasiNama(namaPanggilan)) { namaLengkap = String(pData[p][1] || ""); break; }
        if (normalisasiNama(pData[p][1]) === normalisasiNama(namaPanggilan)) { namaLengkap = String(pData[p][1] || ""); break; }
      }
    }
    if (!bolehAksesAtlet(sesi, namaPanggilan, namaLengkap)) return { error: "AKSES_DITOLAK" };
    var hasil = hitungPresensiAtlet(namaPanggilan, namaLengkap);
    return {
      namaAtlet: String(namaPanggilan || "").toUpperCase(),
      hadir: hasil.hadir,
      izin: hasil.izin,
      alpha: hasil.alpha,
      total: hasil.totalSesi,
      totalSesi: hasil.totalSesi,
      persen: hasil.persen,
      riwayat: hasil.riwayat10,
      riwayat10: hasil.riwayat10
    };
  } catch (err) {
    return { error: err.toString() };
  }
}

function getDaftarIzinLatihan(token) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PRESENSI_LATIHAN");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    var headers = data[0];
    var idxTgl = getColumnIndex(headers, ["TANGGAL"]);
    var idxNama = getColumnIndex(headers, ["NAMA_ATLET", "NAMA"]);
    var idxStatus = getColumnIndex(headers, ["STATUS"]);
    var idxCatatan = getColumnIndex(headers, ["CATATAN"]);
    if (idxTgl === -1) idxTgl = 1;
    if (idxNama === -1) idxNama = 2;
    if (idxStatus === -1) idxStatus = 3;
    if (idxCatatan === -1) idxCatatan = 4;
    var hasil = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxStatus] || "").trim().toLowerCase() !== "izin") continue;
      hasil.push({
        rowIndex: i + 1,
        tanggal: formatTanggalStr(data[i][idxTgl]),
        namaAtlet: String(data[i][idxNama] || ""),
        status: String(data[i][idxStatus] || ""),
        catatan: String(data[i][idxCatatan] || "")
      });
    }
    hasil.reverse();
    return hasil;
  } catch (err) {
    return { error: err.toString() };
  }
}

// ====================================================
// MODUL SPP / IURAN (sheet PEMBAYARAN_SPP)
// Kolom: TIMESTAMP, NAMA PANGGILAN, NAMA LENGKAP,
//        BULAN TAHUN, NOMINAL, STATUS, CATATAN, FOTO BUKTI
// ====================================================

function getSppSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PEMBAYARAN_SPP");
  if (!sheet) {
    sheet = ss.insertSheet("PEMBAYARAN_SPP");
    sheet.appendRow(["TIMESTAMP", "NAMA PANGGILAN", "NAMA LENGKAP", "BULAN TAHUN", "NOMINAL", "STATUS", "CATATAN", "FOTO BUKTI"]);
  }
  return sheet;
}

function submitKonfirmasiSppWali(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["wali"]);
  } catch (e) {
    return { error: e.message };
  }
  if (!paket || !paket.namaPanggilan) return { error: "Nama panggilan wajib diisi." };
  if (!bolehAksesAtlet(sesi, paket.namaPanggilan, paket.namaLengkap)) return { error: "AKSES_DITOLAK" };
  return denganLock(function() {
    try {
      var sheet = getSppSheet();
      // Anti-duplikat: bila sudah ada baris dengan nama panggilan,
      // bulan, dan nominal yang sama berstatus PENDING VERIFIKASI,
      // jangan tambah baris baru (tombol Kirim ditekan dua kali
      // atau retry setelah timeout).
      var dataLama = sheet.getDataRange().getDisplayValues();
      if (dataLama.length > 1) {
        var headersLama = dataLama[0];
        var idxPanggilan = getColumnIndex(headersLama, ["NAMA PANGGILAN", "NAMA_PANGGILAN"]);
        var idxBulan = getColumnIndex(headersLama, ["BULAN TAHUN", "BULAN_TAHUN", "BULAN"]);
        var idxNominal = getColumnIndex(headersLama, ["NOMINAL"]);
        var idxStatus = getColumnIndex(headersLama, ["STATUS"]);
        if (idxPanggilan === -1) idxPanggilan = 1;
        if (idxBulan === -1) idxBulan = 3;
        if (idxNominal === -1) idxNominal = 4;
        if (idxStatus === -1) idxStatus = 5;
        var kunciNama = normalisasiNama(paket.namaPanggilan);
        var bulanFmt = formatBulanTahunIndonesia(paket.bulanTahun);
        var nomBaruStr = String(paket.nominal || "").trim();
        var nomBaruNum = Number(paket.nominal);
        for (var i = 1; i < dataLama.length; i++) {
          var samaNama = normalisasiNama(dataLama[i][idxPanggilan]) === kunciNama;
          var samaBulan = normalisasiNama(formatBulanTahunIndonesia(dataLama[i][idxBulan])) === normalisasiNama(bulanFmt);
          var samaNominal = String(dataLama[i][idxNominal] || "").trim() === nomBaruStr ||
            (!isNaN(nomBaruNum) && nomBaruStr !== "" && Number(dataLama[i][idxNominal]) === nomBaruNum);
          var masihPending = normalisasiNama(dataLama[i][idxStatus]) === "PENDING VERIFIKASI";
          if (samaNama && samaBulan && samaNominal && masihPending) {
            return { success: "Konfirmasi SPP sudah terkirim sebelumnya (PENDING VERIFIKASI)." };
          }
        }
      }
      sheet.appendRow([
        new Date(),
        "'" + String(paket.namaPanggilan || "").toUpperCase(),
        "'" + String(paket.namaLengkap || ""),
        "'" + formatBulanTahunIndonesia(paket.bulanTahun),
        paket.nominal,
        "'PENDING VERIFIKASI",
        "'" + (paket.catatan || ""),
        paket.fotoBukti || ""
      ]);
      return { success: "Konfirmasi SPP berhasil dikirim, menunggu verifikasi pelatih." };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function prosesVerifikasiSppCoach(token, rowIndex, statusBaru) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var sheet = getSppSheet();
      if (!rowIndex || rowIndex <= 1) return { error: "Baris data tidak valid!" };
      if (rowIndex > sheet.getLastRow()) return { error: "Baris data tidak ditemukan!" };
      var data = sheet.getDataRange().getDisplayValues();
      var headers = data[0];
      var idxStatus = getColumnIndex(headers, ["STATUS"]);
      if (idxStatus === -1) idxStatus = 5;
      sheet.getRange(rowIndex, idxStatus + 1).setValue("'" + String(statusBaru || "").trim());
      return { success: "Status SPP berhasil diubah menjadi " + String(statusBaru || "").trim() + "." };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function hapusSpp(token, rowIndex) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var sheet = getSppSheet();
      var baris = Number(rowIndex);
      if (!(baris > 1) || baris % 1 !== 0 || baris > sheet.getLastRow()) {
        return { error: "Baris SPP tidak valid." };
      }
      sheet.deleteRow(baris);
      return { success: "Konfirmasi SPP berhasil dihapus." };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

function getDaftarKonfirmasiSppCoach(token, bulanTahun) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PEMBAYARAN_SPP");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    var headers = data[0];
    var idxPanggilan = getColumnIndex(headers, ["NAMA PANGGILAN", "NAMA_PANGGILAN"]);
    var idxLengkap = getColumnIndex(headers, ["NAMA LENGKAP", "NAMA LENGKAP"]);
    var idxBulan = getColumnIndex(headers, ["BULAN TAHUN", "BULAN_TAHUN", "BULAN"]);
    var idxNominal = getColumnIndex(headers, ["NOMINAL"]);
    var idxStatus = getColumnIndex(headers, ["STATUS"]);
    var idxCatatan = getColumnIndex(headers, ["CATATAN"]);
    var idxFoto = getColumnIndex(headers, ["FOTO BUKTI", "FOTO_BUKTI"]);
    if (idxPanggilan === -1) idxPanggilan = 1;
    if (idxLengkap === -1) idxLengkap = 2;
    if (idxBulan === -1) idxBulan = 3;
    if (idxNominal === -1) idxNominal = 4;
    if (idxStatus === -1) idxStatus = 5;
    if (idxCatatan === -1) idxCatatan = 6;
    if (idxFoto === -1) idxFoto = 7;

    var filterBulan = bulanTahun ? formatBulanTahunIndonesia(bulanTahun) : "";
    var hasil = [];
    for (var i = 1; i < data.length; i++) {
      var bulanFmt = formatBulanTahunIndonesia(data[i][idxBulan]);
      if (filterBulan && normalisasiNama(bulanFmt) !== normalisasiNama(filterBulan)) continue;
      hasil.push({
        rowIndex: i + 1,
        timestamp: data[i][0],
        namaPanggilan: String(data[i][idxPanggilan] || ""),
        namaLengkap: String(data[i][idxLengkap] || ""),
        bulanTahun: bulanFmt,
        nominal: data[i][idxNominal],
        status: String(data[i][idxStatus] || ""),
        catatan: String(data[i][idxCatatan] || ""),
        fotoBukti: String(data[i][idxFoto] || "")
      });
    }
    hasil.reverse();
    return hasil;
  } catch (err) {
    return { error: err.toString() };
  }
}

function getHistoriSppSpesifikAtlet(token, nama) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var namaLengkap = "";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    if (sheetProfil) {
      var pData = sheetProfil.getDataRange().getDisplayValues();
      for (var p = 1; p < pData.length; p++) {
        if (normalisasiNama(pData[p][0]) === normalisasiNama(nama)) { namaLengkap = String(pData[p][1] || ""); break; }
        if (normalisasiNama(pData[p][1]) === normalisasiNama(nama)) { namaLengkap = String(pData[p][1] || ""); break; }
      }
    }
    if (!bolehAksesAtlet(sesi, nama, namaLengkap)) return { error: "AKSES_DITOLAK" };
    var sheet = ss.getSheetByName("PEMBAYARAN_SPP");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    var headers = data[0];
    var idxPanggilan = getColumnIndex(headers, ["NAMA PANGGILAN", "NAMA_PANGGILAN"]);
    var idxLengkap = getColumnIndex(headers, ["NAMA LENGKAP", "NAMA LENGKAP"]);
    var idxBulan = getColumnIndex(headers, ["BULAN TAHUN", "BULAN_TAHUN", "BULAN"]);
    var idxNominal = getColumnIndex(headers, ["NOMINAL"]);
    var idxStatus = getColumnIndex(headers, ["STATUS"]);
    var idxCatatan = getColumnIndex(headers, ["CATATAN"]);
    if (idxPanggilan === -1) idxPanggilan = 1;
    if (idxLengkap === -1) idxLengkap = 2;
    if (idxBulan === -1) idxBulan = 3;
    if (idxNominal === -1) idxNominal = 4;
    if (idxStatus === -1) idxStatus = 5;
    if (idxCatatan === -1) idxCatatan = 6;
    var kunci = normalisasiNama(nama);
    var kunciLengkap = normalisasiNama(namaLengkap);
    var hasil = [];
    for (var i = 1; i < data.length; i++) {
      var pN = normalisasiNama(data[i][idxPanggilan]);
      var lN = normalisasiNama(data[i][idxLengkap]);
      // Pencocokan PERSIS dari panggilan maupun lengkap.
      if (pN !== kunci && lN !== kunci && pN !== kunciLengkap && lN !== kunciLengkap) continue;
      hasil.push({
        rowIndex: i + 1,
        timestamp: data[i][0],
        namaPanggilan: String(data[i][idxPanggilan] || ""),
        namaLengkap: String(data[i][idxLengkap] || ""),
        bulanTahun: formatBulanTahunIndonesia(data[i][idxBulan]),
        nominal: data[i][idxNominal],
        status: String(data[i][idxStatus] || ""),
        catatan: String(data[i][idxCatatan] || "")
      });
    }
    hasil.reverse();
    return hasil;
  } catch (err) {
    return { error: err.toString() };
  }
}

function getRekapMatriksSppAdmin(token, tahun) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
    var sheet = ss.getSheetByName("PEMBAYARAN_SPP");
    var tahunTarget = parseInt(tahun, 10) || new Date().getFullYear();
    var namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    var daftar = [];
    var indeks = {};
    if (sheetProfil) {
      var pData = sheetProfil.getDataRange().getDisplayValues();
      for (var p = 1; p < pData.length; p++) {
        var panggilan = String(pData[p][0] || "").trim().toUpperCase();
        var lengkap = String(pData[p][1] || "").trim();
        if (!panggilan) continue;
        var matriks = {};
        for (var m = 0; m < 12; m++) matriks[namaBulan[m]] = "-";
        var baris = { namaPanggilan: panggilan, namaLengkap: lengkap || panggilan, matriks: matriks, totalLunas: 0 };
        for (var m2 = 0; m2 < 12; m2++) baris[namaBulan[m2]] = "-";
        indeks[normalisasiNama(panggilan)] = daftar.length;
        indeks[normalisasiNama(lengkap)] = daftar.length;
        daftar.push(baris);
      }
    }

    if (sheet) {
      var data = sheet.getDataRange().getDisplayValues();
      if (data.length > 1) {
        var headers = data[0];
        var idxPanggilan = getColumnIndex(headers, ["NAMA PANGGILAN", "NAMA_PANGGILAN"]);
        var idxLengkap = getColumnIndex(headers, ["NAMA LENGKAP", "NAMA LENGKAP"]);
        var idxBulan = getColumnIndex(headers, ["BULAN TAHUN", "BULAN_TAHUN", "BULAN"]);
        var idxStatus = getColumnIndex(headers, ["STATUS"]);
        if (idxPanggilan === -1) idxPanggilan = 1;
        if (idxLengkap === -1) idxLengkap = 2;
        if (idxBulan === -1) idxBulan = 3;
        if (idxStatus === -1) idxStatus = 5;
        for (var i = 1; i < data.length; i++) {
          var parsed = parseBulanTahun(data[i][idxBulan]);
          if (!parsed || parsed.tahun !== tahunTarget) continue;
          var kunci = normalisasiNama(data[i][idxPanggilan]);
          var kunciL = normalisasiNama(data[i][idxLengkap]);
          // Pencocokan PERSIS saja (tanpa fallback substring).
          var idx = indeks.hasOwnProperty(kunci) ? indeks[kunci] : (indeks.hasOwnProperty(kunciL) ? indeks[kunciL] : -1);
          if (idx === -1) continue;
          var bulanNama = namaBulan[parsed.bulanIndex];
          var statusBaru = String(data[i][idxStatus] || "").trim();
          if (!statusBaru) continue;
          var statusLama = daftar[idx].matriks[bulanNama];
          // Jangan timpa status LUNAS dengan baris duplikat.
          if (statusLama === "LUNAS" && statusBaru.toUpperCase() !== "LUNAS") continue;
          daftar[idx].matriks[bulanNama] = statusBaru;
          daftar[idx][bulanNama] = statusBaru;
        }
      }
    }

    for (var d = 0; d < daftar.length; d++) {
      var total = 0;
      for (var m3 = 0; m3 < 12; m3++) {
        if (String(daftar[d].matriks[namaBulan[m3]] || "").toUpperCase() === "LUNAS") total++;
      }
      daftar[d].totalLunas = total;
    }
    return daftar;
  } catch (err) {
    return { error: err.toString() };
  }
}

// ====================================================
// MODUL REGISTRASI ATLET BARU
// ====================================================

function simpanRegistrasiAtletBaru(token, paket) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
      if (!sheet) {
        sheet = ss.insertSheet("PROFIL_ATLET");
        sheet.appendRow(["NAMA_PANGGILAN", "NAMA_LENGKAP", "GENDER", "TGL_LAHIR", "NISN", "NIK", "FOTO_BASE64", "NOMOR HP"]);
      }

      var keyNama = String(paket.namaPanggilan || "").toUpperCase().trim();
      var keyNorm = normalisasiNama(paket.namaPanggilan);
      if (!keyNorm) return { error: "Nama panggilan wajib diisi." };
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) {
        if (normalisasiNama(data[i][0]) === keyNorm) {
          return { error: "Nama panggilan " + keyNama + " sudah terdaftar." };
        }
      }

      var rowContent = [
        "'" + keyNama,
        "'" + paket.namaLengkap,
        "'" + paket.gender,
        "'" + paket.tglLahir,
        "'" + (paket.nisn ? String(paket.nisn).replace(/'/g, '').trim() : ""),
        "'" + (paket.nik ? String(paket.nik).replace(/'/g, '').trim() : ""),
        paket.fotoBase64 || "",
        "'" + (paket.noHp ? String(paket.noHp).replace(/'/g, '').trim() : "")
      ];
      sheet.appendRow(rowContent);

      // Buat akun wali: password ACAK 8 karakter, simpan HANYA
      // hash-nya. Password asli hanya tampil sekali di hasil
      // fungsi ini dan tidak disimpan di mana pun.
      var passwordAcak = buatPasswordAcak();
      var userSheet = getUserAccountSheet();
      var uData = userSheet.getDataRange().getDisplayValues();
      var uHeaders = uData[0] || [];
      var idxUser = getColumnIndex(uHeaders, ["USERNAME"]);
      var idxHash = getColumnIndex(uHeaders, ["PASSWORD_HASH"]);
      var idxRole = getColumnIndex(uHeaders, ["ROLE"]);
      var idxAtlet = getColumnIndex(uHeaders, ["ATLET_KEY"]);
      var idxGanti = getColumnIndex(uHeaders, ["HARUS_GANTI_PASSWORD"]);
      if (idxUser === -1) idxUser = 0;
      if (idxHash === -1) idxHash = 1;
      if (idxRole === -1) idxRole = 2;
      if (idxAtlet === -1) idxAtlet = 3;
      if (idxGanti === -1) idxGanti = 4;
      var sudahAda = false;
      for (var u = 1; u < uData.length; u++) {
        if (String(uData[u][idxUser] || "").trim().toUpperCase() === keyNama) { sudahAda = true; break; }
      }
      if (!sudahAda) {
        userSheet.appendRow([keyNama, hitungHashSha256(passwordAcak), "wali", keyNama, "YA"]);
      }

      hapusCacheDaftarAtlet(keyNama);
      return { success: "Atlet baru " + keyNama + " berhasil didaftarkan dan akun wali dibuat.", username: keyNama, passwordSekali: passwordAcak };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// Internal: hapus semua baris di `sheet` yang nilai kolom
// `idxKolom`-nya cocok PERSIS (setelah normalisasi) dengan `kunci`
// atau `kunciLengkap`. Dihapus dari baris paling bawah ke atas
// supaya nomor baris tidak bergeser. Mengembalikan jumlah baris
// yang dihapus.
function hapusBarisCocokPersis(sheet, idxKolom, kunci, kunciLengkap) {
  var data = sheet.getDataRange().getDisplayValues();
  var dihapus = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    var nilai = normalisasiNama(data[i][idxKolom]);
    if (nilai === kunci || (kunciLengkap && nilai === kunciLengkap)) {
      sheet.deleteRow(i + 1);
      dihapus++;
    }
  }
  return dihapus;
}

function hapusAtlet(token, atletKey) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih"]);
  } catch (e) {
    return { error: e.message };
  }
  return denganLock(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var kunci = normalisasiNama(atletKey);
      if (!kunci) return { error: "Atlet tidak ditemukan." };
      var keyNama = String(atletKey || "").toUpperCase().trim();

      // 1. Cari dulu di PROFIL_ATLET (kolom NAMA_PANGGILAN).
      // Bila tidak ditemukan, JANGAN hapus apa pun.
      var sheetProfil = ss.getSheetByName("PROFIL_ATLET") || ss.getSheetByName("PROFIL");
      if (!sheetProfil) return { error: "Atlet tidak ditemukan." };
      var dataProfil = sheetProfil.getDataRange().getDisplayValues();
      if (dataProfil.length < 2) return { error: "Atlet tidak ditemukan." };
      var headersProfil = dataProfil[0];
      var idxPanggilan = getColumnIndex(headersProfil, ["NAMA_PANGGILAN", "PANGGILAN", "KEY"]);
      var idxLengkap = getColumnIndex(headersProfil, ["NAMA_LENGKAP", "NAMA LENGKAP", "NAMA SISWA", "NAMA"]);
      if (idxPanggilan === -1) idxPanggilan = 0;
      if (idxLengkap === -1) idxLengkap = 1;
      var namaLengkap = "";
      var ditemukan = false;
      for (var p = 1; p < dataProfil.length; p++) {
        if (normalisasiNama(dataProfil[p][idxPanggilan]) === kunci) {
          ditemukan = true;
          namaLengkap = String(dataProfil[p][idxLengkap] || "");
          break;
        }
      }
      if (!ditemukan) return { error: "Atlet tidak ditemukan." };
      // Nama lengkap dipakai sebagai kunci kedua di sheet lain:
      // beberapa sheet menyimpan nama lengkap, bukan panggilan.
      var kunciLengkap = normalisasiNama(namaLengkap);
      if (kunciLengkap === kunci) kunciLengkap = "";

      // 2. Hapus semua baris yang cocok PERSIS di tiap sheet,
      // dari baris paling bawah ke atas. Sheet yang tidak ada
      // dilewati tanpa error.
      var detail = { profil: 0, akunWali: 0, waktu: 0, presensi: 0, spp: 0, pendaftaran: 0 };

      detail.profil = hapusBarisCocokPersis(sheetProfil, idxPanggilan, kunci, "");

      var sheetAkun = getUserAccountSheet();
      var dataAkun = sheetAkun.getDataRange().getDisplayValues();
      if (dataAkun.length > 1) {
        var idxUser = getColumnIndex(dataAkun[0], ["USERNAME"]);
        if (idxUser === -1) idxUser = 0;
        detail.akunWali = hapusBarisCocokPersis(sheetAkun, idxUser, kunci, "");
      }

      var sheetWaktu = ss.getSheetByName("LOG_WAKTU") || ss.getSheetByName("DATABASE_ACS");
      if (sheetWaktu) {
        var dataWaktu = sheetWaktu.getDataRange().getDisplayValues();
        if (dataWaktu.length > 1) {
          var idxWNama = getColumnIndex(dataWaktu[0], ["NAMA_ATLET", "NAMA", "NAMA PANGGILAN", "ATLET"]);
          if (idxWNama === -1) idxWNama = 0;
          detail.waktu = hapusBarisCocokPersis(sheetWaktu, idxWNama, kunci, kunciLengkap);
        }
      }

      var sheetPresensi = ss.getSheetByName("PRESENSI_LATIHAN");
      if (sheetPresensi) {
        var dataPresensi = sheetPresensi.getDataRange().getDisplayValues();
        if (dataPresensi.length > 1) {
          var idxPNama = getColumnIndex(dataPresensi[0], ["NAMA_ATLET", "NAMA"]);
          if (idxPNama === -1) idxPNama = 2;
          detail.presensi = hapusBarisCocokPersis(sheetPresensi, idxPNama, kunci, kunciLengkap);
        }
      }

      var sheetSpp = ss.getSheetByName("PEMBAYARAN_SPP");
      if (sheetSpp) {
        var dataSpp = sheetSpp.getDataRange().getDisplayValues();
        if (dataSpp.length > 1) {
          var idxSppNama = getColumnIndex(dataSpp[0], ["NAMA PANGGILAN", "NAMA_PANGGILAN"]);
          if (idxSppNama === -1) idxSppNama = 1;
          detail.spp = hapusBarisCocokPersis(sheetSpp, idxSppNama, kunci, kunciLengkap);
        }
      }

      var sheetLomba = ss.getSheetByName("PENDAFTARAN_LOMBA");
      if (sheetLomba) {
        var dataLomba = sheetLomba.getDataRange().getDisplayValues();
        if (dataLomba.length > 1) {
          var idxLNama = getColumnIndex(dataLomba[0], ["NAMA_ATLET", "NAMA", "NAMA PANGGILAN"]);
          if (idxLNama === -1) idxLNama = 1;
          detail.pendaftaran = hapusBarisCocokPersis(sheetLomba, idxLNama, kunci, kunciLengkap);
        }
      }

      hapusCacheDaftarAtlet(keyNama);
      var total = detail.profil + detail.akunWali + detail.waktu + detail.presensi + detail.spp + detail.pendaftaran;
      return { success: "Atlet " + keyNama + " beserta " + total + " baris data terkait berhasil dihapus.", detail: detail };
    } catch (err) {
      return { error: err.toString() };
    }
  });
}

// ====================================================
// MODUL KALKULATOR PACING (Rumus Riegel)
// ====================================================

function hitungPacingTarget(token, pbDetik, jarakTarget) {
  var sesi;
  try {
    sesi = butuhPeran(token, ["pelatih", "wali"]);
  } catch (e) {
    return { error: e.message };
  }
  try {
    var pb = Number(pbDetik);
    var jarak = Number(jarakTarget);
    if (!(pb > 0)) return { error: "PB (detik) harus lebih dari 0." };
    if (!(jarak > 0)) return { error: "Jarak target harus lebih dari 0." };
    var total = pb * Math.pow(jarak / 50, 1.06);
    var pacing50 = total / (jarak / 50);
    return {
      totalDetik: Math.round(total * 100) / 100,
      totalWaktuStr: formatWaktuDetik(total),
      pacing50Detik: Math.round(pacing50 * 100) / 100,
      pacing50Str: formatWaktuDetik(pacing50)
    };
  } catch (err) {
    return { error: err.toString() };
  }
}

// ====================================================
// MIGRASI SATU KALI — dari password admin teks biasa (v11)
// ke hash SHA-256 (v12). Dijalankan SEKALI dari editor setelah
// deploy v12. TIDAK pernah mencatat/menampilkan nilai password.
// Prasyarat: Script Properties sudah berisi ADMIN_PASSWORD (teks
// biasa, dibuat oleh setupScriptProperties versi lama) dan
// ADMIN_USERNAME. Setelah berhasil, ADMIN_PASSWORD dihapus.
// ====================================================
function migrasiAdminKeHash() {
  var props = PropertiesService.getScriptProperties();
  var lama = props.getProperty("ADMIN_PASSWORD");
  var sudah = props.getProperty("ADMIN_PASSWORD_HASH");
  if (sudah) {
    Logger.log("Migrasi dilewati: ADMIN_PASSWORD_HASH sudah ada.");
    return "Sudah ada hash, migrasi dilewati.";
  }
  if (!lama) {
    Logger.log("Migrasi GAGAL: ADMIN_PASSWORD tidak ditemukan di Script Properties.");
    return "Gagal: ADMIN_PASSWORD tidak ditemukan.";
  }
  props.setProperty("ADMIN_PASSWORD_HASH", hitungHashSha256(String(lama)));
  props.deleteProperty("ADMIN_PASSWORD");
  Logger.log("Migrasi berhasil: ADMIN_PASSWORD_HASH dibuat, ADMIN_PASSWORD dihapus.");
  return "Migrasi berhasil.";
}

// Cek kesiapan v12 tanpa menampilkan nilai apa pun: hanya true/false.
function cekKesiapanV12() {
  var props = PropertiesService.getScriptProperties();
  var hasil = {
    adaAdminUsername: !!props.getProperty("ADMIN_USERNAME"),
    adaAdminHash: !!props.getProperty("ADMIN_PASSWORD_HASH"),
    adaAdminTeksLama: !!props.getProperty("ADMIN_PASSWORD"),
    adaFonnteToken: !!props.getProperty("FONNTE_TOKEN")
  };
  Logger.log(JSON.stringify(hasil));
  return hasil;
}

// ====================================================
// doPost — API JSON untuk frontend GitHub Pages (v13-pages)
// ====================================================
// Frontend di GitHub Pages memanggil backend lewat HTTP POST ke /exec
// dengan body JSON {fn, args} dan menerima {ok:true,result} atau
// {ok:false,error}. Hanya fungsi yang ada di FUNGSI_DIIZINKAN yang boleh
// dipanggil — yaitu tepat 39 fungsi yang dipakai frontend v13 (13 lewat
// panggilServer + 24 lewat panggilan langsung, 1 irisan: getAthleteData;
// ditambah tambahEventLomba, getFotoAtletMap, dan resetPasswordWali).
// Helper internal (validasiToken, bolehAksesAtlet, dsb.) TIDAK diekspos.
// doGet dan semua fungsi lama di atas TIDAK diubah.
// ====================================================
var FUNGSI_DIIZINKAN = {
  editEventA3Header: editEventA3Header,
  getAthleteData: getAthleteData,
  getDaftarEventBukuAcara: getDaftarEventBukuAcara,
  getDaftarEventLomba: getDaftarEventLomba,
  getDaftarIzinLatihan: getDaftarIzinLatihan,
  getDaftarKonfirmasiSppCoach: getDaftarKonfirmasiSppCoach,
  getFotoAtletMap: getFotoAtletMap,
  getHistoriSppSpesifikAtlet: getHistoriSppSpesifikAtlet,
  getHistoryPresensiSpesifikAtlet: getHistoryPresensiSpesifikAtlet,
  getJadwalBukuAcara: getJadwalBukuAcara,
  getNamaAtletList: getNamaAtletList,
  getRekapMatriksSppAdmin: getRekapMatriksSppAdmin,
  getRekapPendaftaranLomba: getRekapPendaftaranLomba,
  getRekapPresensiAnalisis: getRekapPresensiAnalisis,
  getStatistikPresensiAtlet: getStatistikPresensiAtlet,
  hapusAtlet: hapusAtlet,
  hapusBukuAcara: hapusBukuAcara,
  hapusEventA3: hapusEventA3,
  hapusPendaftaranAtlet: hapusPendaftaranAtlet,
  hapusPresensi: hapusPresensi,
  hapusSpp: hapusSpp,
  hapusWaktuLatihan: hapusWaktuLatihan,
  hitungPacingTarget: hitungPacingTarget,
  kirimIzinLatihanWali: kirimIzinLatihanWali,
  kirimNotifikasiWaTagihanFonnte: kirimNotifikasiWaTagihanFonnte,
  prosesLogin: prosesLogin,
  prosesVerifikasiSppCoach: prosesVerifikasiSppCoach,
  resetPasswordWali: resetPasswordWali,
  simpanAnalisisBukuAcara: simpanAnalisisBukuAcara,
  simpanCatatanKarakter: simpanCatatanKarakter,
  simpanPendaftaranLomba: simpanPendaftaranLomba,
  simpanPresensiSatuAtlet: simpanPresensiSatuAtlet,
  simpanProfilAtlet: simpanProfilAtlet,
  simpanRegistrasiAtletBaru: simpanRegistrasiAtletBaru,
  simpanWaktuLatihan: simpanWaktuLatihan,
  simpanWaktuLatihanMasal: simpanWaktuLatihanMasal,
  submitKonfirmasiSppWali: submitKonfirmasiSppWali,
  tambahEventLomba: tambahEventLomba,
  ubahPassword: ubahPassword
};

function doPost(e) {
  var out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("Body request kosong.");
    var req = JSON.parse(e.postData.contents);
    if (!req || typeof req.fn !== "string" || req.fn === "") throw new Error("Format request tidak valid.");
    var fn = req.fn;
    if (!Object.prototype.hasOwnProperty.call(FUNGSI_DIIZINKAN, fn)) throw new Error("Fungsi tidak dikenal: " + fn);
    var args = Array.isArray(req.args) ? req.args : [];
    var hasil = FUNGSI_DIIZINKAN[fn].apply(null, args);
    return out.setContent(JSON.stringify({ ok: true, result: hasil }));
  } catch (err) {
    return out.setContent(JSON.stringify({ ok: false, error: (err && err.message) ? err.message : String(err) }));
  }
}
