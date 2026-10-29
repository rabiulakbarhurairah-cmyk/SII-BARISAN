/**
 * SI-BARISAN - Backend Google Apps Script
 * Menyimpan data penilaian ke Google Sheets dan bukti dukung ke Google Drive.
 *
 * CARA PAKAI (pemasangan pertama)
 * 1. script.google.com > Proyek baru, hapus isi bawaan, lalu tempel seluruh kode ini.
 * 2. Pastikan TOKEN di bawah sama persis dengan DRIVE_TOKEN di main.js.
 * 3. Pilih fungsi setup, klik Jalankan satu kali, setujui izin akses.
 * 4. Terapkan > Deployment baru > jenis "Aplikasi web":
 *    Jalankan sebagai "Saya", Yang memiliki akses "Siapa saja". Salin URL berakhiran /exec.
 * 5. Tempel URL itu ke DRIVE_URL di main.js.
 *
 * MEMPERBARUI KODE YANG SUDAH TERPASANG
 * Terapkan > Kelola deployment > ikon pensil > Versi: "Versi baru" > Terapkan.
 * Jangan membuat "Deployment baru" karena URL-nya akan berubah.
 */
var TOKEN = "10928300";
var NAMA_SPREADSHEET = "SI-BARISAN Data";
var NAMA_FOLDER = "SI-BARISAN Bukti Dukung";
var AKUN_ADMIN = "kanwil"; // nama pengguna admin: melihat semua data, dan inputnya terlihat semua akun
var AKUN_KAB = "kankemenag"; // akun Kankemenag: input KUA dan MIN dibagikan ke akun KUA dan MIN

var HEADERS = {
  Penilaian: ["ID Penilaian", "Waktu Kirim", "Tanggal", "Satuan Kerja", "Jenis", "Pengguna", "Tahun Pertama",
    "Punya Asrama/Dapur/Kantin", "NB", "NS", "NK", "NT", "IGBS", "Kategori", "Standar BARISAN",
    "Standar SIPAKAINGE", "Standar Kolaborasi", "Butir Berbukti", "Jumlah Berkas",
    "Jalur", "Kab/Kota", "Kecamatan"],
  Butir: ["ID Penilaian", "Tanggal", "Satuan Kerja", "Standar", "Unsur", "No", "Butir", "Skor Input",
    "Skor Dihitung", "Bukti Sah", "Catatan Pengamatan", "Jumlah Berkas"],
  Bukti: ["ID Penilaian", "Tanggal", "Satuan Kerja", "Standar", "Unsur", "Butir", "Nama Berkas",
    "Tanggal Foto", "Lokasi Foto", "Ukuran (KB)", "Tautan Drive"],
  Partisipatif: ["ID", "Waktu Kirim", "Kategori", "Nama Lembaga", "Kab/Kota",
    "Penanggung Jawab dan HP", "Status", "Didaftarkan Oleh",
    "Jenis Lembaga", "Kegiatan", "Jumlah Kegiatan", "Jumlah Berkas",
    "Diverifikasi Oleh", "Waktu Verifikasi"],
    "Bukti Partisipatif": ["ID Pendaftaran", "Tanggal", "Nama Lembaga", "Kategori", "Kegiatan",
    "Jenis Bukti", "Berkas atau Catatan", "Tanggal Foto", "Lokasi Foto", "Ukuran (KB)", "Tautan Drive"]
};
var LEBAR = {
  Penilaian: [150, 150, 100, 230, 110, 110, 90, 130, 70, 70, 70, 70, 70, 110, 130, 140, 130, 110, 100, 70, 190, 150],
  Butir: [150, 100, 230, 160, 140, 50, 420, 80, 90, 80, 300, 90],
  Bukti: [150, 100, 230, 160, 140, 320, 260, 100, 200, 90, 320],
  Partisipatif: [150, 150, 170, 260, 190, 240, 110, 110, 200, 520, 110, 110, 150, 150],
  "Bukti Partisipatif": [150, 100, 230, 170, 420, 90, 320, 100, 200, 90, 320]
};

/* ===== Titik masuk ===== */
function doGet() {
  return json_({ ok: true, app: "SI-BARISAN", info: "Backend aktif" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: "Request POST tidak memiliki data." });
    }

    var p;
    try {
      p = JSON.parse(e.postData.contents);
    } catch (err) {
      return json_({ ok: false, error: "Data POST bukan JSON yang valid." });
    }

    if (!p.token || String(p.token) !== String(TOKEN)) {
      return json_({ ok: false, error: "Token tidak valid." });
    }

    if (p.action === "save") return json_(simpan_(p));
    if (p.action === "upload") return json_(unggah_(p));
    if (p.action === "list") return json_(daftar_(p));
    if (p.action === "part_save") return json_(partSimpan_(p));
    if (p.action === "part_list") return json_(partDaftar_(p));
    if (p.action === "part_upload") return json_(partUnggah_(p));
    if (p.action === "part_verify") return json_(partVerifikasi_(p));

    return json_({ ok: false, error: "Aksi tidak dikenal: " + String(p.action || "") });

  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* Jalankan sekali dari editor untuk membuat spreadsheet dan folder */
function setup() {
  Logger.log("Spreadsheet: " + getSS_().getUrl());
  Logger.log("Folder bukti: " + akar_().getUrl());
}

/* ===== Penyimpanan data ===== */
function simpan_(p) {
  var ss = getSS_(), id = String(p.id), a = p.penilaian || {}, butir = p.butir || [];
  var sp = ss.getSheetByName("Penilaian"), sb = ss.getSheetByName("Butir"), su = ss.getSheetByName("Bukti");

  [sp, sb, su].forEach(function (s) { hapusId_(s, id); });   // kirim ulang = timpa, bukan dobel

  sp.appendRow([id, new Date(), a.tanggal, t_(a.satuan), t_(a.jenis), t_(a.pengguna),
    a.tahunPertama ? "Ya" : "Tidak", a.asrama ? "Ya" : "Tidak",
    n_(a.nb), n_(a.ns), n_(a.nk), n_(a.nt), n_(a.igbs), t_(a.kategori),
    a.okB ? "Memenuhi" : "Belum", a.okS ? "Memenuhi" : "Belum", a.okK ? "Memenuhi" : "Belum",
    n_(a.butirBerbukti), n_(a.jumlahBerkas),
    t_(a.jalur), t_(a.kab), t_(a.kec)]);

  var rows = butir.map(function (b) {
    return [id, a.tanggal, t_(a.satuan), t_(b.standar), t_(b.unsur), n_(b.no), t_(b.butir),
      t_(b.skorInput), t_(b.skorHitung), t_(b.buktiSah), t_(b.catatan), n_(b.berkas)];
  });
  if (rows.length) sb.getRange(sb.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { ok: true, id: id, butir: rows.length };
}

function unggah_(p) {
  if (!p.data) throw new Error("Data berkas kosong");
  var ss = getSS_(), su = ss.getSheetByName("Bukti");
  var folder = sub_(sub_(akar_(), aman_(p.satuan)), String(p.id));
  var nama = aman_(p.kode + "_" + p.nama);
  var ada = folder.getFilesByName(nama), file;
  if (ada.hasNext()) {
    file = ada.next();
  } else {
    file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || "application/octet-stream", nama));
  }
  su.appendRow([String(p.id), p.tanggal, t_(p.satuan), t_(p.standar), t_(p.unsur), t_(p.butir),
    t_(p.nama), p.tgl || "", t_(p.lok), n_(p.ukuran), file.getUrl()]);
  return { ok: true, url: file.getUrl() };
}

/* ===== Membaca kembali data (dipakai saat login) =====
   Admin melihat semua baris. Akun lain melihat baris miliknya
   dan baris yang diinput admin (sama dengan aturan tampilan di main.js). */
function daftar_(p) {
  var ss = getSS_(), sp = ss.getSheetByName("Penilaian"), last = sp.getLastRow();
  var pengguna = String(p.pengguna || "");
  if (last < 2) return { ok: true, data: [] };

  var v = sp.getRange(2, 1, last - 1, HEADERS.Penilaian.length).getValues(), out = [];
    v.forEach(function (r) {
    var by = String(r[5]), jn = String(r[4]);
    var dariKab = by === AKUN_KAB &&
      ((jn === "KUA" && p.role === "unit") || (jn === "MIN" && p.role === "madrasah"));
    var keKab = p.role === "kab" && (jn === "KUA" || jn === "MIN");
    var dariAdmin = by === AKUN_ADMIN && jn !== "Kanwil";
    if (pengguna !== AKUN_ADMIN && by !== pengguna && !dariAdmin && !dariKab && !keKab) return;
    out.push({
      id: String(r[0]),
      tgl: tgl_(r[2]),
      nm: String(r[3]),
      jn: String(r[4]),
      by: by,
      nb: Number(r[8]) || 0,
      ns: Number(r[9]) || 0,
      nk: Number(r[10]) || 0,
      nt: Number(r[11]) || 0,
      jl: String(r[19] || ""),
      kb: String(r[20] || ""),
      kc: String(r[21] || "")
    });
  });
  return { ok: true, data: out };
}

/* ===== Spreadsheet dan folder ===== */
function getSS_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty("SS_ID"), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  var baru = !ss;
  if (baru) { ss = SpreadsheetApp.create(NAMA_SPREADSHEET); props.setProperty("SS_ID", ss.getId()); }
  var awal = baru ? ss.getSheets()[0] : null;
  Object.keys(HEADERS).forEach(function (n) { siapkan_(ss, n); });
  if (awal && ss.getSheets().length > 1) ss.deleteSheet(awal);
  return ss;
}

function siapkan_(ss, nama) {
  var sh = ss.getSheetByName(nama), h = HEADERS[nama];
  if (!sh) sh = ss.insertSheet(nama);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight("bold").setFontColor("#ffffff")
      .setBackground("#0e4f55").setWrap(true).setVerticalAlignment("middle");
    sh.setFrozenRows(1);
    LEBAR[nama].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    var baris = sh.getMaxRows() - 1;
    if (nama === "Penilaian") {
      sh.getRange(2, 2, baris, 1).setNumberFormat("yyyy-mm-dd hh:mm");
      sh.getRange(2, 3, baris, 1).setNumberFormat("yyyy-mm-dd");
      sh.getRange(2, 9, baris, 5).setNumberFormat("0.0");
    } else {
      sh.getRange(2, 2, baris, 1).setNumberFormat("yyyy-mm-dd");
    }
    sh.getRange(1, 1, sh.getMaxRows(), h.length).createFilter();
  } else if (sh.getLastColumn() < h.length) {
    // Spreadsheet lama: tambahkan kolom baru (Jalur, Kab/Kota, Kecamatan) di sebelah kanan
    var k = sh.getLastColumn();
    sh.getRange(1, k + 1, 1, h.length - k).setValues([h.slice(k)]).setFontWeight("bold")
      .setFontColor("#ffffff").setBackground("#0e4f55").setWrap(true).setVerticalAlignment("middle");
    for (var i = k; i < h.length; i++) sh.setColumnWidth(i + 1, LEBAR[nama][i]);
  }
  return sh;
}

function akar_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty("FOLDER_ID"), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); } catch (e) { f = null; } }
  if (!f) { f = DriveApp.createFolder(NAMA_FOLDER); props.setProperty("FOLDER_ID", f.getId()); }
  return f;
}

function sub_(induk, nama) {
  var it = induk.getFoldersByName(nama);
  return it.hasNext() ? it.next() : induk.createFolder(nama);
}

/* ===== Pembantu ===== */
function hapusId_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === id) sh.deleteRow(i + 2);
  }
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function t_(v) {   // teks aman: cegah isi sel dibaca sebagai rumus
  v = (v === null || v === undefined) ? "" : String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function n_(v) { var x = Number(v); return isFinite(x) ? Math.round(x * 10) / 10 : 0; }
function aman_(s) { return String(s || "tanpa-nama").replace(/[\\\/:*?"<>|]/g, "-").slice(0, 120); }
function tgl_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v || "");
}

/* ===== Lembaga Partisipatif ===== */
function partSimpan_(p) {
  partIzin_(p);
  var ss = getSS_(), d = p.data || {}, id = String(p.id || new Date().getTime());
  var sh = ss.getSheetByName("Partisipatif"), sb = ss.getSheetByName("Bukti Partisipatif");
  var keg = (d.kegiatan || []).map(function (x) { return String(x); });

  hapusId_(sh, id);   // kirim ulang = timpa, bukan dobel
  hapusId_(sb, id);

  sh.appendRow([id, new Date(), t_(d.kategori), t_(d.nama), t_(d.kab), t_(d.pj),
    "Menunggu", t_(d.pengguna), t_(d.jenis), t_(keg.join("\n")), keg.length, n_(d.jumlahBerkas)]);

  (d.catatan || []).forEach(function (c) {
    sb.appendRow([id, new Date(), t_(d.nama), t_(d.kategori), t_(c.kegiatan),
      "Catatan", t_(c.teks), "", "", "", ""]);
  });
  return { ok: true, id: id };
}

function partDaftar_(p) {
  var ss = getSS_(), sh = ss.getSheetByName("Partisipatif"), last = sh.getLastRow();
  if (last < 2) return { ok: true, data: [] };

  var role = String(p.role || "");
  var kelola = role === "kanwil" || role === "kab";   // hanya pengelola melihat kontak dan bukti

  var bm = {};
  if (kelola) {
    var sb = ss.getSheetByName("Bukti Partisipatif");
    if (sb.getLastRow() > 1) {
      sb.getRange(2, 1, sb.getLastRow() - 1, HEADERS["Bukti Partisipatif"].length).getValues()
        .forEach(function (r) {
          var k = String(r[0]);
          (bm[k] = bm[k] || []).push({ kegiatan: String(r[4]), jenis: String(r[5]), isi: String(r[6]),
            tgl: tgl_(r[7]), lok: String(r[8] || ""), url: String(r[10] || "") });
        });
    }
  }

  var v = sh.getRange(2, 1, last - 1, HEADERS.Partisipatif.length).getValues(), out = [];
  v.forEach(function (r) {
    out.push({
      id: String(r[0]), tgl: tgl_(r[1]), kategori: String(r[2]), nama: String(r[3]),
      kab: String(r[4]), status: String(r[6] || "Menunggu"),
      jenis: String(r[8] || ""),
      kegiatan: String(r[9] || "").split("\n").filter(function (x) { return x; }),
      berkas: Number(r[11]) || 0,
      oleh: String(r[12] || ""),
      pj: kelola ? String(r[5]) : "",
      by: kelola ? String(r[7]) : "",
      bukti: kelola ? (bm[String(r[0])] || []) : []
    });
  });
  return { ok: true, data: out };
}

function partUnggah_(p) {
  partIzin_(p);
  if (!p.data) throw new Error("Data berkas kosong");
  var sb = getSS_().getSheetByName("Bukti Partisipatif");
  var folder = sub_(sub_(sub_(akar_(), "Partisipatif"), aman_(p.nama)), String(p.id));
  var nama = aman_(p.kode + "_" + p.berkas);
  var ada = folder.getFilesByName(nama), file;
  if (ada.hasNext()) {
    file = ada.next();
  } else {
    file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || "application/octet-stream", nama));
  }
  sb.appendRow([String(p.id), new Date(), t_(p.nama), t_(p.kategori), t_(p.kegiatan),
    "Berkas", t_(p.berkas), p.tgl || "", t_(p.lok), n_(p.ukuran), file.getUrl()]);
  return { ok: true, url: file.getUrl() };
}

function partVerifikasi_(p) {
  var role = String(p.role || "");
  if (role !== "kanwil" && role !== "kab") {
    throw new Error("Hanya Kanwil dan Kankemenag yang dapat memverifikasi.");
  }

  var sh = getSS_().getSheetByName("Partisipatif"), last = sh.getLastRow();
  var status = p.status === "Menunggu" ? "Menunggu" : "Terverifikasi";
  if (last < 2) throw new Error("Data tidak ditemukan.");

  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(p.id)) {
      var baris = i + 2, ok = status === "Terverifikasi";
      sh.getRange(baris, 7).setValue(status);
      sh.getRange(baris, 13, 1, 2).setValues([[ok ? String(p.pengguna || "") : "", ok ? new Date() : ""]]);
      return { ok: true, id: String(p.id), status: status };
    }
  }
  throw new Error("Pendaftaran tidak ditemukan.");
}

function partIzin_(p) {
  var role = String(p.role || "");
  if (role !== "kanwil" && role !== "kab") {
    throw new Error("Hanya Kanwil dan Kankemenag yang dapat mendaftarkan lembaga partisipatif.");
  }
}
