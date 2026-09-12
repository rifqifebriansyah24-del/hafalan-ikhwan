/**
 * Hafalan Kelas Ikhwan - setup script for Google Forms + Sheets.
 *
 * How to use:
 * 1. Go to https://script.google.com, create a new project.
 * 2. Delete the default code, paste this whole file in.
 * 3. Run setupHafalanTracker() once (top toolbar: pick the function, click Run).
 *    First run asks for authorization - allow it (the script only touches
 *    what it creates in your own account).
 * 4. Open View > Logs (or Ctrl+Enter) to get the Form and Spreadsheet links.
 * 5. Run seedInitialData() once to load the 10 Sept 2026 session data.
 * 6. Open the Spreadsheet, Share > set "Anyone with the link" to Viewer.
 *    Share the Form's fill-in link with musyrif, the Spreadsheet's
 *    Dashboard tab with the koordinator, and teach peserta to use
 *    View > Filter views > Create new filter view on "Form Responses 1"
 *    to see just their own rows without affecting anyone else's view.
 */

var PESERTA_NAMES = [
  "Muhammad Rikza", "Ottoh Hidayatullah", "Wawan Abu Kairo", "Hussein Abrar", "Adi Muliadi",
  "Tulus", "Asep Saepuloh", "Andi Andriyanto", "Maulana Hasanuddin", "Budi",
  "Soultan Salam Abu Kholid", "Hamzah", "Sukirno", "Muhammad Rizman", "Muhamad Hidayat",
  "Afwadi", "Agus Tri Yulianto", "Tuhiri Handoyo", "Jati Purnomo", "Tuwan Abu Abiyyu",
  "Ibrahim", "Afwan", "Fajrin Budiawan", "Vito", "Andi Suwardi",
  "Saprul Supardi", "Erwin", "Wawan.S", "Hukma Shobiyya", "Hafid Rizki",
  "Bagus", "Widi", "Bimo", "Ismulyadi", "Muhammad Tahir",
  "Suwandi H", "Imam", "Fuad Adam", "Tri Himawan", "Romli",
  "Hikmah", "Dafa Baihaqi", "Suprapto", "Icen"
];

function setupHafalanTracker() {
  var form = FormApp.create("Hafalan Kelas Ikhwan");
  form.setDescription("Kelas Ikhwan - Rumah Quran Tj. Priok. Isi setelah menyimak satu peserta.");

  form.addListItem().setTitle("Nama peserta").setChoiceValues(PESERTA_NAMES).setRequired(true);
  form.addDateItem().setTitle("Tanggal").setRequired(true);
  form.addMultipleChoiceItem().setTitle("Jenis").setChoiceValues(["Hafalan baru", "Tes", "Murajaah"]).setRequired(true);

  var juzChoices = [];
  for (var j = 1; j <= 30; j++) juzChoices.push(String(j));
  form.addListItem().setTitle("Juz").setChoiceValues(juzChoices).setRequired(false);

  form.addTextItem().setTitle("Surah").setRequired(false);
  form.addTextItem().setTitle("Ayat / batas hafalan").setRequired(false);
  form.addMultipleChoiceItem().setTitle("Kehadiran").setChoiceValues(["Hadir", "Izin"]).setRequired(true);
  form.addCheckboxItem().setTitle("Tandai untuk dicek koordinator").setChoiceValues(["Data belum pasti"]);
  form.addParagraphTextItem().setTitle("Catatan (opsional)");
  form.addTextItem().setTitle("Dicatat oleh (musyrif)");

  var ss = SpreadsheetApp.create("Hafalan Kelas Ikhwan - Data");
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  var responseSheet = findResponseSheet_(ss);
  for (var attempt = 0; attempt < 5 && !responseSheet; attempt++) {
    Utilities.sleep(1000);
    responseSheet = findResponseSheet_(ss);
  }

  var blank = ss.getSheetByName("Sheet1");
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  var pesertaSheet = ss.insertSheet("Peserta");
  pesertaSheet.getRange(1, 1).setValue("Nama peserta");
  pesertaSheet.getRange(2, 1, PESERTA_NAMES.length, 1).setValues(
    PESERTA_NAMES.map(function (n) { return [n]; })
  );

  var dashboard = ss.insertSheet("Dashboard");
  dashboard.getRange(1, 1, 1, 5).setValues([["Nama", "Status", "Terakhir dicatat", "Progres", "Juz disentuh"]]);
  dashboard.setFrozenRows(1);
  dashboard.setColumnWidths(1, 5, 180);

  installTrigger_(ss);
  refreshDashboard_(ss);

  Logger.log("Form (isi jawaban): " + form.getPublishedUrl());
  Logger.log("Form (edit pertanyaan): " + form.getEditUrl());
  Logger.log("Spreadsheet: " + ss.getUrl());
}

function findResponseSheet_(ss) {
  var found = null;
  ss.getSheets().forEach(function (sh) {
    if (sh.getName().indexOf("Form Responses") === 0) found = sh;
  });
  return found;
}

function installTrigger_(ss) {
  ScriptApp.newTrigger("onFormSubmitRefresh").forSpreadsheet(ss).onFormSubmit().create();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Hafalan").addItem("Refresh Dashboard", "refreshDashboardNow").addToUi();
}

function onFormSubmitRefresh() {
  refreshDashboard_(SpreadsheetApp.getActiveSpreadsheet());
}

function refreshDashboardNow() {
  refreshDashboard_(SpreadsheetApp.getActiveSpreadsheet());
}

function refreshDashboard_(ss) {
  var responseSheet = findResponseSheet_(ss);
  if (!responseSheet) return;

  var data = responseSheet.getDataRange().getValues();
  var headers = data[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  var byName = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var name = row[col["Nama peserta"]];
    if (!name) continue;
    if (!byName[name]) byName[name] = [];
    byName[name].push(row);
  }

  var pesertaSheet = ss.getSheetByName("Peserta");
  var roster = pesertaSheet.getRange(2, 1, pesertaSheet.getLastRow() - 1, 1)
    .getValues().map(function (r) { return r[0]; }).filter(String);

  var out = roster.map(function (name) {
    var entries = (byName[name] || []).slice().sort(function (a, b) {
      return new Date(b[col["Timestamp"]]).getTime() - new Date(a[col["Timestamp"]]).getTime();
    });
    var latest = entries[0];
    var status = "Belum ada laporan";
    if (latest) {
      var review = latest[col["Tandai untuk dicek koordinator"]];
      var hadir = latest[col["Kehadiran"]];
      if (review) status = "Perlu dicek";
      else if (hadir === "Izin") status = "Izin";
      else status = "Aktif";
    }
    var latestText = latest
      ? [latest[col["Jenis"]], "Juz " + (latest[col["Juz"]] || "?"), latest[col["Surah"]], latest[col["Ayat"]]]
          .filter(String).join(" - ")
      : "";
    var juzSet = {};
    entries.forEach(function (e) { if (e[col["Juz"]]) juzSet[e[col["Juz"]]] = true; });
    var juzList = Object.keys(juzSet).map(Number).sort(function (a, b) { return a - b; });
    return [name, status, latestText, juzList.length + "/30", juzList.join(", ")];
  });

  var dash = ss.getSheetByName("Dashboard");
  if (dash.getMaxRows() > 1) dash.getRange(2, 1, dash.getMaxRows() - 1, 5).clearContent();
  if (out.length) dash.getRange(2, 1, out.length, 5).setValues(out);
}

function seedInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = findResponseSheet_(ss);
  if (!responseSheet) {
    Logger.log("Sheet 'Form Responses' belum ada - jalankan setupHafalanTracker() dulu.");
    return;
  }

  var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  var seedDate = new Date(2026, 8, 10);
  var seedTimestamp = new Date(2026, 8, 10, 16, 0, 0);

  var rows = SEED_ENTRIES.map(function (e) {
    var row = new Array(headers.length).fill("");
    row[col["Timestamp"]] = seedTimestamp;
    row[col["Nama peserta"]] = e.nama;
    row[col["Tanggal"]] = seedDate;
    row[col["Jenis"]] = e.jenis;
    row[col["Juz"]] = e.juz || "";
    row[col["Surah"]] = e.surah || "";
    row[col["Ayat"]] = e.ayat || "";
    row[col["Kehadiran"]] = e.kehadiran;
    row[col["Tandai untuk dicek koordinator"]] = e.perluCek ? "Data belum pasti" : "";
    row[col["Catatan"]] = e.catatan || "";
    row[col["Dicatat oleh"]] = "";
    return row;
  });

  responseSheet.getRange(responseSheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  refreshDashboard_(ss);
  Logger.log("Selesai menambahkan " + rows.length + " data awal (10 September 2026).");
}

var SEED_ENTRIES = [
  { nama: "Muhammad Rikza", jenis: "Hafalan baru", juz: 30, surah: "An-Naziat", kehadiran: "Hadir" },
  { nama: "Ottoh Hidayatullah", jenis: "Tes", juz: 1, kehadiran: "Hadir" },
  { nama: "Wawan Abu Kairo", jenis: "Hafalan baru", juz: 3, surah: "Al-Baqarah", ayat: "264", kehadiran: "Hadir" },
  { nama: "Hussein Abrar", jenis: "Tes", juz: 28, kehadiran: "Izin", catatan: "Izin, tes dijadwalkan ulang" },
  { nama: "Adi Muliadi", jenis: "Hafalan baru", juz: 29, surah: "Al-Ma'arij", kehadiran: "Hadir" },
  { nama: "Tulus", jenis: "Tes", juz: 5, surah: "An-Nisa", ayat: "140", kehadiran: "Hadir", catatan: "Tes 1 juz" },
  { nama: "Asep Saepuloh", jenis: "Tes", juz: 1, kehadiran: "Hadir" },
  { nama: "Andi Andriyanto", jenis: "Hafalan baru", juz: 12, surah: "Yusuf", kehadiran: "Hadir" },
  { nama: "Maulana Hasanuddin", jenis: "Hafalan baru", juz: 1, surah: "Al-Baqarah", kehadiran: "Hadir", catatan: "2,5 halaman pertama" },
  { nama: "Budi", jenis: "Tes", juz: 28, kehadiran: "Hadir" },
  { nama: "Soultan Salam Abu Kholid", jenis: "Hafalan baru", juz: 2, surah: "Al-Baqarah", ayat: "220", kehadiran: "Hadir" },
  { nama: "Hamzah", jenis: "Hafalan baru", juz: 2, surah: "Al-Baqarah", ayat: "252", kehadiran: "Hadir" },
  { nama: "Sukirno", jenis: "Tes", juz: 26, kehadiran: "Hadir", catatan: "Tes 1 juz" },
  { nama: "Muhammad Rizman", jenis: "Hafalan baru", juz: 4, surah: "An-Nisa", ayat: "19", kehadiran: "Hadir" },
  { nama: "Muhamad Hidayat", jenis: "Hafalan baru", juz: 7, kehadiran: "Hadir" },
  { nama: "Afwadi", jenis: "Tes", juz: 1, surah: "Al-Baqarah", kehadiran: "Hadir" },
  { nama: "Agus Tri Yulianto", jenis: "Hafalan baru", juz: 1, surah: "Al-Baqarah", ayat: "106", kehadiran: "Hadir", catatan: "2,5 halaman akhir" },
  { nama: "Tuhiri Handoyo", jenis: "Hafalan baru", juz: 4, surah: "Ali Imran", ayat: "173", kehadiran: "Hadir" },
  { nama: "Jati Purnomo", jenis: "Hafalan baru", juz: 30, surah: "An-Naba", kehadiran: "Hadir" },
  { nama: "Tuwan Abu Abiyyu", jenis: "Tes", juz: 1, surah: "Al-Baqarah", ayat: "141", kehadiran: "Hadir", catatan: "Tes 2,5 halaman ke-4" },
  { nama: "Ibrahim", jenis: "Hafalan baru", juz: 1, surah: "Al-Baqarah", ayat: "128", kehadiran: "Hadir", perluCek: true, catatan: "Catatan asli tidak jelas, perlu konfirmasi ke musyrif" },
  { nama: "Afwan", jenis: "Hafalan baru", juz: 29, surah: "Al-Jin", ayat: "20", kehadiran: "Hadir" },
  { nama: "Fajrin Budiawan", jenis: "Hafalan baru", juz: 2, surah: "Al-Baqarah", ayat: "230", kehadiran: "Hadir" },
  { nama: "Vito", jenis: "Murajaah", juz: 29, surah: "Al-Jin", kehadiran: "Hadir" },
  { nama: "Andi Suwardi", jenis: "Hafalan baru", juz: 29, surah: "Al-Jin", ayat: "21", kehadiran: "Hadir" },
  { nama: "Saprul Supardi", jenis: "Hafalan baru", juz: 29, surah: "Al-Muddatsir", ayat: "30", kehadiran: "Hadir" },
  { nama: "Erwin", jenis: "Tes", juz: 4, kehadiran: "Hadir" },
  { nama: "Wawan.S", jenis: "Hafalan baru", juz: 29, surah: "Al-Qalam", ayat: "25", kehadiran: "Hadir" },
  { nama: "Hukma Shobiyya", jenis: "Murajaah", juz: 6, kehadiran: "Hadir", catatan: "Murajaah + tes" },
  { nama: "Hafid Rizki", jenis: "Hafalan baru", juz: 5, surah: "An-Nisa", ayat: "128", kehadiran: "Hadir" },
  { nama: "Bagus", jenis: "Hafalan baru", juz: 4, surah: "Ali Imran", ayat: "108", kehadiran: "Hadir" },
  { nama: "Bimo", jenis: "Hafalan baru", juz: 29, surah: "Al-Qalam", kehadiran: "Hadir" },
  { nama: "Ismulyadi", jenis: "Murajaah", juz: 29, surah: "Al-Muddatsir", kehadiran: "Hadir" },
  { nama: "Imam", jenis: "Murajaah", juz: 28, surah: "Al-Hasyr", kehadiran: "Hadir" },
  { nama: "Fuad Adam", jenis: "Hafalan baru", juz: 4, surah: "Ali Imran", ayat: "186", kehadiran: "Hadir" },
  { nama: "Tri Himawan", jenis: "Hafalan baru", juz: 29, surah: "Nuh", ayat: "10", kehadiran: "Hadir" },
  { nama: "Romli", jenis: "Hafalan baru", surah: "Al-Baqarah", kehadiran: "Hadir", perluCek: true, catatan: "Juz & ayat belum tercatat lengkap" },
  { nama: "Hikmah", jenis: "Hafalan baru", juz: 29, surah: "Al-Qalam", ayat: "35", kehadiran: "Hadir" },
  { nama: "Dafa Baihaqi", jenis: "Hafalan baru", juz: 30, kehadiran: "Hadir", perluCek: true, catatan: "Catatan asli tidak jelas, perlu konfirmasi ke musyrif" },
  { nama: "Suprapto", jenis: "Hafalan baru", juz: 30, surah: "An-Naziat", ayat: "25", kehadiran: "Hadir" },
  { nama: "Icen", jenis: "Hafalan baru", juz: 1, surah: "Al-Baqarah", ayat: "83", kehadiran: "Hadir" }
];
