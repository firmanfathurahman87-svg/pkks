/**
 * DASHBOARD ADMINISTRASI PKKS 2026 - SMP NEGERI 3 CISARUA
 * Backend Google Apps Script (Code.gs)
 */

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || "view";
    var callback = params.callback; // Untuk dukungan JSONP jika terkena CORS browser

    // Jika dipanggil via API untuk mengambil JSON data
    if (action === "get_all" || action === "stats" || callback) {
      var responseData = (action === "stats") ? getSummaryStats() : getAllData();
      var jsonString = JSON.stringify(responseData);

      if (callback) {
        return ContentService.createTextOutput(callback + "(" + jsonString + ")")
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }

      return ContentService.createTextOutput(jsonString)
        .setMimeType(ContentService.MimeType.JSON);
    }

    // TAMPILAN DEFAULT (HTML VIEW): Menampilkan antarmuka Dashboard langsung saat Web App dibuka di browser
    var htmlContent = getDashboardHtml();
    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle("DASHBOARD ADMINISTRASI PKKS 2026 - SMP NEGERI 3 CISARUA")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    var errorOutput = { status: "error", message: err.toString() };
    return ContentService.createTextOutput(JSON.stringify(errorOutput))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var postData = {};
    if (e && e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      postData = e.parameter;
    }

    var result = false;
    if (postData.action === "add") {
      result = addDocumentToSheet(postData);
    } else {
      result = updateDocumentStatus(postData);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: result ? "success" : "failed",
      data: postData
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Membaca seluruh data dari Sheet Map 1 s.d Map 12
 */
function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var allDocuments = [];

  var mapSheets = sheets.filter(function(s) {
    return s.getName().toLowerCase().indexOf("map") !== -1;
  });

  if (mapSheets.length === 0) {
    mapSheets = sheets;
  }

  mapSheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    var range = sheet.getDataRange();
    var data = range.getValues();
    var formulas = range.getFormulas();

    if (data.length <= 1) return;

    var startRow = 1; 
    if (data[0] && data[0][0] && isNaN(data[0][0]) && data[0][0].toString().toLowerCase().includes("kode")) {
      startRow = 1;
    } else if (data.length > 2 && data[1][0] && data[1][0].toString().toLowerCase().includes("kode")) {
      startRow = 2;
    }

    for (var i = startRow; i < data.length; i++) {
      var row = data[i];
      var formulaRow = formulas[i] || [];
      
      var kode = (row[0] !== undefined && row[0] !== null) ? row[0].toString().trim() : "";
      if (!kode && row[1]) kode = row[1].toString().trim();
      if (kode === "" && !row[1] && !row[2]) continue;

      var numMapMatch = sheetName.match(/\d+/);
      var numMap = numMapMatch ? parseInt(numMapMatch[0], 10) : 1;
      var bidang = (numMap <= 9) ? "Manajerial" : "Supervisi";

      var rawLink = (row[3] || "").toString().trim();
      var linkFormula = (formulaRow[3] || "").toString().trim();

      if (linkFormula && linkFormula.toUpperCase().indexOf("HYPERLINK") !== -1) {
        var match = linkFormula.match(/HYPERLINK\(\s*"([^"]+)"/i);
        if (match && match[1]) rawLink = match[1];
      }

      allDocuments.push({
        id: sheetName.replace(/\s+/g, '_') + "_" + i,
        kode_bukti: kode || ("ID-" + (i + 1)),
        nomor_map: sheetName.toLowerCase().indexOf("map") !== -1 ? sheetName : ("Map " + numMap),
        bidang: bidang,
        nama_indikator: (row[1] || row[0] || "Indikator PKKS").toString().trim(),
        nama_bukti: (row[2] || row[1] || "Bukti Fisik").toString().trim(),
        link_dokumen: rawLink,
        status: (row[4] || "Belum Diverifikasi").toString().trim(),
        keterangan_sumber: (row[5] || "").toString().trim()
      });
    }
  });

  return allDocuments;
}

function getSummaryStats() {
  var docs = getAllData();
  var total = docs.length;
  var lengkap = 0, proses = 0, belumTersedia = 0, belumDiverifikasi = 0;

  docs.forEach(function(d) {
    if (d.status === "Lengkap") lengkap++;
    else if (d.status === "Proses") proses++;
    else if (d.status === "Belum Tersedia") belumTersedia++;
    else belumDiverifikasi++;
  });

  return {
    total: total,
    lengkap: lengkap,
    proses: proses,
    belumTersedia: belumTersedia,
    belumDiverifikasi: belumDiverifikasi,
    persentase: total > 0 ? Math.round((lengkap / total) * 100) : 0
  };
}

function updateDocumentStatus(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(payload.nomor_map);
  if (!sheet) return false;

  var parts = payload.id.split("_");
  var rowIndex = parseInt(parts[1], 10) + 1;

  if (payload.status) sheet.getRange(rowIndex, 5).setValue(payload.status);
  if (payload.link_dokumen) sheet.getRange(rowIndex, 4).setValue(payload.link_dokumen);

  return true;
}

function addDocumentToSheet(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(payload.nomor_map);
  if (!sheet) {
    sheet = ss.insertSheet(payload.nomor_map);
    sheet.appendRow(["Kode", "Nama Indikator", "Nama Bukti", "Link", "Status", "Keterangan"]);
  }
  
  sheet.appendRow([
    payload.kode_bukti || "",
    payload.nama_indikator || "",
    payload.nama_bukti || "",
    payload.link_dokumen || "",
    payload.status || "Belum Diverifikasi",
    payload.keterangan_sumber || ""
  ]);
  
  return true;
}

function getDashboardHtml() {
  return `<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <title>DASHBOARD ADMINISTRASI PKKS 2026 - SMP NEGERI 3 CISARUA</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-900 text-white min-h-screen flex items-center justify-center p-4">
    <div class="max-w-md w-full bg-slate-800 p-6 rounded-2xl border border-slate-700 text-center space-y-4 shadow-2xl">
        <div class="w-16 h-16 bg-blue-600 rounded-2xl flex items-center justify-center font-bold text-2xl mx-auto">
            S3C
        </div>
        <h1 class="text-xl font-bold">API Google Apps Script PKKS 2026</h1>
        <p class="text-xs text-slate-400">SMP NEGERI 3 CISARUA (KAB. BANDUNG BARAT)</p>
        <div class="p-3 bg-slate-900/60 rounded-xl text-emerald-400 text-xs font-mono border border-slate-700">
            ✓ Status API: Berhasil Aktif & Terhubung ke Spreadsheet
        </div>
        <p class="text-xs text-slate-300">
            Untuk melihat tampilan dashboard lengkap interaktif, gunakan file <b>index.html</b> dan hubungkan menggunakan URL Web App ini.
        </p>
    </div>
</body>
</html>`;
}