/**
 * FUNGSI UTAMA DOGET (Mendukung Web App GET Request & JSONP Fallback)
 * Wajib berada di tingkat paling atas (Top-level scope)
 */
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || "get_all";
    var callback = params.callback; // Untuk dukungan JSONP jika terkena CORS browser
    var responseData;

    if (action === "stats") {
      responseData = getSummaryStats();
    } else {
      responseData = getAllData();
    }

    var jsonString = JSON.stringify(responseData);

    // Jika dipanggil via JSONP script tag
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + jsonString + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    // Mengembalikan Output JSON dengan header CORS yang diizinkan untuk semua origin
    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    var errorOutput = {
      status: "error",
      message: err.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(errorOutput))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * FUNGSI UTAMA DOPOST (Mendukung Web App POST Request)
 * Wajib berada di tingkat paling atas (Top-level scope)
 */
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

  // Filter sheet yang memiliki nama "Map" (abaikan huruf besar/kecil)
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
    var formulas = range.getFormulas(); // Ambil rumus untuk mendeteksi HYPERLINK()

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
      if (!kode && row[1]) {
        kode = row[1].toString().trim();
      }

      if (kode === "" && !row[1] && !row[2]) continue;

      var numMapMatch = sheetName.match(/\d+/);
      var numMap = numMapMatch ? parseInt(numMapMatch[0], 10) : 1;
      var bidang = (numMap <= 9) ? "Manajerial" : "Supervisi";

      // Ekstraksi Link URL asli jika kolom berisi rumus HYPERLINK("url", "text")
      var rawLink = (row[3] || "").toString().trim();
      var linkFormula = (formulaRow[3] || "").toString().trim();

      if (linkFormula && linkFormula.toUpperCase().indexOf("HYPERLINK") !== -1) {
        var match = linkFormula.match(/HYPERLINK\(\s*"([^"]+)"/i);
        if (match && match[1]) {
          rawLink = match[1];
        }
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

/**
 * Menghitung Statistik Ringkasan
 */
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

/**
 * Mengubah Status Dokumen pada Google Spreadsheet
 */
function updateDocumentStatus(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(payload.nomor_map);
  if (!sheet) return false;

  var parts = payload.id.split("_");
  var rowIndex = parseInt(parts[1], 10) + 1; // 1-indexed

  if (payload.status) sheet.getRange(rowIndex, 5).setValue(payload.status);
  if (payload.link_dokumen) sheet.getRange(rowIndex, 4).setValue(payload.link_dokumen);

  return true;
}

/**
 * Menambahkan Dokumen Baru ke Sheet Terkait
 */
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