function doGet() {

  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Buku Tamu Digital Instan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

}

function simpanDataTamu(nama, instansi, keperluan, whatsapp) {

  try {

    var sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Tamu");

    if (!sheet) {
      return "ERROR: Sheet 'Tamu' tidak ditemukan";
    }

    var timestamp = new Date();

    var waktuFormat = Utilities.formatDate(
      timestamp,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm"
    );

    // LEBIH CEPAT DARI appendRow
    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      5
    ).setValues([[
      waktuFormat,
      nama,
      instansi,
      keperluan,
      whatsapp
    ]]);

    return "SUKSES";

  } catch(error) {

    return "ERROR: " + error.toString();

  }

}

function hapusSatuTamu(rowNum) {

  try {

    var sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Tamu");

    if (!sheet) {
      return "ERROR: Sheet tidak ditemukan";
    }

    rowNum = parseInt(rowNum);

    if (isNaN(rowNum) || rowNum <= 1) {
      return "ERROR: Nomor baris tidak valid";
    }

    sheet.deleteRow(rowNum);

    return "SUKSES";

  } catch(error) {

    return "ERROR: " + error.toString();

  }

}

function hapusSemuaTamu() {

  try {

    var sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Tamu");

    if (!sheet) {
      return "ERROR: Sheet tidak ditemukan";
    }

    var lastRow = sheet.getLastRow();

    // Sisakan header
    if (lastRow > 1) {

      sheet.deleteRows(2, lastRow - 1);

    }

    return "SUKSES";

  } catch(error) {

    return "ERROR: " + error.toString();

  }

}

function ambilDataKilat() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tamu"); 
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  
  for (var i = 1; i < data.length; i++) {
    hasil.push({
      rowNum: i + 1,
      waktu: Utilities.formatDate(new Date(data[i][0]), "GMT+7", "dd/MM/yyyy HH:mm"),
      nama: data[i][1],
      instansi: data[i][2],
      keperluan: data[i][3],
      wa: data[i][4],
      disposisi: data[i][5] || "Menunggu Antrean" // <--- Membaca kolom ke-6 (Kolom F)
    });
  }
  return hasil;
}

function simpanDisposisiSheets(rowNum, statusBaru) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tamu");
    
    // Kolom 6 adalah Kolom F (Tempat menyimpan status disposisi)
    sheet.getRange(rowNum, 6).setValue(statusBaru); 
    
    return "SUKSES";
  } catch(e) {
    return e.toString();
  }
}

function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}