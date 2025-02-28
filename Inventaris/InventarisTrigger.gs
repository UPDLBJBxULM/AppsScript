function InventarisTrigger(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Publikasi");
  if (!sheet) {
    Logger.log("Sheet " + sheet + " tidak ditemukan.");
    return;
  }

  var scriptStatus = sheet.getRange(3, 2).getValue(); // Dapatkan status Script di halaman publikasi

  Logger.log("Script status: " + scriptStatus);

  if (scriptStatus == "Aktif") {
    if (!e || !e.range) return Logger.log('e tidak ada');

    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();

    Logger.log("Nama Sheet: " + sheetName);

    if (sheetName == 'LOGOUT02') {
      peminjamanAlat(e);
      pdfPeminjaman(e);
    } else {
      pengembalianAlat(e);
      pdfPengembalian(e);
    }
  } else {
    return Logger.log("Status Script Tidak Aktif");
  }
}
