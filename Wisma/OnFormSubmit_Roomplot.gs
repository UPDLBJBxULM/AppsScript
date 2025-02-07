function onFormSubmit_roomPlot(e) {
  if (!e) {
    Logger.log("Tidak ada objek event.");
    return;
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ROOMPLOT");
    if (!sheet) {
      Logger.log("Sheet 'ROOMPLOT' tidak ditemukan.");
      return;
    }

    // Dapatkan baris terakhir yang berisi data
    var lastRow = getLastDataRow(sheet);

    // Baca nilai dari kolom N (status) dan O (processing)
    var range = sheet.getRange(lastRow, 14, 1, 2); // Kolom N dan O
    var [statusValue, processingValue] = range.getValues()[0];

    // Validasi apakah baris sudah diproses atau sedang diproses
    if (statusValue === "SENT" || processingValue === "PROCESSING") {
      Logger.log("Baris sudah diproses atau sedang diproses.");
      return;
    }

    // Tandai baris sebagai PROCESSING
    range.setValues([["PENDING", "PROCESSING"]]);
    Logger.log("Baris ditandai sebagai PENDING dan PROCESSING.");

    // Validasi status publikasi di sheet 'Publikasi'
    var publikasiSheet = ss.getSheetByName("Publikasi_Roomplot");
    if (!publikasiSheet) {
      Logger.log("Sheet 'Publikasi_Roomplot' tidak ditemukan.");
      resetStatus(range); // Reset status jika gagal
      return;
    }

    var statusPublikasi = publikasiSheet.getRange("B3").getValue();
    Logger.log("Status Publikasi: " + statusPublikasi);

    // Panggil roomPlot dengan parameter status publikasi
    try {
      roomPlot(statusPublikasi === "Aktif"); // Kirim true jika aktif, false jika tidak aktif
      Logger.log("roomPlot selesai diproses.");

      // Update status menjadi SENT dan reset PROCESSING
      range.setValues([["SENT", ""]]);
      Logger.log("Status diperbarui: SENT dan PROCESSING direset.");
    } catch (error) {
      Logger.log("Error saat menjalankan roomPlot: " + error);
      resetStatus(range); // Reset status jika terjadi error
    }
  } catch (error) {
    Logger.log("Error umum: " + error);
  }
}

// Fungsi helper untuk mendapatkan baris terakhir yang berisi data
function getLastDataRow(sheet) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A2:A" + lastRow); // Sesuaikan kolom A sesuai kebutuhan
  var values = range.getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 2; // +2 karena offset header dan index dimulai dari 0
    }
  }
  return 2; // Default ke baris pertama data jika tidak ada data
}

// Fungsi helper untuk mereset status ke PENDING dan mengosongkan PROCESSING
function resetStatus(range) {
  range.setValues([["PENDING", ""]]);
  Logger.log("Status direset: PENDING dan PROCESSING dikosongkan.");
}