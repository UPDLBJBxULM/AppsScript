function telahDiedit(e) {
  // Daftar sheet yang akan didengarkan dan URL Cloud Function yang sesuai
  var sheetConfigs = {
    "RENCANA": "https://asia-southeast2-pln-updl-banjarbaru.cloudfunctions.net/uploadDataExplode",
    "REKAPREALISASI": "https://asia-southeast2-pln-updl-banjarbaru.cloudfunctions.net/uploadDataExplode"
  };

  // Dapatkan sheet aktif dan nama sheet tersebut
  var sheet = e.source.getActiveSheet();
  var activeSheetName = sheet.getName();

  // Cek apakah sheet yang sedang aktif ada di daftar sheet yang didengarkan
  if (sheetConfigs.hasOwnProperty(activeSheetName)) {
    // Cek apakah A1 tidak null atau error
    var cellA1 = sheet.getRange("A1").getValue();
    
    // Pastikan cell A1 tidak null atau error
    if (cellA1 !== null && cellA1 !== "" && !isCellError(cellA1)) {
      // Ambil URL Cloud Function yang sesuai berdasarkan sheet yang sedang aktif
      var url = sheetConfigs[activeSheetName];

      // Payload default adalah objek kosong
      var payload = {};

      // Jika sheet adalah "RAB" atau "REALISASI", tambahkan nama sheet ke payload
      if (activeSheetName === "RENCANA" || activeSheetName === "REKAPREALISASI") {
        payload.sheetName = activeSheetName;
      }

      // Log nama sheet yang mengirimkan payload
      Logger.log("Mengirim payload dari sheet: " + activeSheetName);

      // Kirim request POST ke Cloud Function dengan opsi muteHttpExceptions
      var options = {
        'method' : 'post',
        'contentType': 'application/json',  // Set content type ke JSON
        'payload' : JSON.stringify(payload),  // Kirim payload dengan nama sheet (jika ada)
        'muteHttpExceptions': true  // Tambahkan opsi ini untuk menangani kesalahan HTTP
      };

      try {
        var response = UrlFetchApp.fetch(url, options);
        Logger.log(response.getContentText());  // Log respons jika berhasil
      } catch (error) {
        Logger.log('Error: ' + error.toString());  // Log kesalahan jika terjadi
      }
    }
  }
}

// Fungsi untuk memeriksa apakah ada error di dalam cell
function isCellError(value) {
  // Ini akan menangkap error seperti #N/A, #REF!, #DIV/0!, dll.
  return typeof value === "object" && value instanceof Error;
}
