function senderMessage(pesan, inputNumber) {
  Logger.log("Pesan yang didapatkan: " + pesan);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Publikasi");
  if (!sheet) {
    Logger.log("Sheet 'reportTo' tidak ditemukan.");
    return;
  }

  var tokenFonnte = sheet.getRange(5, 4).getValue(); // Isi dengan token Fonnte Anda
  var url = "https://api.fonnte.com/send";

  var nomor1 = formatPhone(sheet.getRange(8, 4).getValue()); // Kolom D, baris 8
  var nomor2 = formatPhone(sheet.getRange(9, 4).getValue()); // Kolom D, baris 9

  // Nomor-nomor tujuan (2 nomor statik + inputNumber)
  var targetNumbers = [
    nomor1, // Nomor statik pertama
    nomor2, // Nomor statik kedua
    inputNumber // Nomor dinamis dari input
  ];

  targetNumbers.forEach(target => {
    if (target) { // Periksa jika target tidak kosong atau undefined
      var options = { 
        "method": "post",
        "headers": {
          "Authorization": tokenFonnte
        },
        "payload": {
          "target": target,
          "message": pesan
        }
      };

      // Kirim pesan
      try {
        var response = UrlFetchApp.fetch(url, options);
        Logger.log("Pesan berhasil dikirim ke: " + target);
        Logger.log("Response: " + response.getContentText());
      } catch (e) {
        Logger.log("Gagal mengirim pesan ke: " + target);
        Logger.log("Error: " + e.message);
      }
    } else {
      Logger.log("Target kosong, pesan tidak dikirim.");
    }

    Utilities.sleep(10 * 1000);
  });
}

function formatPhone(num) {
  if (!num) return '';
  num = num.toString().trim();
  if (num.startsWith('+628')) {
    return '628' + num.substring(4);
  } else if (num.startsWith('08')) {
    return '628' + num.substring(2);
  } else if (num.startsWith('628')) {
    return num;
  }
  return num;
}