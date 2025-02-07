function roomPlot(isPublikasiAktif) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Tidak dapat memperoleh lock. Proses sudah berjalan.");
    return;
  }
  try {
    var startTime = new Date().getTime();
    var maxRuntime = 6 * 60 * 1000; // 6 menit batas eksekusi
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ROOMPLOT");
    if (!sheet) {
      Logger.log("Sheet 'ROOMPLOT' tidak ditemukan!");
      return;
    }
    var publikasiSheet = ss.getSheetByName("Publikasi_Roomplot");
    var urlFonnte = "https://api.fonnte.com/send";
    var tokenFonnte = publikasiSheet?.getRange('D5').getValue(); // Optional chaining
    if (!tokenFonnte) {
      Logger.log("Token Fonnte belum diisi.");
      return;
    }

    // Helper function untuk format tanggal
    function formatDate(date) {
      if (!date || isNaN(date.getTime())) return "Invalid Date";
      var day = date.getDate().toString().padStart(2, '0');
      var month = (date.getMonth() + 1).toString().padStart(2, '0');
      var year = date.getFullYear();
      return day + '/' + month + '/' + year;
    }

    // Helper function untuk format nomor telepon
    function formatPhone(num) {
      if (!num) return '';
      num = num.toString().trim();
      if (num.startsWith('+628')) return '628' + num.substring(4);
      else if (num.startsWith('08')) return '628' + num.substring(2);
      else if (num.startsWith('8')) return '628' + num.substring(1);
      else if (num.startsWith('628')) return num;
      return num;
    }

    // Ambil semua data sekaligus
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("Tidak ada data lebih lanjut untuk diproses.");
      return;
    }
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 14).getValues(); // Kolom A-N
    var updates = []; // Untuk menyimpan update status

    for (var i = 0; i < dataRange.length; i++) {
      var now = new Date().getTime();
      if (now - startTime > (maxRuntime - 10 * 1000)) {
        Logger.log("Mendekati batas waktu eksekusi, menghentikan proses sementara.");
        break;
      }

      var row = dataRange[i];
      if (row[13] !== "PENDING") continue; // Lewati jika bukan PENDING

      var rowNum = i + 2; // Sesuaikan dengan row spreadsheet
      var rawTimestamp = row[0];
      var timestamp = new Date(rawTimestamp);
      if (isNaN(timestamp.getTime())) timestamp = new Date();

      var reservasiIn = row[7];
      var reservasiOut = row[8];
      var nomorKamar = row[11];
      var guestId = row[12];
      var rawPhone = row[4];
      var phone1 = formatPhone(rawPhone);

      // Ambil nomor dari sheet Publikasi jika ada
      var reportNum1 = '';
      var reportNum2 = '';
      if (publikasiSheet) { // Pastikan sheet Publikasi ada
        reportNum1 = formatPhone(publikasiSheet.getRange("D8").getValue());
        reportNum2 = formatPhone(publikasiSheet.getRange("D9").getValue());
      }

      // Tentukan target penerima berdasarkan status publikasi
      var targets = [];
      if (phone1) targets.push(phone1);
      if (isPublikasiAktif) {
        if (reportNum1) targets.push(reportNum1);
        if (reportNum2) targets.push(reportNum2);
      }

      var phoneNumber = targets.join(",");
      if (!phoneNumber) {
        Logger.log("Nomor telepon tidak valid untuk baris " + rowNum);
        continue;
      }

      var message = "*Laporan CheckIn* \n\n" +
        "🗓️ *Reservasi In*: " + formatDate(reservasiIn) + "\n" +
        "🗓️ *Reservasi Out*: " + formatDate(reservasiOut) + "\n" +
        "🏢 *Nomor Kamar*: " + nomorKamar + "\n" +
        "👤 *Guest ID*: " + guestId + "\n" +
        "_sent by: Auto Report System_";

      var payload = {
        "target": phoneNumber,
        "message": message
      };
      var options = {
        "method": "post",
        "headers": { "Authorization": tokenFonnte },
        "payload": payload,
        "muteHttpExceptions": true
      };

      try {
        Utilities.sleep(1000); // Jeda 1 detik antar request
        var response = UrlFetchApp.fetch(urlFonnte, options);
        var responseText = response.getContentText();
        Logger.log("Pesan terkirim ke " + phoneNumber + ": " + responseText);
      } catch (error) {
        Logger.log("Error saat mengirim pesan: " + error);
      }

      Utilities.sleep(10 * 1000); // Jeda 10 detik antar request
    }

    // Update status secara batch
    if (updates.length > 0) {
      var batchUpdates = updates.map(update => ({
        range: update.range,
        values: [[update.value]]
      }));
      sheet.getRangeList(batchUpdates.map(u => u.range)).setValues(batchUpdates.map(u => u.values));
    }

    Logger.log("Selesai memproses Queue sementara.");
  } finally {
    lock.releaseLock();
  }
}