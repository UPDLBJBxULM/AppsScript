function processQueue() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Tidak dapat memperoleh lock. Proses sudah berjalan.");
    return;
  }

  try {
    var startTime = new Date().getTime();
    var maxRuntime = 6 * 60 * 1000; // 6 menit batas eksekusi

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RENCANA");
    if (!sheet) {
      Logger.log("Sheet 'RENCANA' tidak ditemukan!");
      return;
    }

    var reportSheet = ss.getSheetByName("Publikasi");
    var linkontxtSheet = ss.getSheetByName("linkONTXT");

    var urlFonnte = "https://api.fonnte.com/send";
    var tokenFonnte = reportSheet.getRange("D5").getValue(); // ISI token Fonnte Anda di sini
    if (!tokenFonnte) {
      Logger.log("Token Fonnte belum diisi.");
      return;
    }

    while (true) {
      var now = new Date().getTime();
      if (now - startTime > (maxRuntime - 10 * 1000)) {
        Logger.log("Mendekati batas waktu eksekusi, menghentikan proses sementara.");
        break;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        Logger.log("Tidak ada data lebih lanjut untuk diproses.");
        break;
      }

      var dataRange = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // Dapatkan data A-K
      var pendingIndex = -1;
      for (var i = 0; i < dataRange.length; i++) {
        if (dataRange[i][10] === "PENDING") { // Cari di kolom K
          pendingIndex = i;
          break;
        }
      }

      if (pendingIndex === -1) {
        Logger.log("Tidak ada lagi data PENDING.");
        break;
      }

      var rowNum = pendingIndex + 2;
      var row = dataRange[pendingIndex];

      var rawTimestamp = row[0];
      var timestamp = new Date(rawTimestamp);
      if (isNaN(timestamp.getTime())) {
        timestamp = new Date();
      }

      var requestor = row[3];
      var perihal = row[4];
      var unit = row[6];
      var nominal = row[7];
      var idRencana = row[8];

      // Ambil nomor tujuan dari kolom F dan format ulang
      var rawPhone = row[5];
      var phone1 = formatPhone(rawPhone);

      // Ambil nomor tambahan dari sheet reportTO
      var reportNum1 = ''; // KEU
      var reportNum2 = ''; // MUPDL
      var reportNum3YAN = ''; // Accountable YAN
      var reportNum4K3L = ''; // Accountable K3L
      if (reportSheet) {
        reportNum1 = formatPhone(reportSheet.getRange("D8").getValue());
        reportNum2 = formatPhone(reportSheet.getRange("D9").getValue());

        // Ambil data terakhir dari sheet RENCANA kolom E
        var lastRowRencana = sheet.getLastRow();
        var accountableBag = sheet.getRange(lastRowRencana, 5).getValue();

        // Ambil string pembanding dari sheet Publikasi (misalnya di kolom F1 dan F2)
        var yanConditionString = reportSheet.getRange("C10").getValue(); // string YAN di C10
        var k3lConditionString = reportSheet.getRange("C11").getValue(); // string K3L di C11

        reportNum3YAN = ''; // Inisialisasi sebagai nonaktif default
        reportNum4K3L = ''; // Inisialisasi sebagai nonaktif default

        if (accountableBag === yanConditionString) {
          reportNum3YAN = formatPhone(reportSheet.getRange("D10").getValue()); // Aktifkan reportNum3YAN
          // reportNum4K3L tetap nonaktif karena sudah diinisialisasi di atas
        } else if (accountableBag === k3lConditionString) {
          reportNum4K3L = formatPhone(reportSheet.getRange("D11").getValue()); // Aktifkan reportNum4K3L
          // reportNum3YAN tetap nonaktif karena sudah diinisialisasi di atas
        }
      }

      // Ambil LinkONTXT
      var linkAR = ''; //"https://drive.google.com/drive/folders/1VGU3E8Dv0o0vXs2JXEul-ge-WabHjlah?usp=sharing"
      var linkSRR = ''; //"https://docs.google.com/spreadsheets/d/1QFVJtpfC7E3WomzZCnmsl4ObeMToNSoY4vuGQ_oPj4s/edit?gid=0#gid=0"
      if (linkontxtSheet) {
        linkAR = linkontxtSheet.getRange("B1").getValue();
        linkSRR = linkontxtSheet.getRange("B2").getValue();
      }

      // Kumpulkan semua nomor ke dalam array dan gabungkan menjadi string comma-separated
      var targets = [];
      if (phone1) targets.push(phone1);
      if (reportNum1) targets.push(reportNum1); // KEU
      if (reportNum2) targets.push(reportNum2); // MUPDL
      if (reportNum3YAN) targets.push(reportNum3YAN); // Accountable YAN
      if (reportNum4K3L) targets.push(reportNum4K3L); // Accountable K3L
      var phoneNumber = targets.join(",");

      function calculateStartDate(date) {
        var day = date.getDay();
        if (day === 0) day = 7;
        var daysToAdd = 8 - day;
        var startDate = new Date(date);
        startDate.setDate(date.getDate() + daysToAdd);
        return startDate;
      }

      function calculateEndDate(startDate) {
        var endDate = new Date(startDate);
        endDate.setDate(startDate.getDate() + 4);
        return endDate;
      }

      function formatDate(date) {
        var day = date.getDate().toString().padStart(2, '0');
        var month = (date.getMonth() + 1).toString().padStart(2, '0');
        var year = date.getFullYear();
        return day + '/' + month + '/' + year;
      }

      var startDateAR = calculateStartDate(timestamp);
      var endDateAR = calculateEndDate(startDateAR);

      var message = "🪙 *Pengajuan Petty Cash* \n\n" +
        "📋 *ID Rencana* " + "" + ": " + (idRencana || 'Data tidak tersedia') + "\n\n" +
        "👤 *Requestor* " + "" + ": " + (requestor || 'Data tidak tersedia') + "\n" +
        "🏢 *Unit* " + "" + ": " + (unit || 'Data tidak tersedia') + "\n\n" +
        "❔ *Perihal* " + "" + ": " + (perihal || 'Data tidak tersedia') + "\n" +
        "💰 *Nominal* " + "" + ": Rp. " + (formatRupiah(nominal) || 'Data tidak tersedia') + "\n\n" +
        "🗓️ *Start Date A/R* " + "" + ": " + formatDate(startDateAR) + "\n" +
        "🗓️ *End Date A/R* " + "" + ": " + formatDate(endDateAR) + "\n\n" +
        "silahkan pantau Rekapitulasi Realisasi pada link " + linkSRR + " dan Laporan Pertanggungjawaban (Accontability Report) pada folder link " + linkAR + "\n\n" +
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
        Utilities.sleep(3000);
        var response = UrlFetchApp.fetch(urlFonnte, options);
        sheet.getRange(rowNum, 2).setValue(startDateAR);
        sheet.getRange(rowNum, 3).setValue(endDateAR);
        Logger.log("Pesan terkirim ke " + phoneNumber + ": " + response.getContentText());

        sheet.getRange(rowNum, 11).setValue("SENT");
      } catch (error) {
        Logger.log("Error saat mengirim pesan: " + error);
      }

      Utilities.sleep(10 * 2000);
    }

    Logger.log("Selesai memproses Queue sementara.");
  } finally {
    lock.releaseLock();
  }
}

function formatRupiah(angka) {
  // Menggunakan toLocaleString untuk format ribuan
  return angka.toLocaleString('id-ID');
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