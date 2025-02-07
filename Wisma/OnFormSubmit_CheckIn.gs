function onFormSubmit_checkIn(e) {
  if (!e) {
    Logger.log("Tidak ada objek event.");
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var guestCISheet = ss.getSheetByName("GUESTCI");
    var roomPlotSheet = ss.getSheetByName("ROOMPLOT");

    if (!guestCISheet || !roomPlotSheet) {
      Logger.log("Sheet 'GUESTCI' atau 'ROOMPLOT' tidak ditemukan.");
      return;
    }

    // Dapatkan data dari formulir
    var formGuestID = e.values[1]; // Asumsikan guestID ada di kolom kedua (indeks 1)
    Logger.log("GuestID dari formulir: " + formGuestID);

    // Cari guestID di sheet ROOMPLOT
    var roomPlotData = roomPlotSheet.getRange(2, 1, roomPlotSheet.getLastRow() - 1, 14).getValues(); // Kolom A-N
    var foundRow = null;

    for (var i = 0; i < roomPlotData.length; i++) {
      var row = roomPlotData[i];
      var guestIDInSheet = row[12]; // Kolom M (indeks 12)

      if (guestIDInSheet === formGuestID) {
        foundRow = row;
        break;
      }
    }

    if (!foundRow) {
      Logger.log("GuestID tidak ditemukan di sheet ROOMPLOT.");
      return;
    }

    // Ambil nomor telepon dari kolom E (indeks 4)
    var rawPhone = foundRow[4]; // Kolom E
    var phone1 = formatPhone(rawPhone);

    if (!phone1) {
      Logger.log("Nomor telepon tidak valid untuk GuestID: " + formGuestID);
      return;
    }

    // Ambil nomor Mutu dan JAR dari sheet Publikasi
    var publikasiSheet = ss.getSheetByName("Publikasi_CheckIn");

    // Kirim pesan "Terima kasih" dengan tambahan informasi kritik dan saran
    var urlFonnte = "https://api.fonnte.com/send";
    var tokenFonnte = publikasiSheet?.getRange('D5').getValue(); // Optional chaining

    if (!tokenFonnte) {
      Logger.log("Token Fonnte belum diisi.");
      return;
    }

    var message = publikasiSheet?.getRange('D15').getValue() ;

    var payload = {
      "target": phone1,
      "message": message
    };
    var options = {
      "method": "post",
      "headers": { "Authorization": tokenFonnte },
      "payload": payload,
      "muteHttpExceptions": true
    };

    try {
      Utilities.sleep(1000); // Jeda 1 detik sebelum mengirim pesan
      var response = UrlFetchApp.fetch(urlFonnte, options);
      var responseText = response.getContentText();
      Logger.log("Pesan terkirim ke " + phone1 + ": " + responseText);

      // Cari baris di sheet guestCI berdasarkan guestID
      var guestCIData = guestCISheet.getRange(2, 1, guestCISheet.getLastRow() - 1, 3).getValues(); // Kolom A-C
      var guestCIRowIndex = -1;

      for (var i = 0; i < guestCIData.length; i++) {
        var row = guestCIData[i];
        var guestIDInguestCI = row[1]; // Asumsikan guestID ada di kolom B (indeks 1)

        if (guestIDInguestCI === formGuestID) {
          guestCIRowIndex = i + 2; // Offset header dan index dimulai dari 0
          break;
        }
      }

      if (guestCIRowIndex === -1) {
        Logger.log("GuestID tidak ditemukan di sheet guestCI.");
        return;
      }

      // Tandai status di sheet guestCI sebagai SENT
      guestCISheet.getRange(guestCIRowIndex, 3).setValue("SENT"); // Asumsikan kolom status ada di kolom C
      Logger.log("Status di sheet guestCI diperbarui menjadi SENT untuk GuestID: " + formGuestID);
    } catch (error) {
      Logger.log("Error saat mengirim pesan: " + error);
    }
  } catch (error) {
    Logger.log("Error umum: " + error);
  }
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