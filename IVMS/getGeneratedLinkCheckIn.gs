function getLink(e) {
  const SHEET_NAME = "ivmsCHECKINFORM"; // Validate this sheet name
  const MAX_RETRIES = 10;
  const RETRY_DELAY = 10000; // 10 seconds
  const HEADER_ROW = 1; // Baris yang berisi header

  // Get the sheet and validate its name
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const formData = e.values;
  const pictureLink = formData[14];
  let inputNumber = [];

  Logger.log("formData[16]: " + formData[16]);
  Logger.log("formData: " + formData);

  if (sheetName !== SHEET_NAME) {
    Logger.log("Skipping execution: Sheet name is not '" + SHEET_NAME + "'.");
    return; // Exit the script
  }

  // Ambil semua header pada baris pertama
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

  const namaKolomURL = "Merged Doc URL - VMS_PMK";

  // Cari indeks kolom dengan nama header yang diinginkan (misalnya 'AE' atau yang relevan)
  const columnIndex = headers.indexOf(namaKolomURL) + 1; // Menyesuaikan agar menjadi 1-based index

  if (columnIndex === 0) {
    Logger.log("Header " + (namaKolomURL) + " tidak ditemukan.");
    return;
  }

  if (formData[1] == "Eksternal - Non Diklat") {
    const row = e.range.getRow();
    let retryCount = 0;
    let linkPDF = "";

    while (retryCount < MAX_RETRIES) {
      linkPDF = sheet.getRange(row, columnIndex).getValue();

      if (linkPDF) {
        Logger.log("Value in column AE: " + linkPDF);
        Logger.log("Value linkPicture: " + pictureLink);
        const fileId = extractFileIdFromUrl(pictureLink);
        Logger.log("Value fileId: " + fileId);

        const gcsImageUrl = uploadImageToGCS(fileId);
        inputNumber.push(formData[15]);
        // inputNumber.push("6282251111220-1600942098@g.us");

        let pesan = "📩 *Laporan Pengunjung Masuk* \n" +
          "Pengunjung dari *" + (formData[1] || 'Tidak diketahui') + "* dengan informasi: \n\n" +
          "🪪 *Nama:* " + (formData[3] || 'Tidak diketahui') + "\n" +
          "*Menjumpai:* " + (formData[6] || 'Tidak diketahui') + "\n" +
          "*Keperluan:* " + (formData[7] || 'Tidak diketahui') + "\n" +
          "*Jumlah Rombongan:* " + (formData[8] || 'Tidak diketahui') + "\n" +
          "*Zonasi:* " + (formData[13] || 'Tidak diketahui') + "\n\n" +
          "📜 *Detail Pengunjung:* " + (linkPDF || "Detail pengunjung tidak tersedia") + "\n\n" +
          "🗾 *Peta Zonasi:* https://drive.google.com/drive/folders/12SNVMJpaOQwuTZf41hOzwVJA7L3QR0PZ" + "\n\n" +
          "sent by: Auto Report System";

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const publikasiSheet = ss.getSheetByName('Publikasi_VMS_ExND');

        //Get data from sheet publikasi
        const senderToken = publikasiSheet.getRange(5, 4).getValue();
        const destinationNumber = publikasiSheet.getRange(8, 4).getValue();
        const statusScript = publikasiSheet.getRange(3, 2).getValue();

        Logger.log("Hasil dari senderToken: " + senderToken);
        Logger.log("Hasil dari destinationNumber: " + destinationNumber);
        Logger.log("Hasil dari statusScript: " + statusScript);

        if (statusScript == "Aktif") {
          inputNumber.push(destinationNumber);

          senderMessage(pesan, inputNumber, senderToken, gcsImageUrl);
        }

        // Extract filename from GCS URL (you might need a more robust way to parse URLs in production)
        var gcsFilename = gcsImageUrl.substring(gcsImageUrl.lastIndexOf('/') + 1);

        var deletionSuccessful = deleteImageFromGCS(gcsFilename);

        if (deletionSuccessful) {
          Logger.log("Successfully deleted image from GCS: " + gcsFilename);
        } else {
          Logger.log("Failed to delete image from GCS: " + gcsFilename);
        }
        return;
      } else {
        Logger.log("Cell AE" + row + " is empty. Retrying...");

        SpreadsheetApp.getActiveSpreadsheet().toast("Retrying AE" + row + "...");

        // Ensure all pending changes are applied
        SpreadsheetApp.flush();

        Utilities.sleep(RETRY_DELAY);
        retryCount++;
      }
    }

    Logger.log("No value found after " + MAX_RETRIES + " attempts.");
  } else if (formData[1] == "Eksternal - Diklat") {
    // inputNumber.push("08115153030");

    let pesan = "📩 *Laporan Pengunjung Masuk* \n" +
      "Pengunjung dari *" + (formData[1] || 'Tidak diketahui') + "* dengan informasi: \n\n" +
      "🪪 *Nama:* " + "\n" +
      (diklatGuestSplitting(formData[16])) +
      "*Menjumpai / Keperluan:* Kelas Pembelajaran" + "\n" +
      "*Zonasi:* Biru" + "\n\n" +
      "📜 *Formulir Bertamu:* -" + "\n\n" +
      "🗾 *Peta Zonasi:* https://drive.google.com/drive/folders/12SNVMJpaOQwuTZf41hOzwVJA7L3QR0PZ" + "\n\n" +
      "sent by: Auto Report System";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const publikasiSheet = ss.getSheetByName('Publikasi_VMS_ExD');

    //Get data from sheet publikasi
    const senderToken = publikasiSheet.getRange(5, 4).getValue();
    const destinationNumber = publikasiSheet.getRange(8, 4).getValue();
    const statusScript = publikasiSheet.getRange(3, 2).getValue();

    Logger.log("Hasil dari senderToken: " + senderToken);
    Logger.log("Hasil dari destinationNumber: " + destinationNumber);
    Logger.log("Hasil dari statusScript: " + statusScript);

    if (statusScript == "Aktif") {
      inputNumber.push(destinationNumber);

      senderMessage(pesan, inputNumber, senderToken);
    }
    return;
  } else if (formData[1] == "Eksternal - Khusus") {

    let pesan = "📩 *Laporan Pengunjung Masuk* \n" +
      "Pengunjung dari *" + (formData[1] || 'Tidak diketahui') + "* dengan informasi: \n\n" +
      "🪪 *Nama:* " + (formData[20] || 'Tidak diketahui') + "\n" +
      "*Menjumpai / Keperluan:* " + (formData[21] || "Keperluan tidak berhasil didapatkan") + "\n\n" +
      "sent by: Auto Report System";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const publikasiSheet = ss.getSheetByName('Publikasi_VMS_ExKS');

    //Get data from sheet publikasi
    const senderToken = publikasiSheet.getRange(5, 4).getValue();
    const destinationNumber = publikasiSheet.getRange(8, 4).getValue();
    const statusScript = publikasiSheet.getRange(3, 2).getValue();

    Logger.log("Hasil dari senderToken: " + senderToken);
    Logger.log("Hasil dari destinationNumber: " + destinationNumber);
    Logger.log("Hasil dari statusScript: " + statusScript);
    
    if (statusScript == "Aktif") {
      inputNumber.push(destinationNumber);

      senderMessage(pesan, inputNumber, senderToken);
    }
    return;
  } else {
    // inputNumber.push("08994334111");
    // inputNumber.push("089637137078");

    let pesan = "📩 *Laporan Pengunjung Masuk* \n" +
      "Pengunjung dari *" + (formData[1] || 'Tidak diketahui') + "* dengan informasi: \n\n" +
      "🪪 *Nama:* " + "\n" +
      (diklatGuestSplitting(formData[17])) +
      "*Menjumpai / Keperluan:* Kelas Pembelajaran" + "\n" +
      "*Zonasi:* Biru" + "\n\n" +
      "📜 *Formulir Bertamu:* -" + "\n\n" +
      "🗾 *Peta Zonasi:* https://drive.google.com/drive/folders/12SNVMJpaOQwuTZf41hOzwVJA7L3QR0PZ" + "\n\n" +
      "sent by: Auto Report System";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const publikasiSheet = ss.getSheetByName('Publikasi_VMS_ExKS');

    //Get data from sheet publikasi
    const senderToken = publikasiSheet.getRange(5, 4).getValue();
    const destinationNumber = publikasiSheet.getRange(8, 4).getValue();
    const statusScript = publikasiSheet.getRange(3, 2).getValue();

    Logger.log("Hasil dari senderToken: " + senderToken);
    Logger.log("Hasil dari destinationNumber: " + destinationNumber);
    Logger.log("Hasil dari statusScript: " + statusScript);

    if (statusScript == "Aktif") {
      inputNumber.push(destinationNumber);

      senderMessage(pesan, inputNumber, senderToken);
    }
    return;
  }
}

function senderMessage(pesan, inputNumber, senderNumber, urlPicture) {
  var tokenFonnte = senderNumber;
  var url = "https://api.fonnte.com/send";

  var delayMilliseconds = 5000; // Set the delay duration in milliseconds (e.g., 5000ms = 5 seconds)

  for (i = 0; i < inputNumber.length; i++) {
    const number = formatPhone(inputNumber[i]);

    if (urlPicture) {
      var options = {
        "method": "post",
        "headers": {
          "Authorization": tokenFonnte
        },
        "payload": {
          "target": number,
          "message": pesan,
          "url": urlPicture,
        }
      };
    } else {
      var options = {
        "method": "post",
        "headers": {
          "Authorization": tokenFonnte
        },
        "payload": {
          "target": number,
          "message": pesan,
        }
      };
    }

    // Kirim pesan
    try {
      var response = UrlFetchApp.fetch(url, options);
      Logger.log("Pesan berhasil dikirim ke: " + number);
      Logger.log("Response: " + response.getContentText());
    } catch (e) {
      Logger.log("Gagal mengirim pesan ke: " + number);
      Logger.log("Error: " + e.message);
    }

    Utilities.sleep(delayMilliseconds);
  }
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

function diklatGuestSplitting(formData) {
  let message = "";
  if (!formData) {
    return ""; // Return an empty array if formData or formData[16] is missing
  }

  const data16 = formData;
  if (typeof data16 !== 'string') {
    return ""; // Return an empty array if formData[16] is not a string
  }

  const formData16 = data16.split(',').map(item => item.trim());

  for (i = 0; i < formData16.length; i++) {
    message += "- " + (formData16[i]) + "\n";
  }

  return message;
}

function extractFileIdFromUrl(url) {
  let fileId = null;

  // Try the older "open?id=" format
  let match = url.match(/open\?id=([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    fileId = match[1];
    return fileId; // Return immediately if found
  }

  // Try the newer "file/d/" format if the older format is not matched
  match = url.match(/file\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    fileId = match[1];
    return fileId; //Return immediately if found
  }


  // If the regex from drive.google.com/uc?id= works, you can add that here

  return null; // Or return an error message, if needed
}