function pdfPengembalian(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!e || !e.range) return Logger.log('e tidak ada');

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // Nama worksheet template
  const templateSheetName = "Template Pengembalian";
  const dBSheetName = "DB";
  // Ambil worksheet template
  const secondarySheet = spreadsheet.getSheetByName(templateSheetName);
  const dbSheet = spreadsheet.getSheetByName(dBSheetName);
  if (!secondarySheet || !dbSheet || sheetName !== 'LOGIN02') {
    return Logger.log(`Sheet tidak ditemukan.`);;
  }

  // Ambil data mentah dari event Google Form
  const formData = e.values;
  const dbValues = dbSheet.getDataRange().getValues();

  const dbHeaders = dbValues[1]; // Baris ke-2
  const dbNamaIdx = dbHeaders.indexOf('nama');
  const dbMerkTipeIdx = dbHeaders.indexOf('merk_tipe');
  const dbLabelIdx = dbHeaders.indexOf('label');
  const dbAvailabilityIdx = dbHeaders.indexOf('availabiity');
  const dbIdLogIdx = dbHeaders.indexOf('subGrp');
  const dbJenisIdx = dbHeaders.indexOf('jenis');

  if (dbNamaIdx === -1 || dbMerkTipeIdx === -1 || dbLabelIdx === -1 || dbAvailabilityIdx === -1) {
    SpreadsheetApp.getUi().alert('Pastikan header di sheet "DB" sesuai: nama, merk tipe, label, availability.');
    return;
  }

  // Buat Map untuk menyimpan kombinasi nama + merk tipe + label => availability
  const dbMapp = {};

  for (let i = 1; i < dbValues.length; i++) { // Mulai dari baris ke-3
    const row = dbValues[i];
    const nama = row[dbNamaIdx].toString().trim();
    const merkTipe = row[dbMerkTipeIdx].toString().trim();
    const label = row[dbLabelIdx].toString().trim();
    const availability = row[dbAvailabilityIdx].toString().trim().toUpperCase();
    const idLog = row[dbIdLogIdx].toString().trim();
    const jenis = row[dbJenisIdx].toString().trim();

    const key = `${nama} ${merkTipe}`;
    // Jika key belum ada, buat array baru
    if (!dbMapp[key]) {
      dbMapp[key] = {
        availability: availability,
        nama: nama,
        idLog: idLog,
        jenis: jenis,
        labels: [] // Buat array untuk menyimpan semua label
      };
    }

    if (availability === 'AVAILABLE') {
      // Tambahkan label ke array
      dbMapp[key].labels.push(label)
    }
  }

  // Cleaning map key untuk memudahkan pencocokan
  const dbSimpleMapp = {};
  for (const key in dbMapp) {
    // key format: "nama merkTipe label"
    const parts = key.split(' ');
    const merkTipe = parts.pop(); // Ambil merk tipe
    const nama = parts.join(' '); // Sisa adalah nama
    let simpleKey = null;
    if (merkTipe) {
      simpleKey = `${nama} ${merkTipe}`;
    } else {
      simpleKey = `${nama}`;
    }

    if (!dbSimpleMapp[simpleKey]) {
      dbSimpleMapp[simpleKey] = [];
    }
    dbSimpleMapp[simpleKey].push(dbMapp[key]);
  }

  Logger.log(`Bentuk hasil dari form: "${formData}".`);
  // const timestamp = formData[0];
  // const tujuan = formData[1];
  // const namaIdentitas = formData[2];
  // const lokasiPenyimpanan = formData[3];
  const alat = formData[4];
  const nomorTelp = formData[5];

  Logger.log(`Nomor Telp: "${nomorTelp}".`);

  const groupMap = groupMainInput2(alat);

  const peralatan = [];
  const labelPeralatan = [];
  const apd = [];
  const labelAPD = [];
  const idTransaksiPeralatan = [];
  const idTransaksiAPD = [];
  const idTransaksi = [];

  // Logging hasil untuk setiap entri
  groupMap.forEach(entry => {
    const simpleKey = `${entry.mainInput}`; // Gunakan rowKey sebagai simpleKey

    Logger.log("Bentuk mainInput: " + entry.mainInput);
    Logger.log("Bentuk label: " + entry.labels.join('; '));
    Logger.log("Bentuk idTransaksi: " + entry.idTransaksi);

    // Periksa apakah key ada di dbSimpleMap
    if (dbSimpleMapp[simpleKey]) {
      // Filter hanya entry dengan availability 'AVAILABLE'
      const availableEntries = dbSimpleMapp[simpleKey];

      // const availableEntrie = dbSimpleMapp[simpleKey].filter(entry => entry.availability === 'AVAILABLE');

      // Hitung total labels dari semua availableEntries
      const totalLabels = availableEntries.reduce((sum, entry) => sum + entry.labels.length, 0);

      const namaItemAPD = availableEntries.reduce((sum, entry) => entry.nama, '').replace(/\s/g, '');

      // Ambil nilai 'jenis' (jika ada)
      const jenis = availableEntries.length > 0 ? availableEntries[0].jenis : '';

      // Logger.log("total label: " + availableEntries[0].labels.length);
      Logger.log(`isi availableEntries: ${JSON.stringify(availableEntries)}`);

      if (jenis == 'AKB' || jenis == 'ALT') {
        peralatan.push(entry.mainInput);
        labelPeralatan.push(entry.labels.join('; '));
        idTransaksiPeralatan.push(entry.idTransaksi);
      }

      if (jenis == 'APD') {
        apd.push(entry.mainInput);
        labelAPD.push(entry.labels.join('; '));
        idTransaksiAPD.push(entry.idTransaksi);
        // Logger.log(`Isi APD: ${JSON.stringify(apd)}`);
        // Logger.log(`Isi label APD: ${JSON.stringify(labelAPD)}`);
      }



      // Kembalikan totalLabels
      return [totalLabels];
    } else {
      // Jika key tidak ditemukan, kembalikan jumlah label sebagai 0
      // return [rowNumber, 0, ''];
    }
  });
  idTransaksi.push(...idTransaksiAPD, ...idTransaksiPeralatan);

  const dataLogout = getLogoutData(idTransaksi);

  Logger.log("Data nomor permintaan: "+dataLogout.nomorPermintaan);
  

  const countsPerStringPeralatan = labelPeralatan.map(label => countLabelsInString(label));
  const countsPerStringAPD = labelAPD.map(label => countLabelsInString(label));

  // Menentukan nama untuk sheet baru
  const timestamps = new Date().toISOString().replace(/[-:.TZ]/g, ""); // Contoh: 20231119T123456
  const newSheetName = `Response_${timestamps}`;

  // Menyalin sheet template
  const newSheet = secondarySheet.copyTo(spreadsheet);

  // Mengubah nama sheet baru
  newSheet.setName(newSheetName);

  // Pindahkan sheet baru ke posisi terakhir
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());

  // Menentukan baris mulai dan kolom APD
  const startRowAPD = 30; // Mulai dari baris ke-25

  if (apd.length > 0) {
    const writeAPDWithNumber = apd.map((item, index) => [index + 1, item]);
    const writeLabelAPD = countsPerStringAPD.map((row) => [row]);
    const writeKeterangan = idTransaksiAPD.map((row) => [row]);

    const labelCounts = writeAPDWithNumber.map(([rowNumber, rowKey]) => {
      const simpleKey = `${rowKey}`; // Gunakan rowKey sebagai simpleKey
      Logger.log("simple key dari APD: " + simpleKey);

      // Periksa apakah key ada di dbSimpleMap
      if (dbSimpleMapp[simpleKey]) {
        // Filter hanya entry dengan availability 'AVAILABLE'
        // const availableEntrie = dbSimpleMapp[simpleKey].filter(entry => entry.availability === 'AVAILABLE');

        const availableEntrie = dbSimpleMapp[simpleKey];

        // Hitung total labels dari semua availableEntries
        const totalLabels = availableEntrie.reduce((sum, entry) => sum + entry.labels.length, 0);

        // const namaItemAPD = availableEntrie.reduce((sum, entry) => entry.nama, '').replace(/\s/g, '');

        // let jenisItem = '';

        Logger.log("total label: " + totalLabels);
        Logger.log(`Isi availableEntrie: ${JSON.stringify(availableEntrie)}`);

        // Kembalikan rowNumber dan totalLabels
        return [rowNumber, totalLabels];
      } else {
        // Jika key tidak ditemukan, kembalikan jumlah label sebagai 0
        return [rowNumber, 0];
      }
    });

    newSheet.insertRows(startRowAPD, apd.length);

    // Menulis APD dengan nomor
    const writeRange = newSheet.getRange(startRowAPD, 2, writeAPDWithNumber.length, 2); // (baris, kolom, jumlah baris, jumlah kolom)
    writeRange.setValues(writeAPDWithNumber);

    // b. Mengatur alignment untuk Kolom C (data) ke kiri
    const dataRange = writeRange.offset(0, 1, writeAPDWithNumber.length, 1); // Kolom C
    dataRange.setHorizontalAlignment("left");

    // Jumlah Total
    const startRowDikembalikan = startRowAPD; // Baris mulai untuk dikembalikan
    const numRowsDikembalikan = writeAPDWithNumber.length; // Jumlah baris untuk dikembalikan
    const dikembalikanColumn = 7; // Kolom G adalah kolom ke-7

    // Menentukan rentang untuk checkmark
    const dikembalikanRange = newSheet.getRange(startRowDikembalikan, dikembalikanColumn, numRowsDikembalikan, 1);
    dikembalikanRange.setValues(writeLabelAPD);

    // 6. Menambahkan Simbol Checkmark di Kolom I
    const startRowCheckmark = startRowAPD; // Baris mulai untuk checkmark
    const numRowsCheckmark = writeAPDWithNumber.length; // Jumlah baris untuk checkmark
    const checkmarkColumn = 9; // Kolom I adalah kolom ke-9

    // Menentukan rentang untuk checkmark
    const checkmarkRange = newSheet.getRange(startRowCheckmark, checkmarkColumn, numRowsCheckmark, 1);

    // Membuat array dengan simbol checkmark
    const checkmarkValues = writeAPDWithNumber.map(() => ['✓']); // Atau simbol lain sesuai kebutuhan

    // Menulis simbol checkmark ke rentang
    checkmarkRange.setValues(checkmarkValues);

    // (Opsional) Mengatur alignment ke tengah untuk checkmark
    checkmarkRange.setHorizontalAlignment("center");

    // Jumlah Total
    const startRowKeterangan = startRowAPD; // Baris mulai untuk keterangan
    const numRowsKeterangan = writeAPDWithNumber.length; // Jumlah baris untuk keterangan
    const keteranganColumn = 11; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const keteranganRange = newSheet.getRange(startRowKeterangan, keteranganColumn, numRowsKeterangan, 1);
    keteranganRange.setWrap(true);
    keteranganRange.setValues(writeKeterangan);

    keteranganRange.setHorizontalAlignment("left");
    newSheet.autoResizeColumn(keteranganColumn);

    // Mengatur tinggi baris berdasarkan panjang karakter
    const keteranganValues = keteranganRange.getValues(); // Ambil nilai dari range keterangan
    for (let i = 0; i < keteranganValues.length; i++) {
      const keterangan = keteranganValues[i][0]; // Nilai keterangan pada baris i
      const rowIndex = startRowKeterangan + i; // Baris aktual di sheet
      if (keterangan.length > 27) {
        newSheet.setRowHeight(rowIndex, 30); // Atur tinggi baris menjadi 30 pixel jika panjang karakter > 27
      }
    }

    // Jumlah Total
    const startRowJumlahStock = startRowAPD; // Baris mulai untuk jumlah stock
    const numRowsJumlahStock = writeAPDWithNumber.length; // Jumlah baris untuk jumlah stock
    const jumlahStockColumn = 5; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const jumlahStockRange = newSheet.getRange(startRowJumlahStock, jumlahStockColumn, numRowsJumlahStock, 1);
    jumlahStockRange.setValues(labelCounts.map(([_, totalLabels]) => [totalLabels]));

    // Membuat merged cells pada kolom E dan F untuk setiap baris yang baru disisipkan
    for (let i = 0; i < apd.length; i++) {
      const currentRow = startRowAPD + i;
      const mergeRange = newSheet.getRange(currentRow, 5, 1, 2); // Kolom E (5) dan F
      mergeRange.merge();

      const mergeCD = newSheet.getRange(currentRow, 3, 1, 2); // Kolom C (3) dan D
      mergeCD.merge();

      // Menggabungkan kolom G dan H
      const mergeRangeGH = newSheet.getRange(currentRow, 7, 1, 2); // Kolom G (7) dan H
      mergeRangeGH.merge();
    }
  }

  newSheet.hideRows(29);

  // Menentukan baris mulai dan kolom Peralatan
  const startRowPeralatan = 24; // Mulai dari baris ke-25
  if (peralatan.length > 0) {
    const writePeralatanWithNumber = peralatan.map((item, index) => [index + 1, item]);
    const writeLabelPeralatan = countsPerStringPeralatan.map((row) => [row]);
    const writeKeterangan = idTransaksiPeralatan.map((row) => [row]);

    const labelCounts = writePeralatanWithNumber.map(([rowNumber, rowKey]) => {
      const simpleKey = `${rowKey}`; // Gunakan rowKey sebagai simpleKey

      // Periksa apakah key ada di dbSimpleMap
      if (dbSimpleMapp[simpleKey]) {
        // Filter hanya entry dengan availability 'AVAILABLE'
        // const availableEntrie = dbSimpleMapp[simpleKey].filter(entry => entry.availability === 'AVAILABLE');

        const availableEntrie = dbSimpleMapp[simpleKey];

        // Hitung total labels dari semua availableEntries
        const totalLabels = availableEntrie.reduce((sum, entry) => sum + entry.labels.length, 0);

        const namaItemAPD = availableEntrie.reduce((sum, entry) => entry.nama, '').replace(/\s/g, '');

        let jenisItem = '';

        Logger.log("total label: " + totalLabels);
        Logger.log(`Isi availableEntrie: ${JSON.stringify(availableEntrie)}`);

        // Kembalikan rowNumber dan totalLabels
        return [rowNumber, totalLabels];
      } else {
        // Jika key tidak ditemukan, kembalikan jumlah label sebagai 0
        return [rowNumber, 0];
      }
    });

    newSheet.insertRows(startRowPeralatan, peralatan.length);

    // Menulis peralatan dengan nomor
    const writeRange = newSheet.getRange(startRowPeralatan, 2, writePeralatanWithNumber.length, 2); // (baris, kolom, jumlah baris, jumlah kolom)
    writeRange.setValues(writePeralatanWithNumber);

    // b. Mengatur alignment untuk Kolom C (data) ke kiri
    const dataRange = writeRange.offset(0, 1, writePeralatanWithNumber.length, 1); // Kolom C
    dataRange.setHorizontalAlignment("left");

    // Jumlah Total
    const startRowDisiapkan = startRowPeralatan; // Baris mulai untuk disiapkan
    const numRowsDisiapkan = writePeralatanWithNumber.length; // Jumlah baris untuk disiapkan
    const disiapkanColumn = 7; // Kolom G adalah kolom ke-7

    // Menentukan rentang untuk checkmark
    const disiapkanRange = newSheet.getRange(startRowDisiapkan, disiapkanColumn, numRowsDisiapkan, 1);
    disiapkanRange.setValues(writeLabelPeralatan);

    // 6. Menambahkan Simbol Checkmark di Kolom I
    const startRowCheckmark = startRowPeralatan; // Baris mulai untuk checkmark
    const numRowsCheckmark = writePeralatanWithNumber.length; // Jumlah baris untuk checkmark
    const checkmarkColumn = 9; // Kolom I adalah kolom ke-9

    // Menentukan rentang untuk checkmark
    const checkmarkRange = newSheet.getRange(startRowCheckmark, checkmarkColumn, numRowsCheckmark, 1);

    // Membuat array dengan simbol checkmark
    const checkmarkValues = writePeralatanWithNumber.map(() => ['✓']); // Atau simbol lain sesuai kebutuhan

    // Menulis simbol checkmark ke rentang
    checkmarkRange.setValues(checkmarkValues);

    // (Opsional) Mengatur alignment ke tengah untuk checkmark
    checkmarkRange.setHorizontalAlignment("center");

    // Jumlah Total
    const startRowKeterangan = startRowPeralatan; // Baris mulai untuk keterangan
    const numRowsKeterangan = writePeralatanWithNumber.length; // Jumlah baris untuk keterangan
    const keteranganColumn = 11; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const keteranganRange = newSheet.getRange(startRowKeterangan, keteranganColumn, numRowsKeterangan, 1);
    keteranganRange.setWrap(true);
    keteranganRange.setValues(writeKeterangan);

    keteranganRange.setHorizontalAlignment("left");
    newSheet.autoResizeColumn(keteranganColumn);

    // Mengatur tinggi baris berdasarkan panjang karakter
    const keteranganValues = keteranganRange.getValues(); // Ambil nilai dari range keterangan
    for (let i = 0; i < keteranganValues.length; i++) {
      const keterangan = keteranganValues[i][0]; // Nilai keterangan pada baris i
      const rowIndex = startRowKeterangan + i; // Baris aktual di sheet
      if (keterangan.length > 27) {
        newSheet.setRowHeight(rowIndex, 30); // Atur tinggi baris menjadi 30 pixel jika panjang karakter > 27
      }
    }

    // Jumlah Total
    const startRowJumlahStock = startRowPeralatan; // Baris mulai untuk stock
    const numRowsJumlahStock = writePeralatanWithNumber.length; // Jumlah baris untuk stock
    const jumlahStockColumn = 5; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const jumlahStockRange = newSheet.getRange(startRowJumlahStock, jumlahStockColumn, numRowsJumlahStock, 1);
    jumlahStockRange.setValues(labelCounts.map(([_, totalLabels]) => [totalLabels]));

    // Membuat merged cells pada kolom E dan F untuk setiap baris yang baru disisipkan
    for (let i = 0; i < peralatan.length; i++) {
      const currentRow = startRowPeralatan + i;
      const mergeRange = newSheet.getRange(currentRow, 5, 1, 2); // Kolom E (5) dan F (6)
      mergeRange.merge();

      const mergeCD = newSheet.getRange(currentRow, 3, 1, 2); // Kolom C (3) dan D (4)
      mergeCD.merge();

      // Menggabungkan kolom G dan H
      const mergeRangeGH = newSheet.getRange(currentRow, 7, 1, 2); // Kolom G (7) dan H (8)
      mergeRangeGH.merge();
    }
  }

  newSheet.hideRows(23);

  newSheet.getRange(9, 4).setValue(dataLogout.pembelajaran);
  newSheet.getRange(13, 4).setValue(dataLogout.tanggalSelesai);
  newSheet.getRange(12, 4).setValue(dataLogout.tanggalMulai);
  newSheet.getRange(19, 11).setValue(dataLogout.idTransaction); //Nomor permintaan

  const namePdf = `${dataLogout.transaction}`;

  const inputNumber = formatPhone(nomorTelp);

  // Generate PDF dari sheet "Template"
  let linkPDF = generatePDFPengembalian(newSheetName, namePdf);
  let pesanLink =  "📩 *Inventarisasi Masuk Laboratorium* \n dengan *ID "+ (dataLogout.transaction || 'Data tidak tersedia') + "*\n\n" +
        "📜 *Formulir Masuk:* " + (linkPDF || 'Data tidak tersedia') + "\n\n" +
        "sent by: Auto Report System";

  if (pesanLink) {
    Logger.log("Pesan link berhasil dibuat. Nilai pesanLink: " + pesanLink);
    senderMessage(pesanLink, inputNumber);
  } else {
    Logger.log("Pesan link tidak berhasil dibuat. Nilai pesanLink: " + pesanLink);
  }

  // Menghapus sheet sementara setelah PDF berhasil dibuat
  // Pastikan sheet yang akan dihapus bukan sheet template
  if (spreadsheet.getSheetByName(newSheetName)) {
    spreadsheet.deleteSheet(newSheet);
    Logger.log(`Sheet "${newSheetName}" telah dihapus setelah PDF dibuat.`);
  }
}

/**
 * Fungsi untuk menghasilkan PDF dari sheet tertentu
 * @param {string} sheetName - Nama sheet yang akan dikonversi menjadi PDF
 * @param {string} pdfName - Nama file PDF yang akan disimpan
 */
function generatePDFPengembalian(sheetName, pdfName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet dengan nama "${sheetName}" tidak ditemukan.`);
    return;
  }

  // Force refresh worksheet
  SpreadsheetApp.flush();  // Memastikan semua perubahan diterapkan sebelum lanjut

  const spreadsheetId = spreadsheet.getId();
  const sheetId = sheet.getSheetId();

  // URL dasar untuk mengonversi sheet ke PDF
  const url_base = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?';

  const url_ext = 'exportFormat=pdf&format=pdf' + // Format file
    '&size=letter' + // Ukuran kertas (A4, letter, etc.)
    '&portrait=true' + // Orientasi: true = potrait, false = landscape
    '&fitw=true' + // Fit to width
    '&top_margin=0.1' +
    '&bottom_margin=0.1' +
    '&left_margin=0' +
    '&right_margin=0.5' +
    `&repeatrows=1:17` + // Mengulang baris 1 dan 2 sebagai header di setiap halaman
    '&sheetnames=false&printtitle=false&pagenumbers=false' + // Tidak menyertakan sheet names, title, page numbers
    '&gridlines=false' + // Tidak menyertakan gridlines
    '&gid=' + sheetId; // ID sheet yang akan dikonversi

  const options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };

  // Force refresh worksheet
  SpreadsheetApp.flush();  // Memastikan semua perubahan diterapkan sebelum lanjut

  // Mengambil blob PDF
  const response = UrlFetchApp.fetch(url_base + url_ext, options);
  const blob = response.getBlob().setName(pdfName + '_In.pdf');

  // Menyimpan PDF ke folder tertentu di Drive (opsional)
  const folder = DriveApp.getFolderById('1NJWujNEBwh8YRh72QioRpjQWYbIBROE4'); // Ganti dengan ID folder tujuan
  const file = folder.createFile(blob);

  // Menyimpan PDF di root Drive
  const fileUrl = file.getUrl();

  Logger.log(`PDF "${pdfName}.pdf" berhasil dibuat dan disimpan di Google Drive dengan link "${fileUrl}"`);

  return `${fileUrl}`;
}

function countLabelsInString(labelString) {
  if (!labelString || typeof labelString !== 'string') {
    return 0;
  }

  // Pisahkan string berdasarkan ";" atau "," dan hitung elemen yang valid
  const labels = labelString.split(/[,;]/).map(label => label.trim()).filter(label => label !== '');

  return labels.length;
}

function groupMainInput2(input) {
  const groupMap = [];

  if (!input || typeof input !== 'string') {
    return groupMap;
  }

  const items = input.split(',')
    .map(item => item.trim())
    .filter(item => item !== "");

  items.forEach(item => {
    // Pecahkan bagian sebelum "#" dan setelah "- "
    const nameMatch = item.match(/^[^#]+/); // Nama barang (sebelum "#")
    const labelMatch = item.match(/#(\d+)/); // Label (angka setelah "#")
    const idMatch = item.match(/-\s(.+)/); // ID transaksi (setelah "- ")

    const mainInput = nameMatch ? nameMatch[0].trim() : "";
    const label = labelMatch ? labelMatch[1].trim() : "";
    const idTransaksi = idMatch ? idMatch[1].trim() : "";

    if (mainInput && label && idTransaksi) {
      // Cari entri dengan mainInput dan idTransaksi yang sama
      let existingEntry = groupMap.find(entry => entry.mainInput === mainInput && entry.idTransaksi === idTransaksi);

      if (existingEntry) {
        // Jika ada, tambahkan label ke entri tersebut
        existingEntry.labels.push(label);
      } else {
        // Jika tidak ada, buat entri baru
        groupMap.push({
          mainInput: mainInput,
          labels: [label],
          idTransaksi: idTransaksi
        });
      }
    }
  });

  return groupMap;
}

function itemTerbanyak(arr) {
  const hitungan = {};

  // Hitung kemunculan setiap item
  for (const item of arr) {
    hitungan[item] = (hitungan[item] || 0) + 1;
  }

  let itemTerbanyak = null;
  let jumlahTerbanyak = 0;

  // Cari item dengan kemunculan terbanyak
  for (const item in hitungan) {
    if (hitungan[item] > jumlahTerbanyak) {
      itemTerbanyak = parseInt(item); // Pastikan tipenya angka jika array berisi angka
      jumlahTerbanyak = hitungan[item];
    }
  }

  return itemTerbanyak;
}

function getLogoutData(idTransaksi) {
  // Validasi input
  if (!idTransaksi || !Array.isArray(idTransaksi) || idTransaksi.length === 0) {
    Logger.log('Warning: idTransaksi is empty or invalid');
    throw new Error('ID Transaksi tidak valid atau kosong');
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logoutSheetName = "LOGOUT02";
  const logoutSheet = spreadsheet.getSheetByName(logoutSheetName);
  
  if (!logoutSheet) {
    throw new Error('Sheet LOGOUT02 tidak ditemukan');
  }

  const data = {};
  
  // Mengambil data dari sheet
  const headerIdTransaksi = logoutSheet.getRange("P1:P" + logoutSheet.getLastRow()).getValues();
  const headerNomorPermintaan = logoutSheet.getRange("B1:B" + logoutSheet.getLastRow()).getValues();
  const headerTanggalMulai = logoutSheet.getRange("K1:K" + logoutSheet.getLastRow()).getValues();
  const headerTanggalSelesai = logoutSheet.getRange("L1:L" + logoutSheet.getLastRow()).getValues();
  const headerPembelajaran = logoutSheet.getRange("S1:S" + logoutSheet.getLastRow()).getValues();

  const nomorPermintaan = [];

  // Mencari nomor permintaan untuk setiap ID transaksi
  for (let i = 0; i < idTransaksi.length; i++) {
    const indeks = headerIdTransaksi.flat().indexOf(idTransaksi[i]);
    if (indeks !== -1) {
      nomorPermintaan.push(headerNomorPermintaan[indeks][0]);
    }
  }

  // Validasi hasil pencarian
  if (nomorPermintaan.length === 0) {
    throw new Error('Tidak ada nomor permintaan yang ditemukan untuk ID transaksi yang diberikan');
  }

  Logger.log(`Nomor Permintaan: ${JSON.stringify(nomorPermintaan)}`);

  const terbanyak = itemTerbanyak(nomorPermintaan);
  if (!terbanyak) {
    throw new Error('Tidak dapat menentukan nomor permintaan terbanyak');
  }

  const indeksTerbanyak = headerNomorPermintaan.flat().indexOf(terbanyak);
  if (indeksTerbanyak === -1) {
    throw new Error('Data tidak ditemukan untuk nomor permintaan terbanyak');
  }

  // Mengambil data terkait
  const tanggalSelesai = headerTanggalSelesai[indeksTerbanyak][0];
  const tanggalMulai = headerTanggalMulai[indeksTerbanyak][0];
  const pembelajaran = headerPembelajaran[indeksTerbanyak][0];

  if (!tanggalSelesai) {
    throw new Error('Tanggal selesai tidak ditemukan');
  }

  // Format tanggal
  try {
    const [day, month, year] = tanggalSelesai.split('/');
    const formattedDate = `${year}-${month}-${day}`;
    const tanggalFormatted = new Date(formattedDate).toISOString().split('T')[0].replace(/-/g, '');

    // Menyiapkan data return
    data.nomorPermintaan = terbanyak;
    data.tanggalFormated = tanggalFormatted;
    data.idTransaction = `Id: ${tanggalFormatted}${terbanyak}`;
    data.transaction = `${tanggalFormatted}${terbanyak}`;
    data.tanggalMulai = tanggalMulai;
    data.tanggalSelesai = tanggalSelesai;
    data.pembelajaran = pembelajaran;

    return data;
  } catch (e) {
    throw new Error(`Error saat memformat tanggal: ${e.message}`);
  }
}