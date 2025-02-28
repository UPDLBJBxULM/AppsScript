function pdfPeminjaman(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!e || !e.range) return Logger.log('e tidak ada');

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // Nama worksheet template
  const templateSheetName = "Template Peminjaman";
  const dBSheetName = "DB";

  // Ambil worksheet template
  const secondarySheet = spreadsheet.getSheetByName(templateSheetName);
  const dbSheet = spreadsheet.getSheetByName(dBSheetName);
  if (!secondarySheet || !dbSheet || sheetName !== 'LOGOUT02') {
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
  const dbMap = {};

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
    if (!dbMap[key]) {
      dbMap[key] = {
        availability: availability,
        nama: nama,
        idLog: idLog,
        jenis: jenis,
        labels: [] // Buat array untuk menyimpan semua label
      };
    }

    if (availability === 'AVAILABLE') {
      // Tambahkan label ke array
      dbMap[key].labels.push(label)
    }
  }

  // Cleaning map key untuk memudahkan pencocokan
  const dbSimpleMap = {};
  for (const key in dbMap) {
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

    if (!dbSimpleMap[simpleKey]) {
      dbSimpleMap[simpleKey] = [];
    }
    dbSimpleMap[simpleKey].push(dbMap[key]);
  }

  // Mendapatkan nilai pada indeks tertentu (sesuaikan dengan kebutuhan Anda)
  Logger.log(`Hasil Formdata: ${JSON.stringify(formData)}`);
  const judulPembelajaran = formData[5];
  const tanggalMulai = formData[3];
  const tanggalSelesai = formData[4];
  const nomorPermintaan = formData[1];
  const alatKerjaBantu = formData[8];
  const alatUjidanUkur = formData[9];
  const alatPelindungDiri = formData[10];
  const material = formData[9];
  const nomorTelp = formData[17];

  Logger.log(`Nomor Telpon: "${nomorTelp}".`);
  // Pastikan tanggalSelesai diformat ke "YYYYMMDD"
  const [day, month, year] = tanggalSelesai.split('/'); // Pisahkan tanggal, bulan, tahun
  const formattedDate = `${year}-${month}-${day}`; // Susun kembali dalam format ISO

  // Pastikan tanggal diformat ke "YYYYMMDD"
  const tanggalFormatted = new Date(formattedDate).toISOString().split('T')[0].replace(/-/g, '');

  // Gabungkan kedua variabel
  const idTransaksi = `Id: ${tanggalFormatted}${nomorPermintaan}`;
  const idTransaction = `${tanggalFormatted}${nomorPermintaan}`;
  const namePdf = `${tanggalFormatted}${nomorPermintaan}`;

  const alatKerjaBantuArray = alatKerjaBantu ? alatKerjaBantu.split(',') : [];
  const alatUjidanUkurArray = alatUjidanUkur ? alatUjidanUkur.split(',') : [];

  const peralatan = [
    ...alatKerjaBantuArray,
    ...alatUjidanUkurArray
  ]
    .map(item => item.trim())
    .filter(Boolean);

  // Mengambil data yang sudah dikelompokkan
  const peralatanString = peralatan.join(', ');
  const groupPeralatan = groupMainInput(peralatanString);
  const groupAKB = groupMainInput(alatKerjaBantu);
  const groupAPD = groupMainInput(alatPelindungDiri);
  // const groupMAT = groupMainInput(material);

  // Mengambil kunci dari groupMap1
  const inputPeralatan = Object.keys(groupPeralatan);
  const inputAKB = Object.keys(groupAKB);

  Logger.log(`Panjang AKB: "${inputAKB.length}".`);

  const inputAPD = Object.keys(groupAPD);

  // Menyiapkan data untuk ditulis ke sheet (array dua dimensi)
  const writePeralatan = inputPeralatan.map(key => [key]);
  const labelPeralatan = inputPeralatan.map(key => groupPeralatan[key].join('; '));

  // Perbaikan di sini: Pass setiap label individu, bukan seluruh array
  const countsPerStringPeralatan = labelPeralatan.map(label => countLabelsInString(label));
  const totalPeralatan = inputPeralatan.length;

  const writeAPD = inputAPD.map(key => [key]);
  const labelAPD = inputAPD.map(key => groupAPD[key].join('; '));

  const countsPerStringAPD = labelAPD.map(label => countLabelsInString(label));

  Logger.log(`Item writePeralatan: "${writePeralatan}".`);
  Logger.log(`Label Peralatan: "${labelPeralatan}".`);
  Logger.log(`countsPerStringPeralatan: "${countsPerStringPeralatan}".`);
  Logger.log(`Total Peralatan: "${totalPeralatan}".`);


  // Menentukan nama untuk sheet baru
  const timestamp = new Date().toISOString().replace(/[-:.TZ]/g, ""); // Contoh: 20231119T123456
  const newSheetName = `Response_${timestamp}`;

  // Menyalin sheet template
  const newSheet = secondarySheet.copyTo(spreadsheet);

  // Mengubah nama sheet baru
  newSheet.setName(newSheetName);

  // Pindahkan sheet baru ke posisi terakhir
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());

  // Menentukan baris mulai dan kolom APD
  const startRowAPD = 30; // Mulai dari baris ke-25

  // Menuliskan data secara batch ke sheet dengan menyisipkan baris baru APD
  if (writeAPD.length > 0) {
    const writeAPDWithNumber = writeAPD.map((row, index) => [index + 1, row[0]]);
    const writeLabelAPD = countsPerStringAPD.map((row) => [row]);

    // Iterasi writeAPDWithNumber
    const labelCounts = writeAPDWithNumber.map(([rowNumber, rowKey]) => {
      const simpleKey = `${rowKey}`; // Gunakan rowKey sebagai simpleKey
      Logger.log("Ini Tereksekusi dengan panjang APD: " + writeAPD.length);
      // Periksa apakah key ada di dbSimpleMap
      if (dbSimpleMap[simpleKey]) {
        // Filter hanya entry dengan availability 'AVAILABLE'
        const availableEntry = dbSimpleMap[simpleKey];
        // const availableEntries = dbSimpleMap[simpleKey].filter(entry => entry.availability === 'AVAILABLE');

        // Hitung total labels dari semua availableEntries
        const totalLabels = availableEntry.reduce((sum, entry) => sum + entry.labels.length, 0);

        const namaItemAPD = availableEntry.reduce((sum, entry) => entry.nama, '').replace(/\s/g, '');


        const keterangan = tanggalFormatted + nomorPermintaan + "APD" + namaItemAPD;

        // Kembalikan rowNumber dan totalLabels
        return [rowNumber, totalLabels, keterangan];
      } else {
        // Jika key tidak ditemukan, kembalikan jumlah label sebagai 0
        return [rowNumber, 0, ''];
      }
    });

    Logger.log(`writeAPDWithNumber: "${writeAPDWithNumber}"`);

    // Menyisipkan baris baru
    newSheet.insertRows(startRowAPD, writeAPD.length);

    const writeRange = newSheet.getRange(startRowAPD, 2, writeAPDWithNumber.length, 2); // (baris, kolom, jumlah baris, jumlah kolom)
    writeRange.setValues(writeAPDWithNumber);

    // a. Mengatur alignment untuk Kolom B (nomor) ke tengah
    const nomorRange = writeRange.offset(0, 0, writeAPDWithNumber.length, 1); // Kolom B
    nomorRange.setHorizontalAlignment("center");

    // b. Mengatur alignment untuk Kolom C (data) ke kiri
    const dataRange = writeRange.offset(0, 1, writeAPDWithNumber.length, 1); // Kolom C
    dataRange.setHorizontalAlignment("left");

    // Jumlah Total
    const startRowDisiapkan = startRowAPD; // Baris mulai untuk disiapkan
    const numRowsDisiapkan = writeAPDWithNumber.length; // Jumlah baris untuk disiapkan
    const disiapkanColumn = 7; // Kolom G adalah kolom ke-7

    // Menentukan rentang untuk checkmark
    const disiapkanRange = newSheet.getRange(startRowDisiapkan, disiapkanColumn, numRowsDisiapkan, 1);
    disiapkanRange.setValues(writeLabelAPD);

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
    const startRowJumlahStock = startRowAPD; // Baris mulai untuk stock
    const numRowsJumlahStock = writeAPDWithNumber.length; // Jumlah baris untuk stock
    const jumlahStockColumn = 5; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const jumlahStockRange = newSheet.getRange(startRowJumlahStock, jumlahStockColumn, numRowsJumlahStock, 1);
    jumlahStockRange.setValues(labelCounts.map(([_, totalLabels, keterangan]) => [totalLabels]));

    // Jumlah Total
    const startRowKeterangan = startRowAPD; // Baris mulai untuk keterangan
    const numRowsKeterangan = writeAPDWithNumber.length; // Jumlah baris untuk keterangan
    const keteranganColumn = 11; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const keteranganRange = newSheet.getRange(startRowKeterangan, keteranganColumn, numRowsKeterangan, 1);
    keteranganRange.setValues(labelCounts.map(([_, totalLabels, keterangan]) => [keterangan]));

    keteranganRange.setHorizontalAlignment("left");

    // Mengatur tinggi baris berdasarkan panjang karakter
    const keteranganValues = keteranganRange.getValues(); // Ambil nilai dari range keterangan
    for (let i = 0; i < keteranganValues.length; i++) {
      const keterangan = keteranganValues[i][0]; // Nilai keterangan pada baris i
      const rowIndex = startRowKeterangan + i; // Baris aktual di sheet
      if (keterangan.length > 27) {
        newSheet.setRowHeight(rowIndex, 30); // Atur tinggi baris menjadi 30 pixel jika panjang karakter > 27
      }
    }

    // Membuat merged cells pada kolom E dan F untuk setiap baris yang baru disisipkan
    for (let i = 0; i < writeAPD.length; i++) {
      const currentRow = startRowAPD + i;
      const mergeRange = newSheet.getRange(currentRow, 5, 1, 2); // Kolom E (5) dan F (6)
      mergeRange.merge();

      const mergeCD = newSheet.getRange(currentRow, 3, 1, 2); // Kolom C (3) dan D (4)
      mergeCD.merge();

      // Menggabungkan kolom G dan H
      const mergeRangeGH = newSheet.getRange(currentRow, 7, 1, 2); // Kolom G (7) dan H (8)
      mergeRangeGH.merge();
    }
  }
  newSheet.hideRows(29);

  // Menentukan baris mulai dan kolom Peralatan
  const startRowPeralatan = 24; // Mulai dari baris ke-25

  // Menuliskan data secara batch ke sheet dengan menyisipkan baris baru Peralatan
  if (writePeralatan.length > 0) {
    const writePeralatanWithNumber = writePeralatan.map((row, index) => [index + 1, row[0]]);
    // const writeLabelPeralatan = countsPerString.map((row) => [row[0]]);
    const writeLabelPeralatan = countsPerStringPeralatan.map((row) => [row]);

    // Iterasi writeAPDWithNumber
    const labelCounts = writePeralatanWithNumber.map(([rowNumber, rowKey]) => {
      const simpleKey = `${rowKey}`; // Gunakan rowKey sebagai simpleKey

      // Periksa apakah key ada di dbSimpleMap
      if (dbSimpleMap[simpleKey]) {
        const availableEntry = dbSimpleMap[simpleKey];

        // Filter hanya entry dengan availability 'AVAILABLE'
        const availableEntries = dbSimpleMap[simpleKey].filter(entry => entry.availability === 'AVAILABLE');

        // Hitung total labels dari semua availableEntries
        const totalLabels = availableEntry.reduce((sum, entry) => sum + entry.labels.length, 0);

        // Logger.log(`Jumlah total "${totalLabels}", dengan entry.labels.length: "${entry.labels.length}". Juga dengan entrt.labels: "${entry.labels}"`);

        const namaItemAPD = availableEntry.reduce((sum, entry) => entry.nama, '').replace(/\s/g, '');

        let jenisItem = '';

        if (writePeralatan.length < inputAKB.length) {
          jenisItem = 'AKB';
        } else {
          jenisItem = 'ALT';
        }

        const keterangan = tanggalFormatted + nomorPermintaan + jenisItem + namaItemAPD;

        // Kembalikan rowNumber dan totalLabels
        return [rowNumber, totalLabels, keterangan];
      } else {
        // Jika key tidak ditemukan, kembalikan jumlah label sebagai 0
        return [rowNumber, 0, ''];
      }
    });

    // Menyisipkan baris baru
    newSheet.insertRows(startRowPeralatan, writePeralatan.length);

    const writeRange = newSheet.getRange(startRowPeralatan, 2, writePeralatanWithNumber.length, 2); // (baris, kolom, jumlah baris, jumlah kolom)
    writeRange.setValues(writePeralatanWithNumber);

    // a. Mengatur alignment untuk Kolom B (nomor) ke tengah
    const nomorRange = writeRange.offset(0, 0, writePeralatanWithNumber.length, 1); // Kolom B
    nomorRange.setHorizontalAlignment("center");

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

    // Jumlah Total
    const startRowJumlahStock = startRowPeralatan; // Baris mulai untuk stock
    const numRowsJumlahStock = writePeralatanWithNumber.length; // Jumlah baris untuk stock
    const jumlahStockColumn = 5; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const jumlahStockRange = newSheet.getRange(startRowJumlahStock, jumlahStockColumn, numRowsJumlahStock, 1);
    jumlahStockRange.setValues(labelCounts.map(([_, totalLabels, keterangan]) => [totalLabels]));

    // Jumlah Total
    const startRowKeterangan = startRowPeralatan; // Baris mulai untuk keterangan
    const numRowsKeterangan = writePeralatanWithNumber.length; // Jumlah baris untuk keterangan
    const keteranganColumn = 11; // Kolom E adalah kolom ke-5

    // Menentukan rentang untuk checkmark
    const keteranganRange = newSheet.getRange(startRowKeterangan, keteranganColumn, numRowsKeterangan, 1);
    keteranganRange.setWrap(true);
    keteranganRange.setValues(labelCounts.map(([_, totalLabels, keterangan]) => [keterangan]));

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

    // (Opsional) Mengatur alignment ke tengah untuk checkmark
    checkmarkRange.setHorizontalAlignment("center");

    // Membuat merged cells pada kolom E dan F untuk setiap baris yang baru disisipkan
    for (let i = 0; i < writePeralatan.length; i++) {
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
  // Menuliskan nilai pada spesifik kolom
  newSheet.getRange(9, 4).setValue(judulPembelajaran);
  newSheet.getRange(12, 4).setValue(tanggalMulai);
  newSheet.getRange(13, 4).setValue(tanggalSelesai);
  newSheet.getRange(19, 11).setValue(idTransaksi);

  const inputNumber = formatPhone(nomorTelp);

  let linkPDF = generatePDFPeminjaman(newSheetName, namePdf);
  let pesanLink = "📤 *Inventarisasi Keluar Laboratorium* \n dengan *ID "+ (idTransaction || 'Data tidak tersedia') + "*\n\n" +
        "❔ *Perihal:* " + (judulPembelajaran || 'Data tidak tersedia') + "\n\n" +
        "📜 *Formulir Keluar:* " + (linkPDF || 'Data tidak tersedia') + " \n\n" +
        "‼️ _(pastikan pengembalian Anda sesuai keterangan yang tertera pada Formulir tersebut di atas)_" + "\n\n" +
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

function countLabelsInString(labelString) {
  if (!labelString || typeof labelString !== 'string') {
    return 0;
  }

  // Pisahkan string berdasarkan ";" atau "," dan hitung elemen yang valid
  const labels = labelString.split(/[,;]/).map(label => label.trim()).filter(label => label !== '');

  return labels.length;
}

function processData(data) {
  // Contoh pengolahan data
  const filteredData = data.filter(item => item !== undefined && item !== "");
  filteredData.push("Kolom Tambahan 1", "Kolom Tambahan 2");
  return filteredData;
}

/**
 * Fungsi untuk menghasilkan PDF dari sheet tertentu
 * @param {string} sheetName - Nama sheet yang akan dikonversi menjadi PDF
 * @param {string} pdfName - Nama file PDF yang akan disimpan
 */
function generatePDFPeminjaman(sheetName, pdfName) {
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
  const blob = response.getBlob().setName(pdfName + '_Out.pdf');

  // Menyimpan PDF ke folder tertentu di Drive (opsional)
  const folder = DriveApp.getFolderById('1NJWujNEBwh8YRh72QioRpjQWYbIBROE4'); // Ganti dengan ID folder tujuan
  const file = folder.createFile(blob);

  // Menyimpan PDF di root Drive
  // DriveApp.createFile(blob);
  const fileUrl = file.getUrl();

  Logger.log(`PDF "${pdfName}.pdf" berhasil dibuat dan disimpan di Google Drive dengan link "${fileUrl}"`);

  const message = "Ini adalah pesan otomatis, file peminjaman anda telah selesai dibuat pada link berikut: ";

  return `${fileUrl}`;
}

function messageToWhatsapp(pesan) {

  Logger.log("Pesan yang di dapatkan: " + pesan);

  var TokenFonnte = "6ccP7EgJ544fM3vtG4QZ" //isi dengan token fonnte

  var url = "https://api.fonnte.com/send";

  //kirim pesan ke admin
  var options_admin = {
    "method": "post",
    "headers": {
      "Authorization": TokenFonnte
    },
    "payload": {
      "target": "120363318268953826@g.us",
      "message": pesan
    }
  };
  var response = UrlFetchApp.fetch(url, options_admin);

  Logger.log(response.getContentText());
}

function groupMainInput(input) {
  const groupMap = {};

  if (!input || typeof input !== 'string') {
    return groupMap;
  }

  // Memisahkan input berdasarkan koma dan menghapus spasi ekstra
  const items = input.split(',').map(item => item.trim()).filter(item => item !== "");

  items.forEach(item => {
    if (item.includes('#')) {
      const parts = item.split('#');
      const mainInput = parts[0].trim();
      const label = parts[1].trim();

      if (mainInput) {
        if (!groupMap[mainInput]) {
          groupMap[mainInput] = [];
        }
        groupMap[mainInput].push(prefixLabel(label));
      }
    } else {
      const mainInput = item.trim();
      if (mainInput) {
        if (!groupMap[mainInput]) {
          groupMap[mainInput] = [];
        }
        groupMap[mainInput].push('');
      }
    }
  });

  return groupMap;
}

function prefixLabel(label) {
  // Implementasikan logika prefixLabel sesuai kebutuhan
  return label; // Placeholder
}
