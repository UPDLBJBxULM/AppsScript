function peminjamanAlat(e) {
  // Pastikan event tidak null dan ada range yang diedit  
  if (!e || !e.range) return Logger.log('e tidak ada');

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB');

  // Hanya jalankan untuk sheet "Form Responses 1"  
  // if (sheetName !== 'LOGOUT02') return Logger.log('Sheet tidak ada');

  // Dapatkan baris yang baru saja diedit  
  const editedRow = e.range.getRow();

  // Pastikan baris yang diedit bukan header (baris 1 atau 2)  
  if (editedRow <= 1) return;

  // Jalankan fungsi splitting untuk baris terakhir  
  processAndSplitRow(sheet, dbSheet, editedRow);
}

function processAndSplitRow(sheet, dbSheet, rowToProcess) {
  const data = sheet.getDataRange().getValues();
  const dbValues = dbSheet.getDataRange().getValues();
  const originalHeader = data[0];

  // Fungsi untuk mendapatkan indeks kolom berdasarkan nama header  
  function getColumnIndex(headerRow, columnName) {
    return headerRow.indexOf(columnName);
  }

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

  for (let i = 2; i < dbValues.length; i++) {
    const row = dbValues[i];
    const nama = row[dbNamaIdx].toString().trim();
    const merkTipe = row[dbMerkTipeIdx].toString().trim();
    const label = row[dbLabelIdx].toString().trim();
    const availability = row[dbAvailabilityIdx].toString().trim().toUpperCase();
    const idLog = row[dbIdLogIdx].toString().trim();
    const jenis = row[dbJenisIdx].toString().trim();

    const key = `${nama} ${merkTipe}`;
    if (!dbMap[key]) {
      dbMap[key] = {
        availability: availability,
        idLog: idLog,
        jenis: jenis,
        label: label,
        labels: [] // Buat array untuk menyimpan semua label
      };
    }

    if (availability === 'AVAILABLE') {
      // Tambahkan label ke array
      dbMap[key].labels.push(label)
    }

    if (dbMap[key].availability === 'AVAILABLE') {
      // Logger.log(`Bentuk dbmap: ${JSON.stringify(dbMap[key])}`);
    }

  }

  const input1Name = 'Alat Kerja Bantu Tersedia';
  const input2Name = 'Alat Uji dan Ukur Tersedia';
  const input3Name = 'APD Tersedia';
  const input4Name = 'Material Tersedia';
  const timestamp1Name = 'Timestamp';
  const barangName = 'nama';
  const pembelajaranName = 'Judul Pembelajaran';
  const jumlahAKBName = 'Jumlah Label AKB';
  const jumlahALTName = 'Jumlah Label ALT';
  const totalBarangName = 'Jumlah Keluar';
  const rangkumLabelName = 'No. Label Keluar';
  const peminjamanName = 'Tujuan Peminjam';
  const nomorPermintaanName = 'Nomor Permintaan';
  const tanggalMulaiName = 'Tanggal Mulai';
  const tanggalSelesaiName = 'Tanggal Selesai';
  const idLogName = 'idLOG';
  const jenisName = 'jenis';
  const jumlahTersediaName = 'Jumlah Tersedia';
  const labelTerdatabaseName = 'No. Label Terdatabase';
  const keterangan1Name = 'Nama Identitas Peminjam';
  const keterangan2Name = 'Lokasi Penggunaan';
  const nomorTelpName = 'Nomor Whatsapp';

  // Indeks kolom input  
  const input1Index = getColumnIndex(originalHeader, input1Name);
  const input2Index = getColumnIndex(originalHeader, input2Name);
  const input3Index = getColumnIndex(originalHeader, input3Name);
  const input4Index = getColumnIndex(originalHeader, input4Name);
  const jumlahAKBIndex = getColumnIndex(originalHeader, jumlahAKBName);
  const jumlahALTIndex = getColumnIndex(originalHeader, jumlahALTName);
  const totalBarangIndex = getColumnIndex(originalHeader, totalBarangName);
  const rangkumanLabelIndex = getColumnIndex(originalHeader, rangkumLabelName);
  const peminjamanIndex = getColumnIndex(originalHeader, peminjamanName);
  const nomorPermintaanIndex = getColumnIndex(originalHeader, nomorPermintaanName);
  const tanggalMulaiIndex = getColumnIndex(originalHeader, tanggalMulaiName);
  const tanggalSelesaiIndex = getColumnIndex(originalHeader, tanggalSelesaiName);
  const idLogIndex = getColumnIndex(originalHeader, idLogName);
  const jenisIndex = getColumnIndex(originalHeader, jenisName);
  const jumlahTersediaIndex = getColumnIndex(originalHeader, jumlahTersediaName);
  const labelTerdatabaseIndex = getColumnIndex(originalHeader, labelTerdatabaseName);
  const keterangan1Index = getColumnIndex(originalHeader, keterangan1Name);
  const keterangan2Index = getColumnIndex(originalHeader, keterangan2Name);
  const nomorTelpIndex = getColumnIndex(originalHeader, nomorTelpName);

  // Indeks kolom that is not processed  
  const timestampIndex = getColumnIndex(originalHeader, timestamp1Name);
  const barangIndex = getColumnIndex(originalHeader, barangName);
  const pembelajaranIndex = getColumnIndex(originalHeader, pembelajaranName);

  // Ambil nilai Pembelajaran  
  const valuePembelajaran = sheet.getRange(rowToProcess, pembelajaranIndex + 1).getValue();
  const valuePeminjaman = sheet.getRange(rowToProcess, peminjamanIndex + 1).getValue();
  const valueNomorPermintaan = sheet.getRange(rowToProcess, nomorPermintaanIndex + 1).getValue();
  const valueTanggalMulai = sheet.getRange(rowToProcess, tanggalMulaiIndex + 1).getValue();
  const valueTanggalSelesai = sheet.getRange(rowToProcess, tanggalSelesaiIndex + 1).getValue();
  const valueKeterangan1 = sheet.getRange(rowToProcess, keterangan1Index + 1).getValue();
  const valueKeterangan2 = sheet.getRange(rowToProcess, keterangan2Index + 1).getValue();
  const valueNomorTelp = sheet.getRange(rowToProcess, nomorTelpIndex + 1).getValue();

  // Ambil baris yang akan diproses  
  const row = data[rowToProcess - 1]; // -1 karena array dimulai dari 0  

  const input1 = row[input1Index];
  
  const input2 = row[input2Index];
  const input3 = row[input3Index];
  const input4 = row[input4Index];
  // const jumlahAKB = row[jumlahAKBIndex];
  // const jumlahALT = row[jumlahALTIndex];
  // Logger.log(`Bentuk input1: "${input1}".`);
  // Grup mainInput dan label untuk Input1 hingga Input4  
  const groupMap1 = groupMainInput(input1);
  const groupMap2 = groupMainInput(input2);
  const groupMap3 = groupMainInput(input3);
  const groupMap4 = groupMainInput(input4);

  const mainInputs1 = Object.keys(groupMap1 || {});
  const labels1 = mainInputs1.map(key => groupMap1[key].join('; '));

  Logger.log(`Bentuk mainput1: "${mainInputs1}".`);
  // Logger.log(`Bentuk label1: "${labels1}".`);

  const mainInputs2 = Object.keys(groupMap2);
  const labels2 = mainInputs2.map(key => groupMap2[key].join('; '));

  const mainInputs3 = Object.keys(groupMap3);
  const labels3 = mainInputs3.map(key => groupMap3[key].join('; '));

  const mainInputs4 = Object.keys(groupMap4);
  const labels4 = mainInputs4.map(key => groupMap4[key].join('; '));

  // Gabungkan semua item menjadi array untuk kolom `barang`  
  const allItems = [...mainInputs1, ...mainInputs2, ...mainInputs3, ...mainInputs4]
    .map(item => item.trim())
    .filter(Boolean);

  const allLabels = [...labels1, ...labels2, ...labels3, ...labels4]
    .map(item => item.trim())
    .filter(Boolean);

  const maxGroups = Math.max(mainInputs1.length, mainInputs2.length, mainInputs3.length, mainInputs4.length, allItems.length);
  const timestamp = row[timestampIndex] || new Date();

  // Hapus baris lama dan siapkan untuk menulis ulang  
  const outputRows = [];

  // Buat Map DB berdasarkan nama + label untuk memudahkan pencocokan
  const dbSimpleMap = {};
  for (const key in dbMap) {
    // key format: "nama merkTipe label"
    const parts = key.split(' ');
    const merkTipe = parts.pop(); // Ambil merk tipe
    const nama = parts.join(' '); // Sisa adalah nama
    let simpleKey = null;
    if (merkTipe) {
      simpleKey = `${nama} ${merkTipe}`;
      // Logger.log(`Simple w merktipe Key: "${simpleKey}".`);
    } else {
      simpleKey = `${nama}`;
      // Logger.log(`Simple Key: "${simpleKey}".`);
    }

    // Gabungkan semua label dengan koma
    const labels = dbMap[key].labels.join(', ');
    // Logger.log(`Labels for Key "${key}": "${labels}"`);

    if (!dbSimpleMap[simpleKey]) {
      dbSimpleMap[simpleKey] = [];
    }
    dbSimpleMap[simpleKey].push(dbMap[key]);
  }

  // Persiapkan untuk menulis hasil count
  // const jmlAvailable = [];
  // const idLogList = [];
  // const labelList = [];
  // const jenisList = [];

  for (let j = 0; j < maxGroups; j++) {
    // Logger/log("Ini tereksekusi");
    // Inisialisasi baris output dengan panjang originalHeader  
    const outputRow = Array(originalHeader.length).fill('');

    // Proses Input1  
    if (j < mainInputs1.length) {
      if (j === 0) {
        outputRow[input1Index] = input1;
      }
      outputRow[input1Index + 1] = mainInputs1[j];
      Logger.log(`mainInput1 untuk index "${j}": "${mainInputs1[j]}"`);
      outputRow[input1Index + 2] = prefixLabel(labels1[j]);
      outputRow[jumlahAKBIndex] = countLabelsInString(labels1[j]);
    }
    Logger.log(`outputRow input1 untuk index "${j}": "${outputRow[input1Index + 1]}"`);
    // Proses Input2  
    if (j < mainInputs2.length) {
      if (j === 0) {
        outputRow[input2Index] = input2;
      }
      outputRow[input2Index + 1] = mainInputs2[j];
      outputRow[input2Index + 2] = prefixLabel(labels2[j]);
      outputRow[jumlahALTIndex] = countLabelsInString(labels2[j]);
    }

    // Proses Input3  
    if (j < mainInputs3.length) {
      if (j === 0) {
        outputRow[input3Index] = input3;
      }
      outputRow[input3Index + 1] = mainInputs3[j];
      outputRow[input3Index + 2] = prefixLabel(labels3[j]);
    }

    // Proses Input4  
    if (j < mainInputs4.length) {
      if (j === 0) {
        outputRow[input4Index] = input4;
      }
      outputRow[input4Index + 1] = mainInputs4[j];
      outputRow[input4Index + 2] = prefixLabel(labels4[j]);
    }

    let count = 0;
    let idLogGabungan = '';
    let labelGabungan = [];
    let jenisGabungan = '';
    let jmlLabelGabungan = null;
    const simpleKey = `${allItems[j]}`;

    if (dbSimpleMap[simpleKey]) {
      dbSimpleMap[simpleKey].forEach(entry => {
        // if (entry.availability === 'AVAILABLE') {
        //   count += 1;
        //   Logger.log(`entry label: ${JSON.stringify(entry)}`)
        //   idLogGabungan = entry.idLog; // Gunakan idLog terbaru (atau simpan semua jika perlu)
        //   jenisGabungan = entry.jenis;
        //   jmlLabelGabungan = entry.labels.length;
        //   labelGabungan.push(...entry.labels); // Tambahkan label ke array
        // }
        if (true) {
          count += 1;
          Logger.log(`entry label: ${JSON.stringify(entry)}`)
          idLogGabungan = entry.idLog; // Gunakan idLog terbaru (atau simpan semua jika perlu)
          jenisGabungan = entry.jenis;
          jmlLabelGabungan = entry.labels.length;
          // labelGabungan.push(...entry.labels);
          labelGabungan.push(...entry.labels); // Tambahkan label ke array
        }
      });
    }
    outputRow[timestampIndex] = timestamp;
    outputRow[barangIndex] = allItems[j];
    outputRow[rangkumanLabelIndex] = allLabels[j] + ';';
    outputRow[pembelajaranIndex] = valuePembelajaran;
    outputRow[peminjamanIndex] = valuePeminjaman;
    outputRow[nomorPermintaanIndex] = valueNomorPermintaan;
    outputRow[tanggalMulaiIndex] = valueTanggalMulai;
    outputRow[tanggalSelesaiIndex] = valueTanggalSelesai;
    outputRow[idLogIndex] = idLogGabungan;
    outputRow[jenisIndex] = jenisGabungan;
    const labelTersedia = labelGabungan.join('; ') + ';';
    Logger.log(`Bentuk label tersedia: ${labelTersedia}`);
    outputRow[labelTerdatabaseIndex] = labelTersedia;
    outputRow[jumlahTersediaIndex] = jmlLabelGabungan;
    outputRow[keterangan1Index] = valueKeterangan1;
    outputRow[keterangan2Index] = valueKeterangan2;
    outputRow[totalBarangIndex] = countLabelsInString(allLabels[j]);
    outputRow[nomorTelpIndex] = valueNomorTelp;
    outputRows.push(outputRow);
  }

  Logger.log(`Panjang output: ${outputRows.length}`);
  const lastRow = sheet.getLastRow(); // Ambil baris terakhir yang terisi 

  // Tentukan baris berikutnya yang kosong
  const nextRow = lastRow + 1;

  // Menyisipkan data baru pada baris kosong berikutnya
  sheet.getRange(nextRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

  // Menyisipkan format kolom yang diperlukan
  sheet.getRange(nextRow, 9, outputRows.length, 1).setNumberFormat('@STRING@');  // Label1
  sheet.getRange(nextRow, 12, outputRows.length, 1).setNumberFormat('@STRING@'); // Label2
  sheet.getRange(nextRow, 15, outputRows.length, 1).setNumberFormat('@STRING@'); // Label3
  sheet.getRange(nextRow, 18, outputRows.length, 1).setNumberFormat('@STRING@'); // Label4
  sheet.deleteRows(rowToProcess, 1);

  // sheet.getRange(lastRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

  // // Hapus baris lama  
  // sheet.deleteRows(rowToProcess, 1);
  // // Sisipkan baris baru di posisi yang sama  
  // sheet.getRange(rowToProcess, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

  // // Atur format kolom label sebagai teks  
  // sheet.getRange(rowToProcess, 9, outputRows.length, 1).setNumberFormat('@STRING@'); // Label1  
  // sheet.getRange(rowToProcess, 12, outputRows.length, 1).setNumberFormat('@STRING@'); // Label2  
  // sheet.getRange(rowToProcess, 15, outputRows.length, 1).setNumberFormat('@STRING@'); // Label3  
  // sheet.getRange(rowToProcess, 18, outputRows.length, 1).setNumberFormat('@STRING@'); // Label4

  populateGoogleForms_gformlogout();

  // Jika ada log atau tindakan lain
  Logger.log("populateGoogleForms_gformlogout berhasil dijalankan.");

  populateGoogleForms_gformlogin();
  Logger.log("populateGoogleForms_gformlogin berhasil dijalankan.");
}

function countLabelsInString(labelString) {
  if (!labelString || typeof labelString !== 'string') {
    return 0;
  }

  // Pisahkan string berdasarkan ";" atau "," dan hitung elemen yang valid
  const labels = labelString.split(/[,;]/).map(label => label.trim()).filter(label => label !== '');

  return labels.length;
}

/**
 * Fungsi untuk mengelompokkan mainInput dan mengumpulkan label.
 * @param {string} input - String input yang dipisahkan koma.
 * @returns {Object} - Peta mainInput ke array label.
 */
function groupMainInput(input) {
  const groupMap = {};

  if (!input || typeof input !== 'string') {
    return groupMap;
  }

  const items = input.split(',')
    .map(item => item.trim())
    .filter(item => item !== "");

  items.forEach(item => {
    if (item.includes('#')) {
      const parts = item.split('#');
      const mainInput = parts[0].trim();
      const label = parts[1] ? parts[1].trim() : '';

      if (mainInput) {
        groupMap[mainInput] = groupMap[mainInput] || [];
        groupMap[mainInput].push(prefixLabel(label));
      }
    } else {
      const mainInput = item.trim();
      if (mainInput) {
        groupMap[mainInput] = groupMap[mainInput] || [];
        groupMap[mainInput].push('');
      }
    }
  });

  return groupMap;
}

/**
 * Fungsi untuk menghitung jumlah kemunculan setiap label dalam sebuah array.
 * @param {Array<string>} labels - Array label yang akan dihitung frekuensinya.
 * @returns {Object} - Objek dengan label sebagai kunci dan jumlah kemunculan sebagai nilai.
 */
function countLabels(labels) {
  return labels.reduce((acc, label) => {
    if (!label) return acc; // Lewati label kosong
    if (!acc[label]) {
      acc[label] = 0;
    }
    acc[label]++;
    return acc;
  }, {});
}

/**
 * Fungsi untuk memastikan label diperlakukan sebagai teks.
 * Menambahkan tanda petik tunggal di depan label.
 * @param {string} label - Label yang ingin diproses.
 * @returns {string} - Label yang sudah diproses.
 */
function prefixLabel(label) {
  if (label === '') return '';
  return `${label}`;
}