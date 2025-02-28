function pengembalianAlat(e) {
  // Pastikan event tidak null dan ada range yang diedit  
  if (!e || !e.range) return Logger.log('e tidak ada');

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB');

  // Hanya jalankan untuk sheet "Form Responses 1"  
  // if (sheetName !== 'LOGIN02') return Logger.log('Sheet tidak ada');

  const rowToProcess = e.range.getRow();

  const data = sheet.getDataRange().getValues();
  const originalHeader = data[0];

  // Get index header by name
  const merkTipeName = 'nama merk_tipe';
  const barangDikembalikanName = 'Barang/Material yang akan dikembalikan?';
  const labelBarangKeluarName = 'No. Label Keluar';
  const labelBarangMasukName = 'No. Label Masuk';
  const idTransaksiName = 'idTransaksi';
  const tujuanPengembalianName = 'Tujuan Pengembalian';
  const timestampName = 'Timestamp';
  const namaPenerimaName = 'Nama Identitas Penerima';
  const lokasiPenyimpananName = 'Lokasi Penyimpanan';
  const jumlahName = 'Jumlah';
  const nomorTelpName = 'Nomor Whatsapp';

  const input1Index = getColumnIndex(originalHeader, merkTipeName);
  const barangDikembalikanIndex = getColumnIndex(originalHeader, barangDikembalikanName);
  const labelBarangKeluarIndex = getColumnIndex(originalHeader, labelBarangKeluarName);
  const labelBarangMasukIndex = getColumnIndex(originalHeader, labelBarangMasukName);
  const idTransaksiIndex = getColumnIndex(originalHeader, idTransaksiName);
  const tujuanPengembalianIndex = getColumnIndex(originalHeader, tujuanPengembalianName);
  const timestampIndex = getColumnIndex(originalHeader, timestampName);
  const namaPenerimaIndex = getColumnIndex(originalHeader, namaPenerimaName);
  const lokasiPenyimpananIndex = getColumnIndex(originalHeader, lokasiPenyimpananName);
  const jumlahIndex = getColumnIndex(originalHeader, jumlahName);
  const nomorTelpIndex = getColumnIndex(originalHeader, nomorTelpName);

  // Ambil nilai Edited Row
  const valueBarangDikembalikan = sheet.getRange(rowToProcess, barangDikembalikanIndex + 1).getValues();
  const valueTujuanPengembalian = sheet.getRange(rowToProcess, tujuanPengembalianIndex + 1).getValue();
  const valueNamaPenerima = sheet.getRange(rowToProcess, namaPenerimaIndex + 1).getValue();
  const valueLokasiPenyimpanan = sheet.getRange(rowToProcess, lokasiPenyimpananIndex + 1).getValue();
  const valueNomorTelp = sheet.getRange(rowToProcess, nomorTelpIndex + 1).getValue();

  const valueBarang = sheet.getRange(rowToProcess, barangDikembalikanIndex + 1).getValue();

  Logger.log("Bentuk Input: " + valueBarang);

  // Split and group the input
  const groupMap = groupMainInput2(valueBarang);

  // Logging hasil untuk setiap entri
  groupMap.forEach(entry => {
    Logger.log("Bentuk mainInput: " + entry.mainInput);
    Logger.log("Bentuk label: " + entry.labels.join('; '));
    Logger.log("Bentuk idTransaksi: " + entry.idTransaksi);
  });

  const timestamp = sheet.getRange(rowToProcess, timestampIndex + 1).getValue();

  // Ambil data dari kolom target
  // const barang = sheet.getRange(7, barangDikembalikanIndex + 1).getValues();

  const outputRows = [];

  // Hasil pemisahan
  const namaBarang = [];
  const labelBarang = [];
  const idTransaksi = [];

  groupMap.forEach(entry => {
    const outputRow = Array(originalHeader.length).fill('');
    outputRow[barangDikembalikanIndex] = entry.mainInput;
    outputRow[labelBarangMasukIndex] = entry.labels.join('; ');
    outputRow[jumlahIndex] = entry.labels.length;
    outputRow[idTransaksiIndex] = entry.idTransaksi; // Sama untuk semua label di grup ini
    outputRow[tujuanPengembalianIndex] = valueTujuanPengembalian;
    outputRow[timestampIndex] = timestamp;
    outputRow[namaPenerimaIndex] = valueNamaPenerima;
    outputRow[lokasiPenyimpananIndex] = valueLokasiPenyimpanan;
    outputRow[nomorTelpIndex] = valueNomorTelp;
    outputRows.push(outputRow);
  });

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

  // Validasi apakah kolom ditemukan
  if (input1Index === -1) {
    throw new Error(`Kolom dengan nama '${merkTipeName}' tidak ditemukan.`);
  }

  populateGoogleForms_gformlogout();

  // Jika ada log atau tindakan lain
  Logger.log("populateGoogleForms_gformlogout berhasil dijalankan.");

  populateGoogleForms_gformlogin();
  Logger.log("populateGoogleForms_gformlogin berhasil dijalankan.");
}

// Fungsi untuk mendapatkan indeks kolom berdasarkan nama header  
function getColumnIndex(headerRow, columnName) {
  return headerRow.indexOf(columnName);
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