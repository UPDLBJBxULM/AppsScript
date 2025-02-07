const GENERATE_PDF_HAR = 'HAR';
const SOURCE_SHEET_NAME_HAR = 'HAR_INFRA';
const DROPDOWN_CELL_HAR = 'E6:F6';

function onEditCopyHar(e) {
  try {
    const ss = e.source;
    const editedRange = e.range;
    const editedSheet = editedRange.getSheet();
    var generateSheet = ss.getSheetByName('HAR'); 
    // Mendapatkan data periode dari E6
    var periodeData = generateSheet.getRange('E6').getValue();
    
    // Log data untuk melihat nilai sebelum perubahan
    Logger.log('Data E6: ' + periodeData);

    Logger.log('tes1')

    // Cek apakah sel yang diedit adalah E6 di sheet 'HAR'
    // if (editedSheet.getName() !== GENERATE_PDF_HAR || editedRange.getA1Notation() !== DROPDOWN_CELL_HAR) {
    //   return; // Keluar jika bukan sel yang relevan
    // }
    Logger.log('Data E6: ' + periodeData);

    Logger.log('tes2')

    const selectedValue = editedRange.getValue();
    if (!selectedValue) {
      return; // Keluar jika nilai dropdown kosong atau format tidak valid
    }
    Logger.log('Data E6: ' + periodeData);

    Logger.log('tes3')
    // Tampilkan notifikasi "Processing..."
    ss.toast('Processing...', 'Status', 3);
  const harSheet = ss.getSheetByName(SOURCE_SHEET_NAME_HAR);
  const generatePDFHar = ss.getSheetByName(GENERATE_PDF_HAR);

  const lastRowHar = harSheet.getLastRow();
  if (lastRowHar < 3) {
    ss.toast('No data found in HAR_INFRA sheet.', 'Info', 3);
    return;
  }

  const numRows = lastRowHar - 2;
  const harData = harSheet.getRange(2, 1, numRows, 16).getValues();

  // Array untuk menyimpan hasil yang unik
  const results = [];

  //Logger.log(harData);

  for (let i = 0; i < harData.length; i++) {
    const row = harData[i];
    const periode = row[15]; // Mengambil periode dari kolom ke-14

    // Lewati jika periode tidak ada
    if (!periode) {
        continue;
    }

    // Memeriksa apakah periode cocok dengan nilai yang dipilih
    if (periode.trim() === selectedValue.trim()) {
        // Menambahkan seluruh baris ke dalam results
        results.push(row); // Menyimpan seluruh baris
    }
  }
  // Menampilkan hasil
  //Logger.log(results);

  if(results.length > 0){
    // Tampilkan baris
    generatePDFHar.showRows(17, generatePDFHar.getMaxRows() - 16 );

    // Bersihkan Baris
    const maxRows = generatePDFHar.getMaxRows();
    generatePDFHar.getRange(17, 1, maxRows - 16, generatePDFHar.getLastColumn()).clear();
  }

  const dataToInsert = results.map((row, index) => {
    const ruang = row[3];
    const tanggal = row[0]; // Tanggal dalam format Date

    // Log nilai tanggal untuk pemeriksaan
    // Logger.log(`Tanggal sebelum format: ${tanggal}`);

    // Pastikan tanggal adalah objek Date
    let formattedTanggal = '';
    if (tanggal instanceof Date) {
        const day = String(tanggal.getDate()).padStart(2, '0'); // Ambil hari dan tambahkan nol di depan jika perlu
        const month = String(tanggal.getMonth() + 1).padStart(2, '0'); // Ambil bulan (0-11) dan tambahkan nol di depan
        const year = tanggal.getFullYear(); // Ambil tahun
        formattedTanggal = `${day}.${month}.${year}`; // Format menjadi DD.MM.YYYY
    } else {
        // Jika tanggal tidak valid, bisa diisi dengan string kosong atau nilai default
        formattedTanggal = 'Invalid Date'; // Atau bisa diisi dengan ''
    }

    const device = row[1];
    const ipaddress = row[2]; 
    const statusCCTVAda = row[4];
    const statusCCTVTidak = row[5];
    const kondisiCCTVBaik = row[6];
    const kondisiCCTVTRusak = row[7];
    const kondisiFungsi = row[8];
    const tidakLanjut = row[9];
    const eviden = row[13];

    const columnCD = [
      `Ruang: ${ruang || ''}`,
      `CCTV: ${device || ''} - ${ipaddress || ''}\n`,
      `eviden: ${eviden}`
    ].filter(Boolean).join('\n');
    
    return [
      (index + 1).toString(), // Mengubah nomor urut menjadi string
      formattedTanggal,       // Menggunakan tanggal yang sudah diformat
      columnCD,
      '',
      statusCCTVAda,
      statusCCTVTidak,
      kondisiCCTVBaik,
      kondisiCCTVTRusak,
      kondisiFungsi,
      tidakLanjut
    ];
});
  // Tambahkan border ke data yang baru dimasukkan
    const dataRange = generatePDFHar.getRange(17, 1, dataToInsert.length, 10);
    dataRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);


// Jika Anda ingin memasukkan dataToInsert ke dalam sheet, Anda bisa menambahkannya di sini
  if (dataToInsert.length > 0) {
    const rangeToInsert = generatePDFHar.getRange(17, 1, dataToInsert.length, dataToInsert[0].length);
    rangeToInsert.setValues(dataToInsert);

    // Mengatur format untuk kolom A (index) dalam satu langkah
    const columnA = generatePDFHar.getRange(17, 1, dataToInsert.length);
    columnA.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnB = generatePDFHar.getRange(17, 2, dataToInsert.length);
    columnB.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnCDMerge = generatePDFHar.getRange(17, 3, dataToInsert.length, 2);
    columnCDMerge.mergeAcross();
    columnCDMerge.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
    
    const columnE = generatePDFHar.getRange(17, 5, dataToInsert.length);
    columnE.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnF = generatePDFHar.getRange(17, 6, dataToInsert.length);
    columnF.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnG = generatePDFHar.getRange(17, 7, dataToInsert.length);
    columnG.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnH = generatePDFHar.getRange(17, 8, dataToInsert.length);
    columnH.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnI = generatePDFHar.getRange(17, 9, dataToInsert.length);
    columnI.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnJKMerge = generatePDFHar.getRange(17, 10, dataToInsert.length, 2);
    columnJKMerge.mergeAcross();
    columnJKMerge.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
  }
  const maxRows = generatePDFHar.getMaxRows(); 
    // Sembunyikan baris yang tidak terpakai
    if (dataToInsert.length < maxRows - 16) {
      generatePDFHar.hideRows(
        17 + dataToInsert.length,
        maxRows - dataToInsert.length - 16
      );
    }
  ss.toast('Completed!', 'Status', 3);
  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}

// Fungsi yang dijalankan saat file spreadsheet dibuka
// function onOpen() {
//   const ui = SpreadsheetApp.getUi(); // Mendapatkan antarmuka pengguna

//   // Menambahkan menu baru dengan nama "Custom Menu"
//   ui.createMenu('Generate PDF')
    
//     .addToUi(); // Menambahkan menu ke antarmuka pengguna
// }