const GENERATE_PDF_MRK04 = 'MRK-04';
const SOURCE_SHEET_NAME = 'HAR_INFRA';
const DROPDOWN_CELL_MRK04 = 'E5';

function onEditCopyMrk04(e) {
  const ss = e.source;
  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();

  // Cek apakah sel yang diedit adalah E5 di sheet 'MRK-04'
  if (editedSheet.getName() !== GENERATE_PDF_MRK04 || editedRange.getA1Notation() !== DROPDOWN_CELL_MRK04) {
    return; // Keluar jika bukan sel yang relevan
  }

  const selectedValue = editedRange.getValue();
  if (!selectedValue || !selectedValue.includes('Periode')) {
    return; // Keluar jika nilai dropdown kosong atau format tidak valid
  }

  // Tampilkan notifikasi "Processing..."
  ss.toast('Processing...', 'Status', 3);

  const mrk04Sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const generatePDFMrk04 = ss.getSheetByName(GENERATE_PDF_MRK04);

  const lastRowMrk04 = mrk04Sheet.getLastRow();
  if (lastRowMrk04 < 3) {
    ss.toast('No data found in HAR_INFRA sheet.', 'Info', 3);
    return;
  }

  const numRows = lastRowMrk04 - 2;
  const mrk04Data = mrk04Sheet.getRange(2, 1, numRows, 16).getValues();

  // Array untuk menyimpan hasil yang unik
  const results = [];

  //Logger.log(mrk04Data);

  for (let i = 0; i < mrk04Data.length; i++) {
    const row = mrk04Data[i];
    const periode = row[14]; // Mengambil periode dari kolom ke-14

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
    generatePDFMrk04.showRows(10, generatePDFMrk04.getMaxRows() - 9 );

    // Bersihkan Baris
    const maxRows = generatePDFMrk04.getMaxRows();
    generatePDFMrk04.getRange(10, 1, maxRows - 9, generatePDFMrk04.getLastColumn()).clear();
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
    const tidakLanjut = row[9];

    const columnC = [
      `Ruang: ${ruang || ''}`,
      `CCTV: ${device || ''} - ${ipaddress || ''}`
    ].filter(Boolean).join('\n');
    
    return [
      (index + 1).toString(), // Mengubah nomor urut menjadi string
      formattedTanggal,       // Menggunakan tanggal yang sudah diformat
      columnC,
      statusCCTVAda,
      statusCCTVTidak,
      kondisiCCTVBaik,
      kondisiCCTVTRusak,
      tidakLanjut
    ];
});
  // Tambahkan border ke data yang baru dimasukkan
    const dataRange = generatePDFMrk04.getRange(10, 1, dataToInsert.length, 8);
    dataRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);


// Jika Anda ingin memasukkan dataToInsert ke dalam sheet, Anda bisa menambahkannya di sini
  if (dataToInsert.length > 0) {
    const rangeToInsert = generatePDFMrk04.getRange(10, 1, dataToInsert.length, dataToInsert[0].length);
    rangeToInsert.setValues(dataToInsert);

    // Mengatur format untuk kolom A (index) dalam satu langkah
    const columnA = generatePDFMrk04.getRange(10, 1, dataToInsert.length);
    columnA.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnB = generatePDFMrk04.getRange(10, 2, dataToInsert.length);
    columnB.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnC = generatePDFMrk04.getRange(10, 3, dataToInsert.length);
    columnC.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
    const columnD = generatePDFMrk04.getRange(10, 4, dataToInsert.length);
    columnD.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnE = generatePDFMrk04.getRange(10, 5, dataToInsert.length);
    columnE.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnF = generatePDFMrk04.getRange(10, 6, dataToInsert.length);
    columnF.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnG = generatePDFMrk04.getRange(10, 7, dataToInsert.length);
    columnG.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('center');
    const columnHIJMerge = generatePDFMrk04.getRange(10, 8, dataToInsert.length, 3);
    columnHIJMerge.mergeAcross();
    columnHIJMerge.setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
  }
  const maxRows = generatePDFMrk04.getMaxRows(); 
    // Sembunyikan baris yang tidak terpakai
    if (dataToInsert.length < maxRows - 9) {
      generatePDFMrk04.hideRows(
        10 + dataToInsert.length,
        maxRows - dataToInsert.length - 9
      );
    }
  ss.toast('Completed!', 'Status', 3);
}


// Fungsi yang dijalankan saat file spreadsheet dibuka
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // Mendapatkan antarmuka pengguna

  // Menambahkan menu baru dengan nama "Custom Menu"
  ui.createMenu('Generate PDF')
    .addItem('Generate PDF MRK04', 'GenPDFMRK04') // Menambahkan item menu
    .addItem('Generate PDF HAR', 'GenPDFHAR') // Menambahkan item menu
    .addToUi(); // Menambahkan menu ke antarmuka pengguna
}