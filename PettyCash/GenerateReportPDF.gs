/**
 * @fileOverview Menghasilkan laporan PDF dari Google Sheets berdasarkan pilihan dropdown,
 *             menggunakan template HTML yang di-cache untuk efisiensi.
 */

/**
 * Konstanta konfigurasi untuk caching template HTML.
 *
 * @constant {string} PDF_FOLDER_ID - ID folder Google Drive untuk menyimpan PDF yang dihasilkan.
 * @constant {string} TEMPLATE_CACHE_PREFIX - Prefiks untuk kunci cache yang digunakan untuk template HTML.
 * @constant {number} TEMPLATE_CACHE_EXPIRATION_SECONDS - Waktu kadaluarsa cache dalam detik untuk template.
 */
const PDF_FOLDER_ID = '1VGU3E8Dv0o0vXs2JXEul-ge-WabHjlah';
const TEMPLATE_CACHE_PREFIX = 'TEMPLATE_';
const TEMPLATE_CACHE_EXPIRATION_SECONDS = 30;

/**
 * Konfigurasi statis untuk tampilan dan opsi ekspor PDF.
 *
 * @constant {number[]} PDF_REPORT_COLUMN_WIDTHS - Array lebar kolom untuk laporan PDF dalam poin.
 * @constant {object} PDF_EXPORT_OPTIONS - Objek opsi untuk pengaturan ekspor PDF.
 */
const PDF_REPORT_COLUMN_WIDTHS = [30, 80, 80, 115, 115, 95, 70, 45, 45, 70, 95, 90, 30, 195, 75, 30];
const PDF_EXPORT_OPTIONS = {
  exportFormat: 'pdf',
  format: 'pdf',
  size: 'A4',
  portrait: true,
  fitw: true,
  scale: 2,
  top_margin: 0.5,
  bottom_margin: 0.5,
  left_margin: 0.5,
  right_margin: 0.5,
  sheetnames: false,
  printtitle: false,
  pagenumbers: false,
  gridlines: false,
  fzr: false
};

/**
 * Instance cache script untuk menyimpan template.
 * @type {Cache}
 */
const templateCache = CacheService.getScriptCache();

/**
 * Fungsi utama untuk menghasilkan laporan PDF berdasarkan pilihan dropdown di sheet 'GENERATEPDF'.
 * Fungsi ini membaca data, menggunakan template dari 'TemplatePDF', mengisinya dengan data berdasarkan
 * judul laporan yang dipilih, dan menyimpan PDF yang dihasilkan ke Google Drive.
 *
 * @function generatePdfReportFromDropdown
 * @throws {Error} Jika sheet yang diperlukan tidak ditemukan, nilai dropdown tidak dipilih, atau data tidak ditemukan.
 */
function generatePdfReportFromDropdown() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let newSheet;

  try {
    const sourceSheets = {
      dataSheet: spreadsheet.getSheetByName('GENERATEPDF'),
      templateSheet: spreadsheet.getSheetByName('TemplatePDF')
    };

    if (!sourceSheets.dataSheet) {
      throw new Error("Sheet 'GENERATEPDF' not found."); // Sheet 'GENERATEPDF' tidak ditemukan.
    }

    const allData = sourceSheets.dataSheet.getDataRange().getValues();
    const selectedDropdownValue = allData[0][0].toString().trim();
    const baseReportTitleDropdown = selectedDropdownValue.split('_').shift();
    const reportTitleDropdown = selectedDropdownValue; // Nilai dropdown lengkap
    const planId = selectedDropdownValue.split('_').pop(); // Ekstrak planId

    if (!reportTitleDropdown) {
      throw new Error('Please select a Report Title from the dropdown first.'); // Silakan pilih Judul Laporan dari dropdown terlebih dahulu.
    }

    const headers = allData[2];
    const headerColumnIndexMap = new Map(headers.map((header, index) => [header, index]));

    const filteredDataRow = allData.find((dataRow, rowIndex) => {
      if (rowIndex > 2) {
        const reportTitleValue = dataRow[headerColumnIndexMap.get('Judul Laporan A/R')].toString().trim();
        return reportTitleValue === baseReportTitleDropdown;
      }
      return false;
    });

    if (!filteredDataRow) {
      throw new Error('Data not found for Report Title: ' + reportTitleDropdown); // Data tidak ditemukan untuk Judul Laporan:
    }

    const mainReportData = prepareMainReportData(filteredDataRow, headerColumnIndexMap);
    validateReportData(mainReportData);

    const newSheetName = reportTitleDropdown.replace(/[^a-zA-Z0-9]/g, '').substring(0, 30);
    let existingSheet = spreadsheet.getSheetByName(newSheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    newSheet = sourceSheets.templateSheet.copyTo(spreadsheet).setName(newSheetName);

    applyBasicTemplateUpdates(newSheet, mainReportData, planId, sourceSheets.dataSheet, headerColumnIndexMap);
    if (mainReportData.detailEntries.length > 0) {
      processReportDetailEntries(newSheet, mainReportData.detailEntries);
    }

    const startTime = new Date().getTime();
    const pdfFile = generateAndSavePdfReport(newSheet, reportTitleDropdown, mainReportData);
    updatePdfLinkInSheet(spreadsheet, pdfFile);

    const endTime = new Date().getTime();
    Logger.log(`Execution time: ${(endTime - startTime) / 1000} seconds`); // Waktu eksekusi

  } catch (error) {
    Logger.log('Error: ' + error.toString()); // Error:
    SpreadsheetApp.getUi().alert('Error: ' + error.message); // Error:
    throw error;
  } finally {
    cleanupTemporarySheetAndReset(spreadsheet, newSheet);
  }
}

/**
 * Menerapkan pembaruan dasar pada sheet template dengan data laporan.
 *
 * @function applyBasicTemplateUpdates
 * @param {Sheet} targetSheet - Sheet yang akan diperbarui (salinan template).
 * @param {object} reportData - Data untuk laporan termasuk judul, nama, dan total.
 * @param {string} planId - ID rencana yang akan dimasukkan dalam laporan.
 * @param {Sheet} dataSheet - Sheet 'GENERATEPDF' untuk mengambil data tambahan.
 * @param {Map<string, number>} headerColumnIndexMap - Map header ke indeks kolom untuk pencarian data.
 */
function applyBasicTemplateUpdates(targetSheet, reportData, planId, dataSheet, headerColumnIndexMap) {
  const allDataGeneratePDF = dataSheet.getDataRange().getValues();
  let requestorPositionFromSheet = '';

  for (let i = 3; i < allDataGeneratePDF.length; i++) {
    const requestorNameSheet = allDataGeneratePDF[i][headerColumnIndexMap.get('Nama Requestor')];
    if (requestorNameSheet && requestorNameSheet.toString().trim() === reportData.requestorName.trim()) {
      requestorPositionFromSheet = allDataGeneratePDF[i][5]; // Kolom F (indeks 5) adalah 'Jabatan Requestor'
      if (requestorPositionFromSheet) {
        requestorPositionFromSheet = requestorPositionFromSheet.toString().trim();
      } else {
        requestorPositionFromSheet = ''; // Tangani jika 'Jabatan Requestor' kosong
      }
      break;
    }
  }

  const updates = [
    { range: "A6:N6", value: `Pada Hari, Tanggal ${reportData.endDateFormattedId}`, fontSize: 13 }, // Menetapkan tanggal dan hari
    { range: "A14:P14", value: reportData.reportTitleAr, fontSize: 16 }, // Judul laporan
    { range: "F10:K10", value: reportData.requestorName, fontSize: 13 }, // Nama pemohon
    { range: "F9", value: requestorPositionFromSheet, fontSize: 13 }, // Jabatan pemohon
    { range: "I16", value: reportData.organizationalUnit, fontSize: 13 }, // Unit organisasi
    { range: "I17", value: reportData.endDateFormattedId, fontSize: 13 }, // Tanggal akhir (diulang)
    { range: "I18", value: planId, fontSize: 13 }, // ID Rencana
    { range: "J19:K19", value: reportData.totalNominal, fontSize: 13, format: '#,##0' } // Nominal total
  ];

  updates.forEach(({ range, value, fontSize, format }) => {
    const rangeObj = targetSheet.getRange(range);
    rangeObj.setValue(value).setFontSize(fontSize);
    if (format) rangeObj.setNumberFormat(format);
  });

  PDF_REPORT_COLUMN_WIDTHS.forEach((width, index) => {
    targetSheet.setColumnWidth(index + 1, width);
  });
}

/**
 * Memproses dan memasukkan entri detail laporan ke dalam sheet target.
 *
 * @function processReportDetailEntries
 * @param {Sheet} targetSheet - Sheet tempat entri detail akan dimasukkan.
 * @param {Array<object>} detailEntries - Array objek entri detail.
 */
function processReportDetailEntries(targetSheet, detailEntries) {
  const dataStartRow = 24;
  const dataToInsert = detailEntries.map((entry, index) => [
    index + 1,
    entry.date,
    '',
    entry.description,
    '',
    'Transaksi',
    1,
    'Rp',
    entry.nominal,
    '',
    '',
    entry.account,
    '',
    ''
  ]);

  if (dataToInsert.length > 0) {
    const targetRange = targetSheet.getRange(dataStartRow, 1, dataToInsert.length, 14);
    targetRange.setValues(dataToInsert)
      .setFontSize(13)
      .setVerticalAlignment("middle");

    Logger.log("Data after setValues in sheet:"); // Data setelah setValues di sheet:
    Logger.log(targetSheet.getRange(dataStartRow, 1, dataToInsert.length, 14).getValues()); // Log data

    targetSheet.getRange(dataStartRow, 9, dataToInsert.length, 1)
      .setNumberFormat('#,##0');

    const wrapTextColumns = [4, 12];
    wrapTextColumns.forEach(columnIndex => {
      targetSheet.getRange(dataStartRow, columnIndex, dataToInsert.length, 1).setWrap(true);
    });

    targetSheet.getRange(dataStartRow - 1, 1, dataToInsert.length + 1, 14)
      .setBorder(true, true, true, true, true, true);

    targetSheet.getRange(dataStartRow, 8, dataToInsert.length, 1).setBorder(null, null, null, false, null, null);

    targetSheet.autoResizeRows(dataStartRow, dataToInsert.length);
  }
}

/**
 * Mempersiapkan objek data laporan utama dari baris data.
 *
 * @function prepareMainReportData
 * @param {Array<string>} dataRow - Baris data dari sheet 'GENERATEPDF'.
 * @param {Map<string, number>} headerColumnIndexMap - Map header ke indeks kolom.
 * @returns {object} Objek data laporan utama.
 */
function prepareMainReportData(dataRow, headerColumnIndexMap) {
  const splitAndTrim = (str) => str.toString().split('||').map(item => item.trim());
  const selectedDropdownValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GENERATEPDF').getRange('A1').getValue().toString().trim();
  const planId = selectedDropdownValue.split('_').pop();

  return {
    reportTitleAr: dataRow[headerColumnIndexMap.get('Judul Laporan A/R')],
    endDateFormattedId: formatDateToIndonesian(dataRow[headerColumnIndexMap.get('End Date A/R')]),
    requestorName: dataRow[headerColumnIndexMap.get('Nama Requestor')],
    organizationalUnit: dataRow[headerColumnIndexMap.get('Organizational Unit')],
    totalNominal: dataRow[headerColumnIndexMap.get('Nominal Total')],
    planId: planId,
    detailEntries: createReportDetailEntries(
      splitAndTrim(dataRow[headerColumnIndexMap.get('Tanggal Input')]),
      splitAndTrim(dataRow[headerColumnIndexMap.get('Nominal')]),
      splitAndTrim(dataRow[headerColumnIndexMap.get('Uraian')]),
      splitAndTrim(dataRow[headerColumnIndexMap.get('Account')])
    )
  };
}

/**
 * Memvalidasi data laporan utama untuk memastikan kolom wajib ada.
 *
 * @function validateReportData
 * @param {object} mainReportData - Objek data laporan utama yang akan divalidasi.
 * @throws {Error} Jika kolom data wajib tidak ada.
 */
function validateReportData(mainReportData) {
  if (!mainReportData.reportTitleAr || !mainReportData.endDateFormattedId) {
    throw new Error('Required data missing: Report Title and End Date must be filled.'); // Data wajib tidak lengkap: Judul Laporan dan Tanggal Akhir harus diisi.
  }
}

/**
 * Membuat array objek entri detail laporan dari array input.
 *
 * @function createReportDetailEntries
 * @param {string[]} dateStrings - Array string tanggal.
 * @param {string[]} nominalStrings - Array string nominal.
 * @param {string[]} descriptionStrings - Array string uraian.
 * @param {string[]} accountStrings - Array string akun.
 * @returns {Array<object>} Array objek entri detail.
 */
function createReportDetailEntries(dateStrings, nominalStrings, descriptionStrings, accountStrings) {
  const maxLength = Math.max(
    dateStrings.length,
    nominalStrings.length,
    descriptionStrings.length,
    accountStrings.length
  );

  return Array(maxLength).fill().map((_, i) => ({
    date: formatDateToIndonesian(dateStrings[i] || ''), // Format tanggal, gunakan string kosong jika tidak ada
    nominal: nominalStrings[i] || '', // Gunakan string kosong jika tidak ada nominal
    description: descriptionStrings[i] || '', // Gunakan string kosong jika tidak ada uraian
    account: accountStrings[i] || '' // Gunakan string kosong jika tidak ada akun
  }));
}

/**
 * Menghasilkan file PDF dari sheet laporan dan menyimpannya ke Google Drive.
 *
 * @function generateAndSavePdfReport
 * @param {Sheet} reportSheet - Sheet yang akan dihasilkan PDF-nya.
 * @param {string} reportTitleDropdown - Judul laporan (dari dropdown).
 * @param {object} mainReportData - Data laporan utama untuk penamaan file.
 * @returns {File} Objek file PDF yang dihasilkan di Google Drive.
 * @throws {Error} Jika pembuatan PDF gagal.
 */
function generateAndSavePdfReport(reportSheet, reportTitleDropdown, mainReportData) {
  const spreadsheet = reportSheet.getParent();
  const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);

  const startTime = new Date().getTime();

  reportSheet.activate();
  SpreadsheetApp.flush();
  Utilities.sleep(500);

  Logger.log('Sheet activated and flushed in: ' + (new Date().getTime() - startTime) + ' ms'); // Sheet diaktifkan dan di-flush dalam:

  const pdfExportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?` +
    `format=pdf&size=A4&portrait=true&fitw=true&gid=${reportSheet.getSheetId()}`;

  Logger.log('PDF URL generated in: ' + (new Date().getTime() - startTime) + ' ms'); // URL PDF dibuat dalam:

  const fetchOptions = {
    method: "GET",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };

  const fetchStartTime = new Date().getTime();
  const response = UrlFetchApp.fetch(pdfExportUrl, fetchOptions);
  Logger.log('PDF fetched in: ' + (new Date().getTime() - fetchStartTime) + ' ms'); // PDF di-fetch dalam:

  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to generate PDF file'); // Gagal membuat file PDF
  }

  // Penamaan file PDF, menggunakan judul laporan, tanggal, dan ID Rencana (planId)
  const formattedDateYyyyMmDd = convertIndonesianDateToYyyyMmDd(mainReportData.endDateFormattedId);
  const pdfFileName = `${mainReportData.reportTitleAr}_${formattedDateYyyyMmDd}_${mainReportData.planId}`
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '')
    .substring(0, 100)
    .trim() + '.pdf';

  const blobStartTime = new Date().getTime();
  const pdfBlob = response.getBlob().setName(pdfFileName);
  Logger.log('Blob processed in: ' + (new Date().getTime() - blobStartTime) + ' ms'); // Blob diproses dalam:

  const saveStartTime = new Date().getTime();
  const pdfFileOnDrive = pdfFolder.createFile(pdfBlob);
  Logger.log('File saved to Drive in: ' + (new Date().getTime() - saveStartTime) + ' ms'); // File disimpan ke Drive dalam:

  const endTime = new Date().getTime();
  Logger.log('Total PDF generation time: ' + (endTime - startTime) + ' ms'); // Total waktu pembuatan PDF:

  return pdfFileOnDrive; // Mengembalikan file PDF yang disimpan
}

/**
 * Memperbarui link PDF di sheet 'GENERATEPDF' di cell B2.
 *
 * @function updatePdfLinkInSheet
 * @param {Spreadsheet} spreadsheet - Spreadsheet aktif.
 * @param {File} pdfFile - Objek file PDF dari Google Drive.
 */
function updatePdfLinkInSheet(spreadsheet, pdfFile) {
  spreadsheet.getSheetByName('GENERATEPDF')
    .getRange('B2')
    .setFormula(`=HYPERLINK("${pdfFile.getUrl()}", "${pdfFile.getName()}")`); // Menetapkan formula HYPERLINK
}

/**
 * Membersihkan sheet sementara yang dibuat dan mereset sheet aktif ke 'GENERATEPDF'.
 *
 * @function cleanupTemporarySheetAndReset
 * @param {Spreadsheet} spreadsheet - Spreadsheet aktif.
 * @param {Sheet} temporarySheet - Sheet sementara yang akan dihapus.
 */
function cleanupTemporarySheetAndReset(spreadsheet, temporarySheet) {
  if (temporarySheet) {
    Utilities.sleep(1); // Jeda singkat sebelum menghapus sheet
    try {
      spreadsheet.deleteSheet(temporarySheet); // Hapus sheet sementara
    } catch (e) {
      Logger.log('Error deleting temporary sheet: ' + e.toString()); // Error saat menghapus sheet sementara:
    }
  }
  const configSheet = spreadsheet.getSheetByName('GENERATEPDF');
  if (configSheet) {
    spreadsheet.setActiveSheet(configSheet); // Set 'GENERATEPDF' sebagai sheet aktif
    configSheet.getRange('A1').activate(); // Pilih cell A1 di 'GENERATEPDF'
  }
}

/**
 * Mapping nama bulan Bahasa Indonesia ke nomor bulan untuk konversi tanggal.
 * @constant {object}
 */
const MONTH_INDONESIAN_TO_NUMBER_MAP = {
  'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
  'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
  'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
};

/**
 * Memformat string tanggal ke format tanggal Indonesia (DD Bulan YYYY).
 * Menerima berbagai format string tanggal dan mencoba memparsnya.
 *
 * @function formatDateToIndonesian
 * @param {string} dateString - String tanggal yang akan diformat.
 * @returns {string} Tanggal yang diformat sebagai DD Bulan YYYY (Indonesia).
 */
function formatDateToIndonesian(dateString) {
  const indonesianMonthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  try {
    let date;

    date = new Date(dateString);

    if (isNaN(date.getTime())) {
      const [dayMonthYearPart, timePart] = dateString.split(' ');
      const [day, month, year] = dayMonthYearPart.split('.').map(Number);
      date = new Date(year, month - 1, day);
    }

    if (!isNaN(date.getTime())) {
      return `${String(date.getDate()).padStart(2, '0')} ${indonesianMonthNames[date.getMonth()]} ${date.getFullYear()}`;
    }

    return dateString;
  } catch (error) {
    Logger.log(`Error formatting date: ${error}`); // Error memformat tanggal:
    return dateString;
  }
}

/**
 * Mengkonversi string tanggal Indonesia (DD Bulan YYYY) ke format YYYYMMDD.
 *
 * @function convertIndonesianDateToYyyyMmDd
 * @param {string} indonesianDateStr - String tanggal Indonesia (DD Bulan YYYY).
 * @returns {string} Tanggal yang diformat sebagai YYYYMMDD.
 * @throws {Error} Jika format tanggal tidak valid.
 */
function convertIndonesianDateToYyyyMmDd(indonesianDateStr) {
  try {
    const [day, month, year] = indonesianDateStr.split(' ');
    const monthNumber = MONTH_INDONESIAN_TO_NUMBER_MAP[month];
    if (!monthNumber) throw new Error('Invalid month: ' + month); // Bulan tidak valid:
    return `${year}${monthNumber}${String(day).padStart(2, '0')}`;
  } catch (e) {
    Logger.log('Date conversion error: ' + e.toString()); // Error konversi tanggal:
    throw new Error('Invalid date format: ' + indonesianDateStr); // Format tanggal tidak valid:
  }
}

/**
 * @function onOpen
 * @description Fungsi pemicu otomatis yang berjalan saat spreadsheet dibuka.
 *              Menambahkan menu kustom "GeneratePDF" ke menu bar spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GeneratePDF ✨📄') // Membuat menu dengan nama "GeneratePDF"
    .addItem('Generate PDF 📄⬇️ ', 'generatePdfReportFromDropdown') // Menambahkan item menu untuk menjalankan generatePdfReportFromDropdown
    .addToUi(); // Menambahkan menu ke UI spreadsheet.
}