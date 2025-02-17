/**
 * Konfigurasi cache untuk template HTML.
 * Cache digunakan untuk menyimpan template sementara agar proses lebih cepat.
 */
const PDF_FOLDER_ID = '1VGU3E8Dv0o0vXs2JXEul-ge-WabHjlah';
const TEMPLATE_CACHE_PREFIX = 'TEMPLATE_';
const TEMPLATE_CACHE_EXPIRATION_SECONDS = 30;

/**
 * Konfigurasi statis untuk tampilan dan ekspor PDF.
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
 * Inisialisasi cache script untuk menyimpan template.
 */
const templateCache = CacheService.getScriptCache();

/**
 * Fungsi utama untuk menghasilkan laporan PDF berdasarkan pilihan dropdown di Spreadsheet.
 * Fungsi ini membaca data dari sheet 'GENERATEPDF', menggunakan template dari 'TemplatePDF',
 * mengisi data berdasarkan Judul Laporan A/R yang dipilih dari dropdown, dan menyimpan PDF ke Google Drive.
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
      throw new Error("Sheet 'GENERATEPDF' tidak ditemukan.");
    }

    const allData = sourceSheets.dataSheet.getDataRange().getValues();
    const selectedDropdownValue = allData[0][0].toString().trim();
    const baseReportTitleDropdown = selectedDropdownValue.split('_').shift();
    const reportTitleDropdown = selectedDropdownValue; // Simpan nilai dropdown lengkap
    const planId = selectedDropdownValue.split('_').pop(); // Ekstrak planId dari nilai dropdown

    if (!reportTitleDropdown) {
      throw new Error('Silakan pilih Judul Laporan A/R dari dropdown terlebih dahulu.');
    }

    const headers = allData[2];
    const headerColumnIndexMap = new Map(headers.map((header, index) => [header, index]));

    const filteredDataRow = allData.find((dataRow, rowIndex) => {
      if (rowIndex > 2) {
        const judulLaporanARValue = dataRow[headerColumnIndexMap.get('Judul Laporan A/R')].toString().trim();
        return judulLaporanARValue === baseReportTitleDropdown;
      }
      return false;
    });

    if (!filteredDataRow) {
      throw new Error('Data tidak ditemukan untuk Judul Laporan A/R: ' + reportTitleDropdown);
    }

    const mainReportData = prepareMainReportData(filteredDataRow, headerColumnIndexMap);
    validateReportData(mainReportData);

    const newSheetName = reportTitleDropdown.replace(/[^a-zA-Z0-9]/g, '').substring(0, 30);
    let existingSheet = spreadsheet.getSheetByName(newSheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    newSheet = sourceSheets.templateSheet.copyTo(spreadsheet).setName(newSheetName);

    // Gunakan planId yang diekstrak untuk "Nomor" dan reportTitleDropdown untuk nama sheet dan PDF
    applyBasicTemplateUpdates(newSheet, mainReportData, planId);
    if (mainReportData.detailEntries.length > 0) {
      processReportDetailEntries(newSheet, mainReportData.detailEntries);
    }

    const startTime = new Date().getTime();
    const pdfFile = generateAndSavePdfReport(newSheet, reportTitleDropdown, mainReportData);
    updatePdfLinkInSheet(spreadsheet, pdfFile);

    const endTime = new Date().getTime();
    Logger.log(`Execution time: ${(endTime - startTime) / 1000} seconds`);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    throw error;
  } finally {
    cleanupTemporarySheetAndReset(spreadsheet, newSheet);
  }
}

/**
 * Menerapkan update dasar pada sheet template laporan.
 */
function applyBasicTemplateUpdates(targetSheet, reportData, planId) {
  const updates = [
    { range: "A6:N6", value: `Pada Hari, Tanggal ${reportData.endDateFormattedId}`, fontSize: 13 },
    { range: "A14:P14", value: reportData.reportTitleAr, fontSize: 16 },
    { range: "F10:K10", value: reportData.requestorName, fontSize: 13 },
    { range: "F9", value: reportData.organizationalUnit, fontSize: 13 },
    { range: "I16", value: reportData.organizationalUnit, fontSize: 13 },
    { range: "I17", value: reportData.endDateFormattedId, fontSize: 13 },
    { range: "I18", value: planId, fontSize: 13 },
    { range: "J19:K19", value: reportData.totalNominal, fontSize: 13, format: '#,##0' }
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
 * Memproses dan memasukkan entri detail laporan ke dalam sheet.
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
 * Mempersiapkan data utama laporan dari baris data input.
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
 * Validasi data utama laporan.
 */
function validateReportData(reportData) {
  if (!reportData.reportTitleAr || !reportData.endDateFormattedId) {
    throw new Error('Data wajib tidak lengkap: Judul Laporan dan Tanggal Akhir harus diisi.');
  }
}

/**
 * Membuat array entri detail laporan.
 */
function createReportDetailEntries(dateStrings, nominalStrings, descriptionStrings, accountStrings) {
  const maxLength = Math.max(
    dateStrings.length,
    nominalStrings.length,
    descriptionStrings.length,
    accountStrings.length
  );

  return Array(maxLength).fill().map((_, i) => ({
    date: formatDateToIndonesian(dateStrings[i] || ''),
    nominal: nominalStrings[i] || '',
    description: descriptionStrings[i] || '',
    account: accountStrings[i] || ''
  }));
}

/**
 * Menghasilkan file PDF dari sheet laporan dan menyimpannya ke Google Drive.
 */
function generateAndSavePdfReport(reportSheet, reportId, reportData) {
  const spreadsheet = reportSheet.getParent();
  const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);

  const startTime = new Date().getTime();

  reportSheet.activate();
  SpreadsheetApp.flush();
  Utilities.sleep(500);

  Logger.log('Sheet activated and flushed in: ' + (new Date().getTime() - startTime) + ' ms');

  const pdfExportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?` +
    `format=pdf&size=A4&portrait=true&fitw=true&gid=${reportSheet.getSheetId()}`;

  Logger.log('PDF URL generated in: ' + (new Date().getTime() - startTime) + ' ms');

  const fetchOptions = {
    method: "GET",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };

  const fetchStartTime = new Date().getTime();
  const response = UrlFetchApp.fetch(pdfExportUrl, fetchOptions);
  Logger.log('PDF fetched in: ' + (new Date().getTime() - fetchStartTime) + ' ms');

  if (response.getResponseCode() !== 200) {
    throw new Error('Gagal membuat file PDF');
  }

  // Penamaan file PDF sekarang menggunakan Judul Laporan, Tanggal, dan ID Rencana (planId)
  const formattedDateYyyyMmDd = convertIndonesianDateToYyyyMmDd(reportData.endDateFormattedId);
  const pdfFileName = `${reportData.reportTitleAr}_${formattedDateYyyyMmDd}_${reportData.planId}`
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '')
    .substring(0, 100)
    .trim() + '.pdf';

  const blobStartTime = new Date().getTime();
  const pdfBlob = response.getBlob().setName(pdfFileName);
  Logger.log('Blob processed in: ' + (new Date().getTime() - blobStartTime) + ' ms');

  const saveStartTime = new Date().getTime();
  const pdfFileOnDrive = pdfFolder.createFile(pdfBlob);
  Logger.log('File saved to Drive in: ' + (new Date().getTime() - saveStartTime) + ' ms');

  const endTime = new Date().getTime();
  Logger.log('Total PDF generation time: ' + (endTime - startTime) + ' ms');

  return pdfFileOnDrive;
}

/**
 * Update link PDF pada sheet 'GENERATEPDF' di kolom B2.
 */
function updatePdfLinkInSheet(spreadsheet, pdfFile) {
  spreadsheet.getSheetByName('GENERATEPDF')
    .getRange('B2')
    .setFormula(`=HYPERLINK("${pdfFile.getUrl()}", "${pdfFile.getName()}")`);
}

/**
 * Membersihkan sheet sementara yang dibuat dan reset sheet aktif ke 'GENERATEPDF'.
 */
function cleanupTemporarySheetAndReset(spreadsheet, temporarySheet) {
  if (temporarySheet) {
    Utilities.sleep(1);
    try {
      spreadsheet.deleteSheet(temporarySheet);
    } catch (e) {
      Logger.log('Error deleting temporary sheet: ' + e.toString());
    }
  }
  const configSheet = spreadsheet.getSheetByName('GENERATEPDF');
  if (configSheet) {
    spreadsheet.setActiveSheet(configSheet);
    configSheet.getRange('A1').activate();
  }
}

/**
 * Mapping bulan bahasa Indonesia ke angka bulan (untuk konversi tanggal).
 */
const MONTH_INDONESIAN_TO_NUMBER_MAP = {
  'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
  'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
  'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
};

/**
 * Format string tanggal ke format tanggal Indonesia (DD Bulan YYYY).
 * Menerima berbagai format string tanggal dan mencoba memparsnya.
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
    Logger.log(`Error formatting date: ${error}`);
    return dateString;
  }
}

/**
 * Konversi tanggal format Indonesia (DD Bulan YYYY) ke format YYYYMMDD.
 */
function convertIndonesianDateToYyyyMmDd(indonesianDateStr) {
  try {
    const [day, month, year] = indonesianDateStr.split(' ');
    const monthNumber = MONTH_INDONESIAN_TO_NUMBER_MAP[month];
    if (!monthNumber) throw new Error('Bulan tidak valid: ' + month);
    return `${year}${monthNumber}${String(day).padStart(2, '0')}`;
  } catch (e) {
    Logger.log('Date conversion error: ' + e.toString());
    throw new Error('Format tanggal tidak valid: ' + indonesianDateStr);
  }
}