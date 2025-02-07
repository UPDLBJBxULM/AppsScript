// Cache configuration
const FOLDER_ID = '1VGU3E8Dv0o0vXs2JXEul-ge-WabHjlah';
const CACHE_KEY_PREFIX = 'TEMPLATE_';
const CACHE_EXPIRATION = 30;

// Predefine static configurations
const COLUMN_WIDTHS = [30, 80, 80, 115, 115, 95, 70, 45, 45, 70, 95, 90, 30, 195, 75, 30];
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

// Cache template HTML content
const templateCache = CacheService.getScriptCache();

function generateReportFromDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let newSheet;

  try {
    // Get all required sheets in one batch
    const sheets = {
      data: ss.getSheetByName('GENERATEPDF'),
      template: ss.getSheetByName('TemplatePDF')
    };

    if (!sheets.data) throw new Error("Sheet 'GENERATEPDF' tidak ditemukan.");

    // Read all data at once and process in memory
    const allData = sheets.data.getDataRange().getValues();
    const selectedValue = allData[0][0].toString().trim();
    const idRencana = selectedValue.split('_').pop();

    if (!idRencana) throw new Error('Silakan pilih ID Rencana terlebih dahulu.');

    // Process data in memory
    const headers = allData[2];
    const columnMap = new Map(headers.map((header, idx) => [header, idx]));
    const filteredData = allData.find((row, idx) =>
      idx > 2 && row[columnMap.get('ID Rencana')].toString().trim() === idRencana
    );

    if (!filteredData) throw new Error('Data tidak ditemukan untuk ID: ' + idRencana);

    // Prepare all data in memory
    const mainData = prepareMainData(filteredData, columnMap);
    validateData(mainData);

    // Create new sheet with optimized template
    const newSheetName = idRencana;
    const existingSheet = ss.getSheetByName(newSheetName);
    if (existingSheet) ss.deleteSheet(existingSheet);

    newSheet = sheets.template.copyTo(ss).setName(newSheetName);

    // Apply basic updates
    applyBasicUpdates(newSheet, mainData, idRencana);

    // Process detailed entries
    if (mainData.entries.length > 0) {
      processDetailedEntries(newSheet, mainData.entries);
    }

    const startTime = new Date().getTime();
    // Generate and save PDF
    const pdfFile = generateAndSavePDF(newSheet, idRencana, mainData);
    updatePDFLink(ss, pdfFile);

    const endTime = new Date().getTime();
    Logger.log(`Execution time: ${(endTime - startTime) / 1000} seconds`);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    throw error;
  } finally {
    cleanupAndReset(ss, newSheet);
  }
}

function applyBasicUpdates(sheet, mainData, idRencana) {
  const updates = [
    { range: "A6:N6", value: `Pada Hari, Tanggal ${mainData.endDateAr} ditetapkan oleh dan sebagai berikut`, fontSize: 13 },
    { range: "A14:P14", value: mainData.judulLaporanAr, fontSize: 16 },
    { range: "F10:K10", value: mainData.namaRequestor, fontSize: 13 },
    { range: "F9", value: mainData.orgUnit, fontSize: 13 },
    { range: "I16", value: mainData.orgUnit, fontSize: 13 },
    { range: "I17", value: mainData.endDateAr, fontSize: 13 },
    { range: "I18", value: idRencana, fontSize: 13 },
    { range: "J19:K19", value: mainData.nominalTotal, fontSize: 13, format: '#,##0' }
  ];

  // Apply updates in batch
  updates.forEach(({ range, value, fontSize, format }) => {
    const rangeObj = sheet.getRange(range);
    rangeObj.setValue(value).setFontSize(fontSize);
    if (format) rangeObj.setNumberFormat(format);
  });

  // Set column widths in batch
  COLUMN_WIDTHS.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
}

function processDetailedEntries(sheet, entries) {
  const startRow = 24;
  const dataToInsert = entries.map((entry, index) => [
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
    // Insert data in batch
    const targetRange = sheet.getRange(startRow, 1, dataToInsert.length, 14);
    targetRange.setValues(dataToInsert)
      .setFontSize(13)
      .setVerticalAlignment("middle");

    // Apply number format to nominal column
    sheet.getRange(startRow, 9, dataToInsert.length, 1)
      .setNumberFormat('#,##0');

    // Apply text wrapping ke kolom Uraian (4) dan Account (12)
    const wrapColumns = [4, 12];
    wrapColumns.forEach(col => {
      sheet.getRange(startRow, col, dataToInsert.length, 1).setWrap(true);
    });

    // Terapkan borders setelah pembungkusan teks
    sheet.getRange(startRow - 1, 1, dataToInsert.length + 1, 14)
      .setBorder(true, true, true, true, true, true);

    sheet.getRange(startRow, 8, dataToInsert.length, 1).setBorder(null, null, null, false, null, null);

    // Sesuaikan tinggi baris agar teks terbungkus terlihat
    sheet.autoResizeRows(startRow, dataToInsert.length);

    // Logging untuk memastikan pembungkusan teks diterapkan
    //Logger.log(`Pembungkusan teks diterapkan pada baris ${startRow} hingga ${startRow + dataToInsert.length - 1}`);
  }
}

function prepareMainData(row, columnMap) {
  const splitAndTrim = str => str.toString().split('||').map(item => item.trim());

  return {
    judulLaporanAr: row[columnMap.get('Judul Laporan A/R')],
    endDateAr: formatDateString(row[columnMap.get('End Date A/R')]),
    namaRequestor: row[columnMap.get('Nama Requestor')],
    orgUnit: row[columnMap.get('Organizational Unit')],
    nominalTotal: row[columnMap.get('Nominal Total')],
    entries: processEntries(
      splitAndTrim(row[columnMap.get('Tanggal Input')]),
      splitAndTrim(row[columnMap.get('Nominal')]),
      splitAndTrim(row[columnMap.get('Uraian')]),
      splitAndTrim(row[columnMap.get('Account')])
    )
  };
}

function validateData(mainData) {
  if (!mainData.judulLaporanAr || !mainData.endDateAr) {
    throw new Error('Data wajib tidak lengkap.');
  }
}

function processEntries(dates, nominals, descriptions, accounts) {
  const maxLength = Math.max(
    dates.length,
    nominals.length,
    descriptions.length,
    accounts.length
  );

  return Array(maxLength).fill().map((_, i) => ({
    date: formatDateString(dates[i] || ''),
    nominal: nominals[i] || '',
    description: descriptions[i] || '',
    account: accounts[i] || ''
  }));
}

function generateAndSavePDF(sheet, idRencana, mainData) {
  const ss = sheet.getParent();
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // Mulai logging waktu
  const startTime = new Date().getTime();

  // Aktifkan sheet untuk memastikan fokus
  sheet.activate();
  SpreadsheetApp.flush(); // Pastikan perubahan disimpan ke server
  Utilities.sleep(500);  // Tunggu sedikit untuk sinkronisasi

  // Logging waktu flushing
  Logger.log('Sheet activated and flushed in: ' + (new Date().getTime() - startTime) + ' ms');

  // URL untuk ekspor ke PDF
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    `format=pdf&size=A4&portrait=true&fitw=true&gid=${sheet.getSheetId()}`;
  
  Logger.log('PDF URL generated in: ' + (new Date().getTime() - startTime) + ' ms');

  // Konfigurasi Fetch
  const options = {
    method: "GET",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };

  // Mulai waktu untuk fetch
  const fetchStartTime = new Date().getTime();
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('PDF fetched in: ' + (new Date().getTime() - fetchStartTime) + ' ms');

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF generation failed');
  }

  // Penamaan file
  const formattedDate = convertIndonesianDateToYYYYMMDD(mainData.endDateAr);
  const filename = `${mainData.judulLaporanAr}_${formattedDate}_${idRencana}`
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '')
    .trim() + '.pdf';
  
  // Logging waktu untuk memproses blob
  const blobStartTime = new Date().getTime();
  const pdfBlob = response.getBlob().setName(filename);
  Logger.log('Blob processed in: ' + (new Date().getTime() - blobStartTime) + ' ms');

  // Menyimpan file ke folder
  const saveStartTime = new Date().getTime();
  const file = folder.createFile(pdfBlob);
  Logger.log('File saved to Drive in: ' + (new Date().getTime() - saveStartTime) + ' ms');

  // Logging total waktu
  const endTime = new Date().getTime();
  Logger.log('Total PDF generation time: ' + (endTime - startTime) + ' ms');

  return file;
}


function updatePDFLink(ss, file) {
  ss.getSheetByName('GENERATEPDF')
    .getRange('B2')
    .setFormula(`=HYPERLINK("${file.getUrl()}", "${file.getName()}")`);
}

function cleanupAndReset(ss, sheet) {
  if (sheet) {
    Utilities.sleep(1);
    try {
      ss.deleteSheet(sheet);
    } catch (e) {
      Logger.log('Error deleting sheet: ' + e.toString());
    }
  }
  const dataSheet = ss.getSheetByName('GENERATEPDF');
  if (dataSheet) {
    ss.setActiveSheet(dataSheet);
    dataSheet.getRange('A1').activate();
  }
}

// Existing date formatting functions
const MONTH_MAPPING = {
  'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
  'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
  'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
};

function formatDateString(dateString) {
  const monthsIndonesian = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  try {
    let date;

    // First, try direct parsing
    date = new Date(dateString);

    // If direct parsing fails, try splitting by dot
    if (isNaN(date.getTime())) {
      const [dayMonthYear, time] = dateString.split(' ');
      const [day, month, year] = dayMonthYear.split('.').map(Number);
      date = new Date(year, month - 1, day);
    }

    // Log for debugging
    // Logger.log(`Original dateString: ${dateString}`);
    // Logger.log(`Parsed date: ${date}`);

    if (!isNaN(date.getTime())) {
      return `${String(date.getDate()).padStart(2, '0')} ${monthsIndonesian[date.getMonth()]} ${date.getFullYear()}`;
    }

    return dateString;
  } catch (error) {
    Logger.log(`Error formatting date: ${error}`);
    return dateString;
  }
}

function convertIndonesianDateToYYYYMMDD(dateStr) {
  try {
    const [day, month, year] = dateStr.split(' ');
    const monthNum = MONTH_MAPPING[month];
    if (!monthNum) throw new Error('Invalid month: ' + month);
    return `${year}${monthNum}${String(day).padStart(2, '0')}`;
  } catch (e) {
    Logger.log('Date conversion error: ' + e.toString());
    throw new Error('Invalid date format: ' + dateStr);
  }
}