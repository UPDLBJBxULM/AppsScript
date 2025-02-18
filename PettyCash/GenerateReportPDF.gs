/**
 * @fileOverview Generates PDF reports from Google Sheets based on dropdown selections,
 *             utilizing cached HTML templates for efficiency.
 */

/**
 * Configuration constants for HTML template caching.
 *
 * @constant {string} PDF_FOLDER_ID - ID of the Google Drive folder to store generated PDFs.
 * @constant {string} TEMPLATE_CACHE_PREFIX - Prefix for cache keys used for HTML templates.
 * @constant {number} TEMPLATE_CACHE_EXPIRATION_SECONDS - Cache expiration time in seconds for templates.
 */
const PDF_FOLDER_ID = '1VGU3E8Dv0o0vXs2JXEul-ge-WabHjlah';
const TEMPLATE_CACHE_PREFIX = 'TEMPLATE_';
const TEMPLATE_CACHE_EXPIRATION_SECONDS = 30;

/**
 * Static configuration for PDF appearance and export options.
 *
 * @constant {number[]} PDF_REPORT_COLUMN_WIDTHS - Array of column widths for the PDF report in points.
 * @constant {object} PDF_EXPORT_OPTIONS - Options object for PDF export settings.
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
 * Script cache instance for storing templates.
 * @type {Cache}
 */
const templateCache = CacheService.getScriptCache();

/**
 * Main function to generate a PDF report based on a dropdown selection in the 'GENERATEPDF' sheet.
 * It reads data, uses a template from 'TemplatePDF', populates it with data based on the selected
 * report title, and saves the generated PDF to Google Drive.
 *
 * @function generatePdfReportFromDropdown
 * @throws {Error} If required sheets are missing, dropdown value is not selected, or data is not found.
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
      throw new Error("Sheet 'GENERATEPDF' not found.");
    }

    const allData = sourceSheets.dataSheet.getDataRange().getValues();
    const selectedDropdownValue = allData[0][0].toString().trim();
    const baseReportTitleDropdown = selectedDropdownValue.split('_').shift();
    const reportTitleDropdown = selectedDropdownValue; // Full dropdown value
    const planId = selectedDropdownValue.split('_').pop(); // Extract planId

    if (!reportTitleDropdown) {
      throw new Error('Please select a Report Title from the dropdown first.');
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
      throw new Error('Data not found for Report Title: ' + reportTitleDropdown);
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
 * Applies basic updates to the template sheet with report data.
 *
 * @function applyBasicTemplateUpdates
 * @param {Sheet} targetSheet - The sheet to update (template copy).
 * @param {object} reportData - Data for the report including titles, names, and totals.
 * @param {string} planId - The plan ID to be included in the report.
 * @param {Sheet} dataSheet - The 'GENERATEPDF' sheet for retrieving additional data.
 * @param {Map<string, number>} headerColumnIndexMap - Map of headers to column indices for data lookup.
 */
function applyBasicTemplateUpdates(targetSheet, reportData, planId, dataSheet, headerColumnIndexMap) {
  const allDataGeneratePDF = dataSheet.getDataRange().getValues();
  let requestorPositionFromSheet = '';

  for (let i = 3; i < allDataGeneratePDF.length; i++) {
    const requestorNameSheet = allDataGeneratePDF[i][headerColumnIndexMap.get('Nama Requestor')];
    if (requestorNameSheet && requestorNameSheet.toString().trim() === reportData.requestorName.trim()) {
      requestorPositionFromSheet = allDataGeneratePDF[i][5]; // Column F (index 5) is 'Jabatan Requestor'
      if (requestorPositionFromSheet) {
        requestorPositionFromSheet = requestorPositionFromSheet.toString().trim();
      } else {
        requestorPositionFromSheet = ''; // Handle empty 'Jabatan Requestor'
      }
      break;
    }
  }

  const updates = [
    { range: "A6:N6", value: `Pada Hari, Tanggal ${reportData.endDateFormattedId}`, fontSize: 13 }, // Sets date and day
    { range: "A14:P14", value: reportData.reportTitleAr, fontSize: 16 }, // Report title
    { range: "F10:K10", value: reportData.requestorName, fontSize: 13 }, // Requestor name
    { range: "F9", value: requestorPositionFromSheet, fontSize: 13 }, // Requestor position
    { range: "I16", value: reportData.organizationalUnit, fontSize: 13 }, // Organizational unit
    { range: "I17", value: reportData.endDateFormattedId, fontSize: 13 }, // End date (repeated)
    { range: "I18", value: planId, fontSize: 13 }, // Plan ID
    { range: "J19:K19", value: reportData.totalNominal, fontSize: 13, format: '#,##0' } // Total nominal
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
 * Processes and inserts report detail entries into the target sheet.
 *
 * @function processReportDetailEntries
 * @param {Sheet} targetSheet - The sheet to insert detail entries into.
 * @param {Array<object>} detailEntries - Array of detail entry objects.
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

    Logger.log("Data after setValues in sheet:");
    Logger.log(targetSheet.getRange(dataStartRow, 1, dataToInsert.length, 14).getValues());

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
 * Prepares the main report data object from a data row.
 *
 * @function prepareMainReportData
 * @param {Array<string>} dataRow - Row of data from the 'GENERATEPDF' sheet.
 * @param {Map<string, number>} headerColumnIndexMap - Map of headers to column indices.
 * @returns {object} Main report data object.
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
 * Validates the main report data to ensure required fields are present.
 *
 * @function validateReportData
 * @param {object} mainReportData - The main report data object to validate.
 * @throws {Error} If required data fields are missing.
 */
function validateReportData(mainReportData) {
  if (!mainReportData.reportTitleAr || !mainReportData.endDateFormattedId) {
    throw new Error('Required data missing: Report Title and End Date must be filled.');
  }
}

/**
 * Creates an array of report detail entry objects from input arrays.
 *
 * @function createReportDetailEntries
 * @param {string[]} dateStrings - Array of date strings.
 * @param {string[]} nominalStrings - Array of nominal strings.
 * @param {string[]} descriptionStrings - Array of description strings.
 * @param {string[]} accountStrings - Array of account strings.
 * @returns {Array<object>} Array of detail entry objects.
 */
function createReportDetailEntries(dateStrings, nominalStrings, descriptionStrings, accountStrings) {
  const maxLength = Math.max(
    dateStrings.length,
    nominalStrings.length,
    descriptionStrings.length,
    accountStrings.length
  );

  return Array(maxLength).fill().map((_, i) => ({
    date: formatDateToIndonesian(dateStrings[i] || ''), // Format date, use empty string if missing
    nominal: nominalStrings[i] || '', // Use empty string if missing nominal
    description: descriptionStrings[i] || '', // Use empty string if missing description
    account: accountStrings[i] || '' // Use empty string if missing account
  }));
}

/**
 * Generates a PDF file from the report sheet and saves it to Google Drive.
 *
 * @function generateAndSavePdfReport
 * @param {Sheet} reportSheet - The sheet to generate the PDF from.
 * @param {string} reportTitleDropdown - The title of the report (from dropdown).
 * @param {object} mainReportData - Main report data for file naming.
 * @returns {File} The generated PDF file object in Google Drive.
 * @throws {Error} If PDF generation fails.
 */
function generateAndSavePdfReport(reportSheet, reportTitleDropdown, mainReportData) {
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
    throw new Error('Failed to generate PDF file');
  }

  // PDF file naming, using report title, date, and Plan ID
  const formattedDateYyyyMmDd = convertIndonesianDateToYyyyMmDd(mainReportData.endDateFormattedId);
  const pdfFileName = `${mainReportData.reportTitleAr}_${formattedDateYyyyMmDd}_${mainReportData.planId}`
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

  return pdfFileOnDrive; // Returns the saved PDF file
}

/**
 * Updates the PDF link in the 'GENERATEPDF' sheet in cell B2.
 *
 * @function updatePdfLinkInSheet
 * @param {Spreadsheet} spreadsheet - The active spreadsheet.
 * @param {File} pdfFile - The PDF file object from Google Drive.
 */
function updatePdfLinkInSheet(spreadsheet, pdfFile) {
  spreadsheet.getSheetByName('GENERATEPDF')
    .getRange('B2')
    .setFormula(`=HYPERLINK("${pdfFile.getUrl()}", "${pdfFile.getName()}")`); // Sets HYPERLINK formula
}

/**
 * Cleans up the temporary sheet and resets the active sheet to 'GENERATEPDF'.
 *
 * @function cleanupTemporarySheetAndReset
 * @param {Spreadsheet} spreadsheet - The active spreadsheet.
 * @param {Sheet} temporarySheet - The temporary sheet to be deleted.
 */
function cleanupTemporarySheetAndReset(spreadsheet, temporarySheet) {
  if (temporarySheet) {
    Utilities.sleep(1); // Short pause before deleting sheet
    try {
      spreadsheet.deleteSheet(temporarySheet); // Delete temporary sheet
    } catch (e) {
      Logger.log('Error deleting temporary sheet: ' + e.toString());
    }
  }
  const configSheet = spreadsheet.getSheetByName('GENERATEPDF');
  if (configSheet) {
    spreadsheet.setActiveSheet(configSheet); // Set 'GENERATEPDF' as active sheet
    configSheet.getRange('A1').activate(); // Select cell A1 in 'GENERATEPDF'
  }
}

/**
 * Mapping of Indonesian month names to month numbers for date conversion.
 * @constant {object}
 */
const MONTH_INDONESIAN_TO_NUMBER_MAP = {
  'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
  'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
  'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
};

/**
 * Formats a date string into Indonesian date format (DD Month YYYY).
 * Accepts various date string formats and attempts to parse them.
 *
 * @function formatDateToIndonesian
 * @param {string} dateString - Date string to format.
 * @returns {string} Date formatted as DD Month YYYY (Indonesian).
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
 * Converts an Indonesian date string (DD Month YYYY) to YYYYMMDD format.
 *
 * @function convertIndonesianDateToYyyyMmDd
 * @param {string} indonesianDateStr - Indonesian date string (DD Month YYYY).
 * @returns {string} Date formatted as YYYYMMDD.
 * @throws {Error} If the date format is invalid.
 */
function convertIndonesianDateToYyyyMmDd(indonesianDateStr) {
  try {
    const [day, month, year] = indonesianDateStr.split(' ');
    const monthNumber = MONTH_INDONESIAN_TO_NUMBER_MAP[month];
    if (!monthNumber) throw new Error('Invalid month: ' + month);
    return `${year}${monthNumber}${String(day).padStart(2, '0')}`;
  } catch (e) {
    Logger.log('Date conversion error: ' + e.toString());
    throw new Error('Invalid date format: ' + indonesianDateStr);
  }
}

/**
 * @function onOpen
 * @description Automatic trigger function that runs when the spreadsheet is opened.
 *              Adds a custom menu "GeneratePDF" to the spreadsheet menu bar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GeneratePDF ✨📄') // Creates menu named "GeneratePDF"
    .addItem('Generate PDF 📄⬇️ ', 'generatePdfReportFromDropdown') // Adds menu item to run generatePdfReportFromDropdown
    .addToUi(); // Adds menu to spreadsheet UI.
}