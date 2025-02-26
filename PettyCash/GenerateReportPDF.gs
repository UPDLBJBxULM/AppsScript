// Constants for configuration
const PDF_FOLDER_ID = '1VGU3E8Dv0o0vXs2JXEul-ge-WabHjlah';
const TEMPLATE_HAL2_NAME = 'TemplatePDF_hal2';
const TEMPLATE_HAL3_NAME = 'TemplatePDF_hal3';
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

const MONTH_INDONESIAN_TO_NUMBER_MAP = {
  'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
  'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
  'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
};

// Constants for sheet structure
const DETAIL_DATA_START_ROW = 27;
const NOMINAL_PERMOHONAN_COLUMN_INDEX_RENCANA = 8;

// Main function to generate PDF reports from dropdown selection
function generatePdfReportFromDropdown() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let newSheet;
  let pdfBlobHal1;
  let pdfFileHal2 = null;
  let pdfFileHal3 = null;

  try {
    // Validate and retrieve source sheets
    const sourceSheets = getSourceSheets(spreadsheet);
    const { reportTitleDropdown, baseReportTitleDropdown, planId } = getDropdownValues(sourceSheets.dataSheet);
    const { headers, headerColumnIndexMap, filteredDataRow } = getReportData(sourceSheets.dataSheet, baseReportTitleDropdown, planId);

    // Prepare report data
    const mainReportData = prepareMainReportData(filteredDataRow, headerColumnIndexMap, planId);
    validateReportData(mainReportData);

    // Create and update new sheet from template
    newSheet = createNewSheet(spreadsheet, sourceSheets.templateSheet, reportTitleDropdown);
    applyBasicTemplateUpdates(newSheet, mainReportData, planId, sourceSheets.dataSheet, headerColumnIndexMap);
    if (mainReportData.detailEntries.length > 0) {
      processReportDetailEntries(newSheet, mainReportData.detailEntries);
    }

    // Generate PDFs
    const startTime = new Date().getTime();
    pdfBlobHal1 = generateAndSavePdfReport(newSheet, reportTitleDropdown, mainReportData, true, 'Hal1');
    pdfFileHal2 = generatePdfReportHalPage(
      spreadsheet,
      baseReportTitleDropdown,
      planId, // Kirim planId
      sourceSheets,
      TEMPLATE_HAL2_NAME,
      'Hal2',
      'Scan Nota',
      'B3'
    );

    pdfFileHal3 = generatePdfReportHalPage(
      spreadsheet,
      baseReportTitleDropdown,
      planId, // Kirim planId
      sourceSheets,
      TEMPLATE_HAL3_NAME,
      'Hal3',
      'Gambar Barang',
      'B4'
    );

    // Update PDF links in sheet
    updatePdfLinks(spreadsheet, pdfBlobHal1, pdfFileHal2, pdfFileHal3);

    Logger.log(`Total Execution time: ${(new Date().getTime() - startTime) / 1000} seconds`);
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
    throw error;
  } finally {
    cleanupTemporarySheetAndReset(spreadsheet, newSheet);
  }
}

// Retrieve and validate source sheets
function getSourceSheets(spreadsheet) {
  const sources = {
    dataSheet: spreadsheet.getSheetByName('GENERATEPDF'),
    templateSheet: spreadsheet.getSheetByName('TemplatePDF_hal1'),
    templateHal2Sheet: spreadsheet.getSheetByName(TEMPLATE_HAL2_NAME),
    templateHal3Sheet: spreadsheet.getSheetByName(TEMPLATE_HAL3_NAME),
    rekapRealisasiSheet: spreadsheet.getSheetByName('REKAPREALISASI')
  };

  const missingSheets = Object.entries(sources)
    .filter(([_, sheet]) => !sheet)
    .map(([name]) => name.replace('Sheet', ''));

  if (missingSheets.length > 0) {
    throw new Error(`Missing sheets: ${missingSheets.join(', ')}`);
  }

  return sources;
}

// Retrieve dropdown values
function getDropdownValues(dataSheet) {
  const allData = dataSheet.getDataRange().getValues();
  const selectedDropdownValue = allData[0][0].toString().trim();
  if (!selectedDropdownValue) throw new Error('Please select a Report Title from the dropdown.');

  const parts = selectedDropdownValue.split('_');
  const baseReportTitleDropdown = parts[0];
  const planId = parts[parts.length - 1];

  return { reportTitleDropdown: selectedDropdownValue, baseReportTitleDropdown, planId };
}

// Retrieve report data
function getReportData(dataSheet, baseReportTitleDropdown, planId) {
  const allData = dataSheet.getDataRange().getValues();
  const headers = allData[4];
  const headerColumnIndexMap = new Map(headers.map((header, index) => [header, index]));

  // Validate required columns
  const requiredColumns = ['Judul Laporan A/R', 'ID Rencana'];
  requiredColumns.forEach(column => {
    if (!headerColumnIndexMap.has(column)) {
      throw new Error(`Missing required column: ${column}`);
    }
  });

  const filteredDataRow = allData.find((dataRow, rowIndex) => {
    if (rowIndex > 4) {
      const reportTitleValue = dataRow[headerColumnIndexMap.get('Judul Laporan A/R')].toString().trim();
      const idRencanaValue = dataRow[headerColumnIndexMap.get('ID Rencana')].toString().trim();
      return reportTitleValue === baseReportTitleDropdown && idRencanaValue === planId;
    }
    return false;
  });

  if (!filteredDataRow) {
    throw new Error(`Data not found for Report Title: "${baseReportTitleDropdown}" and Plan ID: "${planId}"`);
  }

  return { headers, headerColumnIndexMap, filteredDataRow };
}

// Create new sheet from template
function createNewSheet(spreadsheet, templateSheet, reportTitleDropdown) {
  const newSheetName = reportTitleDropdown
    .replace(/[^a-zA-Z0-9]/g, '')
    .substring(0, 30);
  let existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  return templateSheet.copyTo(spreadsheet).setName(newSheetName);
}

// Generate PDF for Hal2 or Hal3
// Modifikasi fungsi generatePdfReportHalPage
function generatePdfReportHalPage(
  spreadsheet,
  baseReportTitleDropdown,
  planId, // Tambahkan parameter planId
  sourceSheets,
  templateHalName,
  pageNameSuffix,
  linkColumnName,
  linkCellRange
) {
  const sourceSheetsLocal = {
    rekapRealisasiSheet: sourceSheets.rekapRealisasiSheet,
    templateHalSheet: spreadsheet.getSheetByName(templateHalName)
  };

  try {
    // Perubahan di sini: tambahkan planId
    const { links, filteredRekapRealisasiRow } = getLinksForHalPage(
      sourceSheetsLocal.rekapRealisasiSheet,
      baseReportTitleDropdown,
      planId, // Kirim planId ke fungsi
      linkColumnName
    );

    if (!filteredRekapRealisasiRow || links.length === 0) {
      Logger.log(`No data or links found for ${pageNameSuffix}`);
      return null;
    }

    const pdfFilesHal = processTemplatePagesForHal(
      sourceSheetsLocal.templateHalSheet,
      links,
      pageNameSuffix,
      spreadsheet,
      baseReportTitleDropdown
    );

    return pdfFilesHal && pdfFilesHal.length > 0 ? pdfFilesHal[0] : null;
  } catch (error) {
    Logger.log(`Error generating PDF for ${pageNameSuffix}: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error generating PDF for ${pageNameSuffix}: ${error.message}`);
    throw error;
  }
}

// Modifikasi fungsi getLinksForHalPage
function getLinksForHalPage(
  rekapRealisasiSheet,
  baseReportTitleDropdown,
  planId, // Tambahkan parameter planId
  linkColumnName
) {
  const allRekapRealisasiData = rekapRealisasiSheet.getDataRange().getValues();
  const rekapRealisasiHeader = allRekapRealisasiData[0];
  const rekapRealisasiHeaderMap = new Map(
    rekapRealisasiHeader.map((header, index) => [header, index])
  );

  // Validasi kolom yang diperlukan
  const requiredColumns = ['Judul Laporan (A/R)', 'id rencana', linkColumnName];
  requiredColumns.forEach(column => {
    if (!rekapRealisasiHeaderMap.has(column)) {
      throw new Error(`Kolom '${column}' tidak ditemukan di sheet REKAPREALISASI`);
    }
  });

  const judulLaporanColumnIndex = rekapRealisasiHeaderMap.get('Judul Laporan (A/R)');
  const idRencanaColumnIndex = rekapRealisasiHeaderMap.get('id rencana');
  const linkColumnIndex = rekapRealisasiHeaderMap.get(linkColumnName);

  const matchingRows = allRekapRealisasiData.filter((dataRow, rowIndex) => {
    if (rowIndex > 0) {
      const reportTitleValue = dataRow[judulLaporanColumnIndex].toString().trim();
      const idRencanaValue = dataRow[idRencanaColumnIndex].toString().trim();
      return (
        reportTitleValue === baseReportTitleDropdown &&
        idRencanaValue === planId // Filter tambahan untuk ID Rencana
      );
    }
    return false;
  });

  const links = matchingRows
    .flatMap(row => (row[linkColumnIndex] || '').toString().split(','))
    .map(link => link.trim())
    .filter(link => link !== '');

  return {
    links,
    filteredRekapRealisasiRow: matchingRows.length > 0 ? matchingRows[0] : null
  };
}

// Apply basic updates to template sheet
function applyBasicTemplateUpdates(targetSheet, reportData, planId, dataSheet, headerColumnIndexMap) {
  const requestorPosition = getRequestorPosition(dataSheet, headerColumnIndexMap, reportData.requestorName);
  const { accountableName, accountablePosition } = parseAccountableCombined(reportData.accountableCombined);

  const updates = [
    { range: "A6:N6", value: `Pada Hari, Tanggal ${reportData.endDateFormattedId}`, fontSize: 13 },
    { range: "A16:P16", value: reportData.reportTitleAr, fontSize: 16 },
    { range: "F10:K10", value: reportData.requestorName, fontSize: 13 },
    { range: "F9", value: requestorPosition, fontSize: 13 },
    { range: "F12:K12", value: accountableName, fontSize: 13 },
    { range: "F11:H11", value: accountablePosition, fontSize: 13 },
    { range: "I18", value: reportData.organizationalUnit, fontSize: 13 },
    { range: "I19", value: reportData.endDateFormattedId, fontSize: 13 },
    { range: "I20", value: planId, fontSize: 13 },
    { range: "J21:K21", value: reportData.totalNominalPermohonan, fontSize: 13, format: '#,##0' },
    { range: "J22:K22", value: reportData.totalNominal, fontSize: 13, format: '#,##0' },
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

// Retrieve requestor position from data sheet
function getRequestorPosition(dataSheet, headerColumnIndexMap, requestorName) {
  const allData = dataSheet.getDataRange().getValues();
  for (let i = 5; i < allData.length; i++) {
    const requestorNameSheet = allData[i][headerColumnIndexMap.get('Nama Pemohon')];
    if (requestorNameSheet && requestorNameSheet.toString().trim() === requestorName.trim()) {
      return (allData[i][5] || '').toString().trim();
    }
  }
  return '';
}

// Parse accountable combined field
function parseAccountableCombined(accountableCombined) {
  if (!accountableCombined) return { accountableName: '', accountablePosition: '' };
  const parts = accountableCombined.split('_');
  return {
    accountableName: parts.slice(0, parts.length - 1).join('_'),
    accountablePosition: parts.pop()
  };
}

// Process report detail entries
function processReportDetailEntries(targetSheet, detailEntries) {
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
    const targetRange = targetSheet.getRange(DETAIL_DATA_START_ROW, 1, dataToInsert.length, 14);
    targetRange.setValues(dataToInsert)
      .setFontSize(13)
      .setVerticalAlignment("middle");

    targetSheet.getRange(DETAIL_DATA_START_ROW, 9, dataToInsert.length, 1).setNumberFormat('#,##0');

    const wrapTextColumns = [4, 12];
    wrapTextColumns.forEach(columnIndex => {
      targetSheet.getRange(DETAIL_DATA_START_ROW, columnIndex, dataToInsert.length, 1).setWrap(true);
    });

    targetSheet.getRange(DETAIL_DATA_START_ROW - 1, 1, dataToInsert.length + 1, 14)
      .setBorder(true, true, true, true, true, true);

    targetSheet.getRange(DETAIL_DATA_START_ROW, 8, dataToInsert.length, 1).setBorder(null, null, null, false, null, null);

    targetSheet.autoResizeRows(DETAIL_DATA_START_ROW, dataToInsert.length);
  }
}

// Prepare main report data
function prepareMainReportData(dataRow, headerColumnIndexMap, planId) {
  const splitAndTrim = (str) => str.toString().split('||').map(item => item.trim());
  const accountableCombined = dataRow[headerColumnIndexMap.get('Nama Accountable')];

  const totalNominalPermohonan = getTotalNominalPermohonan(planId);

  return {
    reportTitleAr: dataRow[headerColumnIndexMap.get('Judul Laporan A/R')],
    endDateFormattedId: formatDateToIndonesian(dataRow[headerColumnIndexMap.get('End Date A/R')]),
    requestorName: dataRow[headerColumnIndexMap.get('Nama Pemohon')],
    organizationalUnit: dataRow[headerColumnIndexMap.get('Organizational Unit')],
    totalNominal: dataRow[headerColumnIndexMap.get('Nominal Total')],
    totalNominalPermohonan,
    planId,
    accountableCombined,
    detailEntries: createReportDetailEntries(
      splitAndTrim(dataRow[headerColumnIndexMap.get('Tanggal Input')]),
      splitAndTrim(dataRow[headerColumnIndexMap.get('Nominal')]),
      splitAndTrim(dataRow[headerColumnIndexMap.get('Uraian')]),
      splitAndTrim(dataRow[headerColumnIndexMap.get('Account')])
    )
  };
}

// Retrieve total nominal permohonan from RENCANA sheet
function getTotalNominalPermohonan(planId) {
  try {
    const rencanaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RENCANA');
    if (!rencanaSheet) throw new Error("Sheet 'RENCANA' not found.");
    const rencanaData = rencanaSheet.getDataRange().getValues();
    const rencanaHeader = rencanaData[0];
    const rencanaHeaderMap = new Map(rencanaHeader.map((header, index) => [header, index]));
    const planIdColumnIndex = rencanaHeaderMap.get('id rencana');
    if (planIdColumnIndex === undefined) throw new Error("Header 'id rencana' not found in 'RENCANA'.");

    const rencanaRow = rencanaData.find((row, rowIndex) => {
      if (rowIndex > 0 && row[planIdColumnIndex] && row[planIdColumnIndex].toString().trim() === planId) return true;
      return false;
    });

    return rencanaRow ? (rencanaRow[NOMINAL_PERMOHONAN_COLUMN_INDEX_RENCANA] || 0) : 0;
  } catch (error) {
    Logger.log(`Error fetching Total Nominal Permohonan: ${error.message}`);
    throw error;
  }
}

// Validate report data
function validateReportData(mainReportData) {
  if (!mainReportData.reportTitleAr || !mainReportData.endDateFormattedId) {
    throw new Error('Required data missing: Report Title and End Date must be filled.');
  }
}

// Create report detail entries
function createReportDetailEntries(dateStrings, nominalStrings, descriptionStrings, accountStrings) {
  const maxLength = Math.max(dateStrings.length, nominalStrings.length, descriptionStrings.length, accountStrings.length);
  return Array(maxLength).fill().map((_, i) => ({
    date: formatDateToIndonesian(dateStrings[i] || ''),
    nominal: nominalStrings[i] || '',
    description: descriptionStrings[i] || '',
    account: accountStrings[i] || ''
  }));
}

// Generate and save PDF report
function generateAndSavePdfReport(reportSheet, reportTitleDropdown, mainReportData, saveToDrive = true, pageNameSuffix = 'Hal1') {
  const spreadsheet = reportSheet.getParent();
  const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);

  reportSheet.activate();
  SpreadsheetApp.flush();

  const pdfExportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?` +
    `format=pdf&size=A4&portrait=true&fitw=true&gid=${reportSheet.getSheetId()}`;

  const fetchOptions = {
    method: "GET",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };

  const response = UrlFetchApp.fetch(pdfExportUrl, fetchOptions);
  if (response.getResponseCode() !== 200) throw new Error('Failed to generate PDF file');

  const formattedDateYyyyMmDd = convertIndonesianDateToYyyyMmDd(mainReportData.endDateFormattedId);
  const pdfFileName = `${mainReportData.reportTitleAr}_${pageNameSuffix}_${formattedDateYyyyMmDd}_${mainReportData.planId}`
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '')
    .substring(0, 100)
    .trim() + '.pdf';

  const pdfBlob = response.getBlob().setName(pdfFileName);
  return saveToDrive ? pdfFolder.createFile(pdfBlob) : pdfBlob;
}

// Process template pages for Hal2 or Hal3
function processTemplatePagesForHal(templateSheet, links, pageNameSuffix, spreadsheet, baseReportTitleDropdown) {
  if (links.length === 0) return [];

  const pdfFiles = [];
  const baseSheetName = templateSheet.getName().replace('TemplatePDF_', '');
  const newSheetName = `${baseReportTitleDropdown}_${baseSheetName}`;

  let existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);

  const numImageRows = Math.ceil(links.length / 2);

  for (let i = 0; i < numImageRows; i++) {
    newSheet.setRowHeight(8 + i, 400);
  }

  let currentRow = 8;
  let currentColumnIndex = 1;

  links.forEach((link, index) => {
    const imageUrl = convertDriveLinkToImageUrl(link);
    const targetColumn = currentColumnIndex === 1 ? 2 : 10;

    try {
      const image = newSheet.insertImage(imageUrl, targetColumn, currentRow);
      image.setWidth(390).setHeight(390);
      image.setAnchorCell(targetColumn, currentRow);
    } catch (error) {
      Logger.log(`Error inserting image from URL ${imageUrl}: ${error.message}`);
    }

    currentColumnIndex = currentColumnIndex === 1 ? 10 : 1;
    if (currentColumnIndex === 1) currentRow++;
  });

  const lastImageRow = currentRow;
  const maxRows = newSheet.getMaxRows();
  if (lastImageRow < maxRows) {
    const rowsToHide = maxRows - lastImageRow;
    newSheet.hideRows(lastImageRow + 1, rowsToHide);
  }

  const dummyMainReportData = {
    reportTitleAr: newSheetName,
    endDateFormattedId: formatDateToIndonesian(new Date()),
    planId: 'N/A'
  };

  const pdfFileHal = generateAndSavePdfReport(newSheet, newSheetName, dummyMainReportData, true, pageNameSuffix);
  pdfFiles.push(pdfFileHal);

  cleanupTemporarySheetAndReset(spreadsheet, newSheet);
  return pdfFiles;
}

// Convert Google Drive link to image URL
function convertDriveLinkToImageUrl(driveLink) {
  const fileId = driveLink.split('/d/')[1].split('/')[0];
  return `https://drive.google.com/uc?export=download&id=${fileId}`;
}

// Update PDF links in sheet
function updatePdfLinks(spreadsheet, pdfFileHal1, pdfFileHal2, pdfFileHal3) {
  const dataSheet = spreadsheet.getSheetByName('GENERATEPDF');
  updatePdfLink(dataSheet, 'B2', pdfFileHal1);
  updatePdfLink(dataSheet, 'B3', pdfFileHal2);
  updatePdfLink(dataSheet, 'B4', pdfFileHal3);
}

// Update single PDF link
function updatePdfLink(dataSheet, cellRange, pdfFile) {
  if (pdfFile) {
    dataSheet.getRange(cellRange)
      .setFormula(`=HYPERLINK("${pdfFile.getUrl()}", "${pdfFile.getName()}")`);
  } else {
    dataSheet.getRange(cellRange).clearContent();
  }
}

// Clean up temporary sheet and reset active sheet
function cleanupTemporarySheetAndReset(spreadsheet, temporarySheet) {
  if (temporarySheet) {
    try {
      spreadsheet.deleteSheet(temporarySheet);
    } catch (error) {
      Logger.log(`Error deleting temporary sheet: ${error.message}`);
    }
  }

  const configSheet = spreadsheet.getSheetByName('GENERATEPDF');
  if (configSheet) {
    spreadsheet.setActiveSheet(configSheet);
    configSheet.getRange('A1').activate();
  }
}

// Format date to Indonesian format
function formatDateToIndonesian(dateString) {
  try {
    let date;
    date = new Date(dateString);
    if (isNaN(date.getTime())) {
      const [dayMonthYearPart, timePart] = dateString.split(' ');
      const [day, month, year] = dayMonthYearPart.split('.').map(Number);
      date = new Date(year, month - 1, day);
    }
    if (!isNaN(date.getTime())) {
      const indonesianMonthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
      return `${String(date.getDate()).padStart(2, '0')} ${indonesianMonthNames[date.getMonth()]} ${date.getFullYear()}`;
    }
    return dateString;
  } catch (error) {
    Logger.log(`Error formatting date: ${error.message}`);
    return dateString;
  }
}

// Convert Indonesian date to YYYYMMDD format
function convertIndonesianDateToYyyyMmDd(indonesianDateStr) {
  try {
    const [day, month, year] = indonesianDateStr.split(' ');
    const monthNumber = MONTH_INDONESIAN_TO_NUMBER_MAP[month];
    if (!monthNumber) throw new Error(`Invalid month: ${month}`);
    return `${year}${monthNumber}${String(day).padStart(2, '0')}`;
  } catch (error) {
    Logger.log(`Date conversion error: ${error.message}`);
    throw new Error(`Invalid date format: ${indonesianDateStr}`);
  }
}

// Initialize menu on spreadsheet open
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GeneratePDF 📄')
    .addItem('Generate PDF 📄🔍', 'generatePdfReportFromDropdown')
    .addToUi();
}