// Constants for configuration
const PDF_FOLDER_ID = '1VGU3E8Dv0o0vXs2JXEul-ge-WabHjlah';
const TEMPLATE_HAL2_NAME = 'TemplatePDF_hal2';
const TEMPLATE_HAL3_NAME = 'TemplatePDF_hal3';
const PDF_REPORT_COLUMN_WIDTHS = [30, 80, 80, 115, 115, 95, 70, 45, 45, 70, 95, 90, 30, 195, 75, 30];
const PDF_EXPORT_OPTIONS = {
  exportFormat: 'pdf', format: 'pdf', size: 'A4', portrait: true, fitw: true, scale: 2,
  top_margin: 0.5, bottom_margin: 0.5, left_margin: 0.5, right_margin: 0.5,
  sheetnames: false, printtitle: false, pagenumbers: false, gridlines: false, fzr: false
};
const MONTH_INDONESIAN_TO_NUMBER_MAP = {
  'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04', 'Mei': '05', 'Juni': '06',
  'Juli': '07', 'Agustus': '08', 'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
};
const DETAIL_DATA_START_ROW = 27;
const NOMINAL_PERMOHONAN_COLUMN_INDEX_RENCANA = 8;

// Global cache
let rencanaCache = null;

function initializeRencanaCache(spreadsheet) {
  if (!rencanaCache) {
    const rencanaSheet = spreadsheet.getSheetByName('RENCANA');
    if (!rencanaSheet) throw new Error("Sheet 'RENCANA' tidak ditemukan.");
    const rencanaData = rencanaSheet.getDataRange().getValues();
    rencanaCache = new Map();
    const header = rencanaData[0];
    const planIdIndex = header.indexOf('id rencana');
    const nominalIndex = header.indexOf('Nominal') || NOMINAL_PERMOHONAN_COLUMN_INDEX_RENCANA;
    for (let i = 1; i < rencanaData.length; i++) {
      const planId = rencanaData[i][planIdIndex]?.toString().trim();
      const nominal = rencanaData[i][nominalIndex] || 0;
      if (planId) rencanaCache.set(planId, nominal);
    }
  }
}

function generatePdfReportFromDropdown() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let newSheet;
  const startTime = new Date().getTime();
  Logger.log(`Mulai eksekusi: ${new Date().toISOString()}`);

  try {
    initializeRencanaCache(spreadsheet);
    const sourceSheets = getSourceSheets(spreadsheet);
    const { reportTitleDropdown, baseReportTitleDropdown, planId } = getDropdownValues(sourceSheets.dataSheet);
    const { headers, headerColumnIndexMap, filteredDataRow } = getReportData(sourceSheets.dataSheet, baseReportTitleDropdown, planId);
    const mainReportData = prepareMainReportData(filteredDataRow, headerColumnIndexMap, planId);
    validateReportData(mainReportData);

    newSheet = createNewSheet(spreadsheet, sourceSheets.templateSheet, reportTitleDropdown);
    applyBasicTemplateUpdates(newSheet, mainReportData, planId, sourceSheets.dataSheet, headerColumnIndexMap);
    if (mainReportData.detailEntries.length > 0) {
      processReportDetailEntries(newSheet, mainReportData.detailEntries);
    }

    const pdfBlobHal1 = generateAndSavePdfReport(newSheet, mainReportData.reportTitleAr, mainReportData, true, 'Hal1');
    Logger.log(`Selesai Hal1: ${((new Date().getTime() - startTime) / 1000)} detik`);

    const pdfFileHal2 = generatePdfReportHalPage(spreadsheet, baseReportTitleDropdown, planId, sourceSheets, TEMPLATE_HAL2_NAME, 'Hal2', 'Scan Nota', 'B3', mainReportData);
    Logger.log(`Selesai Hal2: ${((new Date().getTime() - startTime) / 1000)} detik`);

    const pdfFileHal3 = generatePdfReportHalPage(spreadsheet, baseReportTitleDropdown, planId, sourceSheets, TEMPLATE_HAL3_NAME, 'Hal3', 'Gambar Barang', 'B4', mainReportData);
    Logger.log(`Selesai Hal3: ${((new Date().getTime() - startTime) / 1000)} detik`);

    updatePdfLinks(spreadsheet, pdfBlobHal1, pdfFileHal2, pdfFileHal3);
    Logger.log(`Total waktu eksekusi: ${((new Date().getTime() - startTime) / 1000)} detik`);
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Terjadi kesalahan: ${error.message}`);
    throw error;
  } finally {
    cleanupTemporarySheetAndReset(spreadsheet, newSheet);
  }
}

function getSourceSheets(spreadsheet) {
  const sources = {
    dataSheet: spreadsheet.getSheetByName('GENERATEPDF'),
    templateSheet: spreadsheet.getSheetByName('TemplatePDF_hal1'),
    templateHal2Sheet: spreadsheet.getSheetByName(TEMPLATE_HAL2_NAME),
    templateHal3Sheet: spreadsheet.getSheetByName(TEMPLATE_HAL3_NAME),
    rekapRealisasiSheet: spreadsheet.getSheetByName('REKAPREALISASI')
  };
  const missingSheets = Object.entries(sources).filter(([_, sheet]) => !sheet).map(([name]) => name.replace('Sheet', ''));
  if (missingSheets.length > 0) throw new Error(`Sheet tidak ditemukan: ${missingSheets.join(', ')}`);
  return sources;
}

function getDropdownValues(dataSheet) {
  const allData = dataSheet.getDataRange().getValues();
  const selectedDropdownValue = allData[0][0]?.toString().trim();
  if (!selectedDropdownValue) throw new Error('Pilih Judul Laporan dari dropdown.');
  const parts = selectedDropdownValue.split('_');
  const baseReportTitleDropdown = parts[0];
  const planId = parts[parts.length - 1];
  return { reportTitleDropdown: selectedDropdownValue, baseReportTitleDropdown, planId };
}

function getReportData(dataSheet, baseReportTitleDropdown, planId) {
  const allData = dataSheet.getDataRange().getValues();
  const headers = allData[4];
  const headerColumnIndexMap = new Map(headers.map((header, index) => [header, index]));
  const requiredColumns = ['Judul Laporan A/R', 'ID Rencana'];
  requiredColumns.forEach(column => {
    if (!headerColumnIndexMap.has(column)) throw new Error(`Kolom wajib tidak ditemukan: ${column}`);
  });

  const filteredDataRow = allData.find((dataRow, rowIndex) => rowIndex > 4 &&
    dataRow[headerColumnIndexMap.get('Judul Laporan A/R')]?.toString().trim() === baseReportTitleDropdown &&
    dataRow[headerColumnIndexMap.get('ID Rencana')]?.toString().trim() === planId);

  if (!filteredDataRow) throw new Error(`Data tidak ditemukan untuk Judul: "${baseReportTitleDropdown}" dan ID: "${planId}"`);
  return { headers, headerColumnIndexMap, filteredDataRow };
}

function createNewSheet(spreadsheet, templateSheet, reportTitleDropdown) {
  const newSheetName = reportTitleDropdown.replace(/[^a-zA-Z0-9]/g, '').substring(0, 30);
  let existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  return templateSheet.copyTo(spreadsheet).setName(newSheetName);
}

function generatePdfReportHalPage(spreadsheet, baseReportTitleDropdown, planId, sourceSheets, templateHalName, pageNameSuffix, linkColumnName, linkCellRange, mainReportData) {
  const sourceSheetsLocal = { rekapRealisasiSheet: sourceSheets.rekapRealisasiSheet, templateHalSheet: spreadsheet.getSheetByName(templateHalName) };
  try {
    const { links, filteredRekapRealisasiRow } = getLinksForHalPage(sourceSheetsLocal.rekapRealisasiSheet, baseReportTitleDropdown, planId, linkColumnName);
    if (!filteredRekapRealisasiRow || links.length === 0) {
      Logger.log(`Tidak ada data atau link untuk ${pageNameSuffix}`);
      return null;
    }
    return processTemplatePagesForHal(sourceSheetsLocal.templateHalSheet, links, pageNameSuffix, spreadsheet, baseReportTitleDropdown, mainReportData)[0] || null;
  } catch (error) {
    Logger.log(`Error membuat PDF untuk ${pageNameSuffix}: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error membuat PDF untuk ${pageNameSuffix}: ${error.message}`);
    throw error;
  }
}

function getLinksForHalPage(rekapRealisasiSheet, baseReportTitleDropdown, planId, linkColumnName) {
  const allRekapRealisasiData = rekapRealisasiSheet.getDataRange().getValues();
  const rekapRealisasiHeader = allRekapRealisasiData[0];
  const rekapRealisasiHeaderMap = new Map(rekapRealisasiHeader.map((header, index) => [header, index]));
  const requiredColumns = ['Judul Laporan (A/R)', 'id rencana', linkColumnName];
  requiredColumns.forEach(column => {
    if (!rekapRealisasiHeaderMap.has(column)) throw new Error(`Kolom '${column}' tidak ditemukan di REKAPREALISASI`);
  });

  const judulLaporanColumnIndex = rekapRealisasiHeaderMap.get('Judul Laporan (A/R)');
  const idRencanaColumnIndex = rekapRealisasiHeaderMap.get('id rencana');
  const linkColumnIndex = rekapRealisasiHeaderMap.get(linkColumnName);

  const matchingRows = allRekapRealisasiData.filter((dataRow, rowIndex) => rowIndex > 0 &&
    dataRow[judulLaporanColumnIndex]?.toString().trim() === baseReportTitleDropdown &&
    dataRow[idRencanaColumnIndex]?.toString().trim() === planId);

  const links = matchingRows.flatMap(row => (row[linkColumnIndex] || '').toString().split(',')).map(link => link.trim()).filter(link => link);
  return { links, filteredRekapRealisasiRow: matchingRows[0] || null };
}

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

  PDF_REPORT_COLUMN_WIDTHS.forEach((width, index) => targetSheet.setColumnWidth(index + 1, width));
}

function getRequestorPosition(dataSheet, headerColumnIndexMap, requestorName) {
  const allData = dataSheet.getDataRange().getValues();
  for (let i = 5; i < allData.length; i++) {
    if (allData[i][headerColumnIndexMap.get('Nama Pemohon')]?.toString().trim() === requestorName.trim()) {
      return (allData[i][5] || '').toString().trim();
    }
  }
  return '';
}

function parseAccountableCombined(accountableCombined) {
  if (!accountableCombined) return { accountableName: '', accountablePosition: '' };
  const parts = accountableCombined.split('_');
  return { accountableName: parts.slice(0, -1).join('_'), accountablePosition: parts.pop() };
}

function processReportDetailEntries(targetSheet, detailEntries) {
  const dataToInsert = detailEntries.map((entry, index) => [
    index + 1, entry.date, '', entry.description, '', 'Transaksi', 1, 'Rp', entry.nominal, '', '', entry.account, '', ''
  ]);

  if (dataToInsert.length === 0) return;

  const targetRange = targetSheet.getRange(DETAIL_DATA_START_ROW, 1, dataToInsert.length, 14);
  targetRange.setValues(dataToInsert)
    .setFontSize(13)
    .setVerticalAlignment("middle")
    .setNumberFormats(Array(dataToInsert.length).fill(['@', '@', '@', '@', '@', '@', '@', '@', '#,##0', '@', '@', '@', '@', '@']));

  targetSheet.getRange(DETAIL_DATA_START_ROW, 4, dataToInsert.length, 1).setWrap(true);
  targetSheet.getRange(DETAIL_DATA_START_ROW, 12, dataToInsert.length, 1).setWrap(true);
  targetSheet.getRange(DETAIL_DATA_START_ROW - 1, 1, dataToInsert.length + 1, 14).setBorder(true, true, true, true, true, true);
  targetSheet.getRange(DETAIL_DATA_START_ROW, 8, dataToInsert.length, 1).setBorder(null, null, null, false, null, null);

  targetSheet.autoResizeRows(DETAIL_DATA_START_ROW, dataToInsert.length);
}

function prepareMainReportData(dataRow, headerColumnIndexMap, planId) {
  const splitAndTrim = (str) => (str || '').toString().split('||').map(item => item.trim());
  const accountableCombined = dataRow[headerColumnIndexMap.get('Nama Accountable')];
  const totalNominalPermohonan = getTotalNominalPermohonan(planId);

  return {
    reportTitleAr: dataRow[headerColumnIndexMap.get('Judul Laporan A/R')] || '',
    endDateFormattedId: formatDateToIndonesian(dataRow[headerColumnIndexMap.get('End Date A/R')] || ''),
    requestorName: dataRow[headerColumnIndexMap.get('Nama Pemohon')] || '',
    organizationalUnit: dataRow[headerColumnIndexMap.get('Organizational Unit')] || '',
    totalNominal: dataRow[headerColumnIndexMap.get('Nominal Total')] || 0,
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

function getTotalNominalPermohonan(planId) {
  return rencanaCache.get(planId) || 0;
}

function validateReportData(mainReportData) {
  if (!mainReportData.reportTitleAr || !mainReportData.endDateFormattedId) {
    throw new Error('Data wajib hilang: Judul Laporan dan Tanggal Akhir harus diisi.');
  }
}

function createReportDetailEntries(dateStrings, nominalStrings, descriptionStrings, accountStrings) {
  const maxLength = Math.max(dateStrings.length, nominalStrings.length, descriptionStrings.length, accountStrings.length);
  return Array(maxLength).fill().map((_, i) => ({
    date: formatDateToIndonesian(dateStrings[i] || ''),
    nominal: nominalStrings[i] || 0,
    description: descriptionStrings[i] || '',
    account: accountStrings[i] || ''
  }));
}

function generateAndSavePdfReport(reportSheet, baseReportTitle, mainReportData, saveToDrive = true, pageNameSuffix = 'Hal1') {
  const spreadsheet = reportSheet.getParent();
  const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
  reportSheet.activate();
  SpreadsheetApp.flush();

  const pdfExportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&size=A4&portrait=true&fitw=true&gid=${reportSheet.getSheetId()}`;
  const fetchOptions = { method: "GET", headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } };
  const response = UrlFetchApp.fetch(pdfExportUrl, fetchOptions);
  if (response.getResponseCode() !== 200) throw new Error('Gagal menghasilkan file PDF');

  const formattedDateYyyyMmDd = convertIndonesianDateToYyyyMmDd(mainReportData.endDateFormattedId);
  const pdfFileName = `${baseReportTitle}_${pageNameSuffix}_${formattedDateYyyyMmDd}_${mainReportData.planId}`
    .replace(/[^\w\s-]/g, '').replace(/\s+/g, '').substring(0, 100).trim() + '.pdf';

  const pdfBlob = response.getBlob().setName(pdfFileName);
  return saveToDrive ? pdfFolder.createFile(pdfBlob) : pdfBlob;
}

function processTemplatePagesForHal(templateSheet, links, pageNameSuffix, spreadsheet, baseReportTitleDropdown, mainReportData) {
  if (links.length === 0) return [];
  const MAX_IMAGES = 20;
  if (links.length > MAX_IMAGES) {
    Logger.log(`Terlalu banyak gambar (${links.length}). Dibatasi hingga ${MAX_IMAGES}.`);
    links = links.slice(0, MAX_IMAGES);
  }

  const newSheetName = `${baseReportTitleDropdown}_${pageNameSuffix}`;
  let existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);

  const numImageRows = Math.ceil(links.length / 2);
  for (let i = 0; i < numImageRows; i++) newSheet.setRowHeight(8 + i, 400);

  let currentRow = 8;
  let currentColumnIndex = 1;
  const startRow = 8;
  const imagesPerRow = 2;
  setRowVisibility(newSheet, links.length, startRow, imagesPerRow);

  links.forEach((link, index) => {
    const imageUrl = convertDriveLinkToImageUrl(link);
    if (!imageUrl.includes('drive.google.com')) {
      Logger.log(`URL gambar tidak valid: ${link}`);
      return;
    }
    const targetColumn = currentColumnIndex === 1 ? 2 : 10;
    try {
      const image = newSheet.insertImage(imageUrl, targetColumn, currentRow);
      if (image) {
        image.setWidth(390).setHeight(390);
        const anchorCell = newSheet.getRange(currentRow, targetColumn);
        image.setAnchorCell(anchorCell);
      } else {
        Logger.log(`Gagal menyisipkan gambar dari ${imageUrl}: Gambar null`);
      }
    } catch (error) {
      Logger.log(`Error saat menyisipkan gambar dari ${imageUrl}: ${error.message}`);
    }
    currentColumnIndex = currentColumnIndex === 1 ? 10 : 1;
    if (currentColumnIndex === 1) currentRow++;
  });

  const pdfFileHal = generateAndSavePdfReport(newSheet, mainReportData.reportTitleAr, mainReportData, true, pageNameSuffix);
  cleanupTemporarySheetAndReset(spreadsheet, newSheet);
  return [pdfFileHal];
}

function setRowVisibility(newSheet, numImages, startRow, imagesPerRow) {
  const rowsNeeded = Math.ceil(numImages / imagesPerRow);
  const maxRows = newSheet.getMaxRows();

  for (let i = 0; i < rowsNeeded; i++) {
    newSheet.unhideRow(newSheet.getRange(startRow + i, 1));
  }

  if (startRow + rowsNeeded <= maxRows) {
    newSheet.hideRows(startRow + rowsNeeded, maxRows - (startRow + rowsNeeded) + 1);
  }
}

function convertDriveLinkToImageUrl(driveLink) {
  const fileIdMatch = driveLink.match(/\/d\/([^\/]+)/);
  if (!fileIdMatch) throw new Error(`Link Drive tidak valid: ${driveLink}`);
  return `https://drive.google.com/uc?export=download&id=${fileIdMatch[1]}`;
}

function updatePdfLinks(spreadsheet, pdfFileHal1, pdfFileHal2, pdfFileHal3) {
  const dataSheet = spreadsheet.getSheetByName('GENERATEPDF');
  updatePdfLink(dataSheet, 'B2', pdfFileHal1);
  updatePdfLink(dataSheet, 'B3', pdfFileHal2);
  updatePdfLink(dataSheet, 'B4', pdfFileHal3);
}

function updatePdfLink(dataSheet, cellRange, pdfFile) {
  if (pdfFile) dataSheet.getRange(cellRange).setFormula(`=HYPERLINK("${pdfFile.getUrl()}", "${pdfFile.getName()}")`);
  else dataSheet.getRange(cellRange).clearContent();
}

function cleanupTemporarySheetAndReset(spreadsheet, temporarySheet) {
  if (temporarySheet) {
    try {
      spreadsheet.deleteSheet(temporarySheet);
    } catch (error) {
      Logger.log(`Error menghapus sheet sementara: ${error.message}`);
    }
  }
  const configSheet = spreadsheet.getSheetByName('GENERATEPDF');
  if (configSheet) {
    spreadsheet.setActiveSheet(configSheet);
    configSheet.getRange('A1').activate();
  }
}

function formatDateToIndonesian(dateString) {
  try {
    let date = new Date(dateString);
    if (isNaN(date.getTime())) {
      const [dayMonthYearPart] = dateString.split(' ');
      const [day, month, year] = dayMonthYearPart.split('.').map(Number);
      date = new Date(year, month - 1, day);
    }
    if (!isNaN(date.getTime())) {
      const indonesianMonthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
      return `${String(date.getDate()).padStart(2, '0')} ${indonesianMonthNames[date.getMonth()]} ${date.getFullYear()}`;
    }
    return dateString || '';
  } catch (error) {
    Logger.log(`Error memformat tanggal: ${error.message}`);
    return dateString || '';
  }
}

function convertIndonesianDateToYyyyMmDd(indonesianDateStr) {
  try {
    const [day, month, year] = indonesianDateStr.split(' ');
    const monthNumber = MONTH_INDONESIAN_TO_NUMBER_MAP[month];
    if (!monthNumber) throw new Error(`Bulan tidak valid: ${month}`);
    return `${year}${monthNumber}${String(day).padStart(2, '0')}`;
  } catch (error) {
    Logger.log(`Error konversi tanggal: ${error.message}`);
    throw new Error(`Format tanggal tidak valid: ${indonesianDateStr}`);
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('GeneratePDF 📄').addItem('Generate PDF 📄🔍', 'generatePdfReportFromDropdown').addToUi();
}