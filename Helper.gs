function getActiveSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetByName(sheetName) {
  return getActiveSpreadsheet().getSheetByName(sheetName);
}

function ensureSheet(sheetName, headers) {
  let sheet = getSheetByName(sheetName);
  if (!sheet) {
    sheet = getActiveSpreadsheet().insertSheet(sheetName);
    if (headers && headers.length) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

function normalizeStatus(status) {
  const value = (status || 'Proses').toString().trim();
  return value || 'Proses';
}

function formatDateIndo(dateValue) {
  if (dateValue instanceof Date && !isNaN(dateValue)) {
    return `${dateValue.getDate()} ${MONTH_NAMES[dateValue.getMonth()]} ${dateValue.getFullYear()}`;
  }
  return dateValue;
}

function getYearFromValue(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value.getFullYear().toString();
  }
  const match = String(value || '').match(/\d{4}/);
  return match ? match[0] : '';
}
