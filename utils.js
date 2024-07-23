/**
 * Gets the column index for a given header name.
 * @param {string} headerName The name of the header to find.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
 * @return {number|null} The column index of the header, or null if not found.
 */
function getColumnIndexByHeader(headerName, sheet) {
    sheet = sheet || SpreadsheetApp.getActiveSheet();
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headerRow.indexOf(headerName) + 1;

    if (columnIndex === 0) {
      Logger.log(`Header "${headerName}" not found in the sheet.`);
      return null;
    }

    return columnIndex;
  }

  /**
   * Converts a column index to a letter.
   * @param {number} columnIndex The index of the column.
   * @return {string} The letter representation of the column.
   */
  function getColumnLetter(columnIndex) {
    let temp, letter = '';
    while (columnIndex > 0) {
      temp = (columnIndex - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      columnIndex = (columnIndex - temp - 1) / 26;
    }
    return letter;
  }

  /**
   * Gets the last row of data in the active sheet.
   * @return {number} The last row number with data.
   * note: need the usabilety of this function or use the default sheet.getLastRow()
   */
  function getLastRow() {
    const sheet = getOrCreateMainSheet();
    return Math.max(1, sheet.getLastRow());
  }

  /**
   * Generates a unique ID.
   * @return {string} A unique 6-digit ID.
   */
  function generateUniqueId() {
    return Math.floor(100000 + Math.random() * 900000).toString();
  }