/**
 * Sets data validation for a specific column.
 * @param {string} columnName The name of the column to set validation for.
 * @param {string[]} columnValueOptions The list of valid options for the column.
 */
function setFieldValidation(columnName, columnValueOptions, sheet) {
  sheet = sheet || getOrCreateMainSheet();
  const columnIndex = getColumnIndexByHeader(columnName);
  ErrorHandler.log(`Setting validation for column: ${columnName}, index: ${columnIndex}`, 'INFO');

  if (columnIndex > 0) {
    const lastRow = sheet.getLastRow();
    const columnRange = sheet.getRange(2, columnIndex, Math.max(lastRow - 1, 1), 1);
    const columnRule = SpreadsheetApp.newDataValidation().requireValueInList(columnValueOptions).build();
    columnRange.setDataValidation(columnRule);
    ErrorHandler.log(`Data validation set for column ${columnName} from row 2 to ${lastRow}`, 'INFO');
  } else {
    ErrorHandler.log(`The header '${columnName}' was not found. Data validation not set.`, 'WARNING');
  }
}

/**
 * Clears all data validations from the active sheet.
 */
function clearAllDataValidations() {
  const sheet = getOrCreateMainSheet();
  const range = sheet.getDataRange();
  range.setDataValidation(null);

  ErrorHandler.log(`Cleared all data validations from the entire sheet`, 'INFO');
}

/**
 * Applies data validation to all relevant columns.
 */
function applyDataValidationToAllColumns() {
  clearAllDataValidations();
  setFieldValidation("currency", CURRENCY_LIST);
  setFieldValidation("category_id", CATEGORY_LIST);
  setFieldValidation("availability", AVAILABILITY_LIST);
  setFieldValidation("condition", CONDITION_LIST);
  ErrorHandler.log("Applied data validation to all relevant columns", 'INFO');
}

/**
 * Generates and sets a unique ID for a specific row if not already present.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to update.
 * @param {number} row The row number to update.
 * @return {string|null} The generated or existing ID, or null if an error occurred.
 */
function generateAndSetUniqueId(sheet, row) {
  const idColumnIndex = getColumnIndexByHeader('id', sheet);
  if (!idColumnIndex) {
    ErrorHandler.log("ID column not found", 'WARNING');
    return null;
  }

  const idCell = sheet.getRange(row, idColumnIndex);

  if (idCell.getValue()) {
    return idCell.getValue();
  }

  const existingIds = sheet.getRange(2, idColumnIndex, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String);

  let uniqueId;
  do {
    uniqueId = generateUniqueId();
  } while (existingIds.includes(uniqueId));

  idCell.setValue(uniqueId);

  ErrorHandler.log(`Generated and set unique ID ${uniqueId} for row ${row}`, 'INFO');
  return uniqueId;
}

/**
 * Initializes the thumbnail column in the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] The sheet to initialize. If not provided, uses the active sheet.
 */
function thumbnailColumnInit(sheet) {
  sheet = sheet || getOrCreateMainSheet();
  setupThumbnailColumn("thumbnail", "image_url", 2, 100, sheet);
}

/**
 * Generates and sets unique IDs for all rows with image URLs.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] The sheet to update. If not provided, uses the active sheet.
 */
function generateAndSetUniqueIds(sheet) {
  sheet = sheet || getOrCreateMainSheet();
  const idColumnIndex = getColumnIndexByHeader('id', sheet);
  const imageUrlColumnIndex = getColumnIndexByHeader('image_url', sheet);

  if (!idColumnIndex || !imageUrlColumnIndex) {
    ErrorHandler.log("Could not find 'id' or 'image_url' columns", 'WARNING');
    return;
  }

  const lastRow = sheet.getLastRow();
  const idRange = sheet.getRange(2, idColumnIndex, lastRow - 1, 1);
  const imageUrlRange = sheet.getRange(2, imageUrlColumnIndex, lastRow - 1, 1);

  const ids = idRange.getValues();
  const imageUrls = imageUrlRange.getValues();

  let changed = false;

  for (let i = 0; i < ids.length; i++) {
    if (imageUrls[i][0] && !ids[i][0]) {
      ids[i][0] = generateUniqueId();
      changed = true;
    }
  }

  if (changed) {
    idRange.setValues(ids);
    ErrorHandler.log(`Generated and set unique IDs for ${ids.filter(id => id[0]).length} rows`, 'INFO');
  } else {
    ErrorHandler.log("No new unique IDs needed to be generated", 'INFO');
  }
}

/**
 * Sets up the thumbnail column with image formulas.
 * @param {string} thumbnailColumnName The name of the thumbnail column.
 * @param {string} imageUrlColumnName The name of the image URL column.
 * @param {number} startRow The starting row for applying formulas.
 * @param {number} imageSize The size of the thumbnail images.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to update.
 */
function setupThumbnailColumn(thumbnailColumnName, imageUrlColumnName, startRow = 2, imageSize = 100, sheet) {
  try {
    const thumbnailColumnIndex = getColumnIndexByHeader(thumbnailColumnName, sheet);
    const imageUrlColumnIndex = getColumnIndexByHeader(imageUrlColumnName, sheet);

    ErrorHandler.log(`Thumbnail column index: ${thumbnailColumnIndex}, Image URL column index: ${imageUrlColumnIndex}`, 'INFO');

    if (!thumbnailColumnIndex || !imageUrlColumnIndex) {
      throw new Error("Could not find specified columns");
    }

    const columnDifference = imageUrlColumnIndex - thumbnailColumnIndex;

    const thumbnailFormula = `=IFERROR(IMAGE(SUBSTITUTE(SUBSTITUTE(INDIRECT("R[0]C[${columnDifference}]", FALSE), "/view?usp=drivesdk", ""), "https://drive.google.com/file/d/", "https://drive.google.com/uc?export=view&id="),4,${imageSize},${imageSize}), "Unable to load image")`;

    ErrorHandler.log(`Thumbnail formula: ${thumbnailFormula}`, 'INFO');

    const lastRow = sheet.getLastRow();
    const numRows = Math.max(1, lastRow - startRow + 1);
    const thumbnailRange = sheet.getRange(startRow, thumbnailColumnIndex, numRows, 1);

    ErrorHandler.log(`Setting formula for ${numRows} rows`, 'INFO');

    thumbnailRange.setFormula(thumbnailFormula);

    sheet.setRowHeights(startRow, numRows, imageSize + 10);
    sheet.setColumnWidth(thumbnailColumnIndex, imageSize + 10);

    SpreadsheetApp.flush();

    ErrorHandler.log(`Thumbnail column "${thumbnailColumnName}" set up successfully.`, 'INFO');
  } catch (error) {
    ErrorHandler.handleError(error, `Error in setupThumbnailColumn: ${error.message}`);
  }
}

/**
 * Clears all content and formatting from the active sheet.
 */
function clearSheetCompletely(sheet) {
  sheet = sheet || getOrCreateMainSheet();

  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getDataRange().clearDataValidations();

  ErrorHandler.log("Sheet cleared completely", 'INFO');
}

/**
 * Sets a value for all cells in a specific column.
 * @param {string} columnName The name of the column to update.
 * @param {string} value The value to set in the column.
 */
function setValuesToColumn(columnName, value) {
  const sheet = getOrCreateMainSheet();
  const columnIndex = getColumnIndexByHeader(columnName);

  if (columnIndex > 0) {
    const lastRow = getLastRow();
    const columnRange = sheet.getRange(2, columnIndex, lastRow - 1, 1);
    columnRange.setValue(value);
    ErrorHandler.log(`Set value '${value}' to column '${columnName}'`, 'INFO');
  } else {
    ErrorHandler.log(`Column '${columnName}' not found`, 'WARNING');
  }
}

/**
 * Sets up the spreadsheet with headers, data validation, and formatting.
 */
function setupSpreadsheet() {
  try {
    ErrorHandler.log('Starting setupSpreadsheet function', 'INFO');

    const sheet = getOrCreateMainSheet();
    ErrorHandler.log(`Using sheet: ${sheet.getName()}`, 'INFO');

    clearSheetCompletely(sheet);

    // Set headers
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    ErrorHandler.log('Headers set successfully', 'INFO');

    applyDataValidationToAllColumns(sheet);
    ErrorHandler.log('Data validation applied to all columns', 'INFO');

    // Apply conditional formatting for required fields
    const requiredHeaders = ["id", "name", "price", "currency", "image_url"];
    let rules = sheet.getConditionalFormatRules();
    requiredHeaders.forEach(header => {
      const columnIndex = getColumnIndexByHeader(header, sheet);
      if (columnIndex) {
        const range = sheet.getRange(2, columnIndex, Math.max(1, sheet.getLastRow() - 1), 1);
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=ISBLANK(INDIRECT("R[0]C[0]", FALSE))`)
          .setBackground("#FFB3BA")
          .setRanges([range])
          .build();
        rules.push(rule);
      }
    });
    sheet.setConditionalFormatRules(rules);

    // Set number format for price columns
    ["price", "sale_price"].forEach(header => {
      const columnIndex = getColumnIndexByHeader(header, sheet);
      if (columnIndex) {
        sheet.getRange(2, columnIndex, Math.max(1, sheet.getLastRow() - 1), 1).setNumberFormat("#,##0.00");
      }
    });

    setupThumbnailColumn("thumbnail", "image_url", 2, 150, sheet);

    sheet.setFrozenRows(1);
    const thumbnailColumnIndex = getColumnIndexByHeader("thumbnail", sheet);
    if (thumbnailColumnIndex) {
      sheet.setColumnWidth(thumbnailColumnIndex, 120);
    }

    // Protect header row
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const protection = headerRange.protect().setDescription("Header Row");
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    // Set warning for thumbnail column
    if (thumbnailColumnIndex) {
      const thumbnailRange = sheet.getRange(2, thumbnailColumnIndex, Math.max(1, sheet.getLastRow() - 1), 1);
      const thumbnailProtection = thumbnailRange.protect().setDescription("Thumbnail Column");
      thumbnailProtection.setWarningOnly(true);
    }

    sheet.autoResizeColumns(1, sheet.getLastColumn());
    hideIrrelevantColumns(sheet);

    ErrorHandler.log("Spreadsheet setup completed successfully", 'INFO');
  } catch (error) {
    ErrorHandler.handleError(error, `Error in setupSpreadsheet: ${error.message}`);
    throw error;
  }
}