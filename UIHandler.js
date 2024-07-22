/**
 * Applies an advanced filter to the active sheet based on the given configuration
 * @param {Object} filterConfig The filter configuration
 * @param {number} filterConfig.minPrice The minimum price
 * @param {number} filterConfig.maxPrice The maximum price
 * @param {string[]} filterConfig.categories The categories to include
 * @param {boolean} filterConfig.showHidden Whether to show hidden items
 * @returns {number} The number of filtered results
 */
function applyAdvancedFilter(filterConfig) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const [headers, ...data] = sheet.getDataRange().getValues();

  const columnIndices = {
    price: headers.indexOf('price'),
    category: headers.indexOf('category_id'),
    availability: headers.indexOf('is_hidden')
  };

  const filteredData = data.filter(row => {
    const price = Number(row[columnIndices.price]);
    const category = row[columnIndices.category];
    const availability = row[columnIndices.availability];

    return (
      price >= filterConfig.minPrice &&
      price <= filterConfig.maxPrice &&
      filterConfig.categories.includes(category) &&
      (filterConfig.showHidden || availability !== 'true')
    );
  });

  if (filteredData.length > 0) {
    const resultSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Filtered Results');
    resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    resultSheet.getRange(2, 1, filteredData.length, headers.length).setValues(filteredData);
  }

  logEvent(`Advanced filter applied: ${filteredData.length} results`);
  return filteredData.length;
}

/**
 * Retrieves the configuration dropdown lists and preselected values
 * @returns {Object} The configuration data
 */
function getConfigurationDropdownLists() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const defaultValues = {
    productType: scriptProperties.getProperty('DEFAULT_PRODUCT_TYPE') || "",
    currency: scriptProperties.getProperty('DEFAULT_CURRENCY') || "",
    category: scriptProperties.getProperty('DEFAULT_CATEGORY') || "",
    availability: scriptProperties.getProperty('DEFAULT_AVAILABILITY') || "",
    condition: scriptProperties.getProperty('DEFAULT_CONDITION') || ""
  };

  return {
    currencyList: CURRENCY_LIST,
    categoryList: CATEGORY_LIST,
    productTypeList: PRODUCT_TYPE_LIST,
    availabilityList: AVAILABILITY_LIST,
    conditionList: CONDITION_LIST,
    preselectedProductType: defaultValues.productType,
    preselectedCurrency: defaultValues.currency,
    preselectedCategory: defaultValues.category,
    preselectedAvailability: defaultValues.availability,
    preselectedCondition: defaultValues.condition
  };
}

/**
 * Processes the configuration form submission
 * @param {Object} formObject The form data object
 */
function processForm(formObject) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSheet();

  // Update script properties
  scriptProperties.setProperties({
    'DEFAULT_PRODUCT_TYPE': formObject.product_type,
    'DEFAULT_CURRENCY': formObject.currency,
    'DEFAULT_CATEGORY': formObject.category,
    'DEFAULT_AVAILABILITY': formObject.availability,
    'DEFAULT_CONDITION': formObject.condition
  });

  // Update sheet data
  updateColumnValues(sheet, 'currency', formObject.currency);
  updateColumnValues(sheet, 'category_id', formObject.category);
  updateColumnValues(sheet, 'availability', formObject.availability);
  updateColumnValues(sheet, 'condition', formObject.condition);

  setDefaultValuesForProductType(sheet);
  applyDataValidationToAllColumns(sheet);

  logEvent('Configuration updated successfully', 'INFO');
}

/**
 * Updates values in a specific column
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to update
 * @param {string} columnName The name of the column to update
 * @param {string} value The value to set
 */
function updateColumnValues(sheet, columnName, value) {
  const columnIndex = HEADERS.indexOf(columnName) + 1;
  if (columnIndex > 0) {
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const range = sheet.getRange(2, columnIndex, lastRow - 1, 1);
    range.setValue(value);
  }
}

/**
 * Sets default values for the selected product type
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to update
 */
function setDefaultValuesForProductType(sheet) {
  const productType = PropertiesService.getScriptProperties().getProperty('DEFAULT_PRODUCT_TYPE') || 'default';
  const relevantColumns = PRODUCT_TYPE_COLUMNS[productType] || PRODUCT_TYPE_COLUMNS['default'];

  relevantColumns.forEach(columnName => {
    const columnIndex = HEADERS.indexOf(columnName) + 1;
    if (columnIndex > 0) {
      let defaultValue = '';
      switch (columnName) {
        case 'currency':
          defaultValue = PropertiesService.getScriptProperties().getProperty('DEFAULT_CURRENCY') || '';
          break;
        case 'category_id':
          defaultValue = PropertiesService.getScriptProperties().getProperty('DEFAULT_CATEGORY') || '';
          break;
        case 'availability':
          defaultValue = PropertiesService.getScriptProperties().getProperty('DEFAULT_AVAILABILITY') || '';
          break;
        case 'condition':
          defaultValue = PropertiesService.getScriptProperties().getProperty('DEFAULT_CONDITION') || '';
          break;
      }
      if (defaultValue) {
        updateColumnValues(sheet, columnName, defaultValue);
      }
    }
  });

  hideIrrelevantColumns(sheet);

  logEvent(`Default values set for product type: ${productType}`);
}

/**
 * Handles the edit event on the spreadsheet
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object
 */
// function onEdit(e) {
//   const sheet = e.source.getActiveSheet();
//   const range = e.range;

//   if (range.getRow() === 1) {
//     range.setValue(e.oldValue);
//     SpreadsheetApp.getUi().alert("You cannot edit the header row.");
//     return;
//   }

//   const columnIndex = range.getColumn();
//   const headerName = HEADERS[columnIndex - 1];

//   if (headerName === 'image_url') {
//     handleImageUrlEdit(sheet, range);
//   } else if (headerName === 'thumbnail') {
//     handleThumbnailEdit(range);
//   }

//   if (range.getRow() === sheet.getLastRow() && sheet.getLastRow() > sheet.getMaxRows() - 1) {
//     applyDataValidationToAllColumns(sheet);
//   }

//   validateRow(range.getRow());
// }

/**
 * Handles edits to the image_url column
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet
 * @param {GoogleAppsScript.Spreadsheet.Range} range The edited range
 */
function handleImageUrlEdit(sheet, range) {
  const imageUrl = range.getValue();
  if (imageUrl) {
    generateAndSetUniqueId(sheet, range.getRow());
    updateThumbnail(sheet, range.getRow(), imageUrl);
  } else {
    clearThumbnail(sheet, range.getRow());
  }
}

/**
 * Handles edits to the thumbnail column
 * @param {GoogleAppsScript.Spreadsheet.Range} range The edited range
 */
function handleThumbnailEdit(range) {
  range.setValue(range.getOldValue());
  SpreadsheetApp.getUi().alert("The thumbnail column is automatically generated and cannot be edited directly.");
}

/**
 * Updates the thumbnail for a given row
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet
 * @param {number} row The row number to update
 * @param {string} imageUrl The image URL
 */
function updateThumbnail(sheet, row, imageUrl) {
  const thumbnailColumnIndex = HEADERS.indexOf('thumbnail') + 1;
  if (thumbnailColumnIndex > 0) {
    const thumbnailCell = sheet.getRange(row, thumbnailColumnIndex);
    thumbnailCell.setFormula(`=IMAGE("${imageUrl}",4,100,100)`);
  }
}

/**
 * Clears the thumbnail for a given row
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet
 * @param {number} row The row number to clear
 */
function clearThumbnail(sheet, row) {
  const thumbnailColumnIndex = HEADERS.indexOf('thumbnail') + 1;
  if (thumbnailColumnIndex > 0) {
    const thumbnailCell = sheet.getRange(row, thumbnailColumnIndex);
    thumbnailCell.clearContent();
  }
}

/**
 * Generates and sets a unique ID for a given row
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet
 * @param {number} row The row number to update
 * @returns {string|null} The generated ID or null if unsuccessful
 */
function generateAndSetUniqueId(sheet, row) {
  const idColumnIndex = HEADERS.indexOf('id') + 1;
  if (idColumnIndex === 0) {
    logEvent("ID column not found", 'WARNING');
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
    uniqueId = Math.floor(100000 + Math.random() * 900000).toString();
  } while (existingIds.includes(uniqueId));

  idCell.setValue(uniqueId);

  logEvent(`Generated and set unique ID ${uniqueId} for row ${row}`, 'INFO');
  return uniqueId;
}

// ... (other functions as needed)