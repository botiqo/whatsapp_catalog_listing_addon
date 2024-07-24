/**
 * Centralized error handling and logging system.
 */
const ErrorHandler = (function() {
  const LOG_SHEET_NAME = 'Error_Log';

  /**
   * Log levels enum
   */
  const LogLevel = {
    DEBUG: 'DEBUG',
    INFO: 'INFO',
    WARN: 'WARN',
    ERROR: 'ERROR'
  };

  /**
   * Gets or creates the log sheet.
   * @return {GoogleAppsScript.Spreadsheet.Sheet} The log sheet.
   */
  function getLogSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Level', 'Message', 'Function', 'File', 'Stack']);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  /**
   * Logs a message to the error log sheet and console.
   * @param {string} message - The message to log.
   * @param {LogLevel} level - The log level.
   * @param {Error} [error] - The error object, if applicable.
   */
  function log(message, level = LogLevel.INFO, error = null) {
    const sheet = getLogSheet();
    const timestamp = new Date().toISOString();
    const functionName = getFunctionName();
    const fileName = getFileName();
    const stack = error ? error.stack : '';

    sheet.appendRow([timestamp, level, message, functionName, fileName, stack]);

    // Log to console as well
    console.log(`[${level}] ${message}`);
    if (error) {
      console.error(error);
    }
  }

  /**
   * Gets the name of the function that called the error handler.
   * @return {string} The function name.
   */
  function getFunctionName() {
    try {
      throw new Error();
    } catch (e) {
      const stack = e.stack.split('\n');
      // The function name is typically on the third line of the stack trace
      const functionCallLine = stack[3];
      const functionName = functionCallLine.trim().split(' ')[1];
      return functionName || 'Unknown Function';
    }
  }

  /**
   * Gets the name of the file that called the error handler.
   * @return {string} The file name.
   */
  function getFileName() {
    try {
      throw new Error();
    } catch (e) {
      const stack = e.stack.split('\n');
      // The file name is typically on the third line of the stack trace
      const fileCallLine = stack[3];
      const fileName = fileCallLine.trim().split('/').pop().split(':')[0];
      return fileName || 'Unknown File';
    }
  }

  /**
   * Handles an error by logging it and optionally displaying a user-friendly message.
   * @param {Error} error - The error object.
   * @param {string} [userMessage] - A user-friendly message to display.
   */
  function handleError(error, userMessage = 'An error occurred. Please try again or contact support.') {
    log(error.message, LogLevel.ERROR, error);

    // Display a user-friendly message
    SpreadsheetApp.getUi().alert(userMessage);
  }

  // Public API
  return {
    LogLevel: LogLevel,
    log: log,
    handleError: handleError
  };
})();

/**
 * Displays the add-on homepage.
 * @param {Object} e The event object.
 * @return {Card} The card to display.
 */
function onHomepage(e) {
  return createHomepageCard();
}

/**
 * Gets or creates the main data sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The main data sheet.
 */
function getOrCreateMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("WhatsApp Catalog");

  if (!sheet) {
    sheet = ss.insertSheet("WhatsApp Catalog");
    ErrorHandler.log("Created new 'WhatsApp Catalog' sheet", 'INFO');
  } else {
    ErrorHandler.log("Using existing 'WhatsApp Catalog' sheet", 'INFO');
  }

  return sheet;
}

/**
 * Runs when the add-on is installed.
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {}

/**
 * Runs when the document is opened, creating the add-on's menu.
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {}

/**
 * Gets the difference between all headers and product type headers.
 * @param {Array} allHeaders All possible headers.
 * @param {Array} productTypeHeaders Headers for a specific product type.
 * @return {Array} The headers that are not in the product type headers.
 */
function getHeaderDiff(allHeaders, productTypeHeaders) {
  return allHeaders.filter(header => !productTypeHeaders.includes(header));
}

/**
 * Gets the default values for various fields.
 * @return {Object} An object containing default values.
 */
function getDefaultValues() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    product_type: scriptProperties.getProperty('DEFAULT_PRODUCT_TYPE') || 'default',
    currency: scriptProperties.getProperty('DEFAULT_CURRENCY') || 'USD',
    category: scriptProperties.getProperty('DEFAULT_CATEGORY') || '',
    availability: scriptProperties.getProperty('DEFAULT_AVAILABILITY') || 'in stock',
    condition: scriptProperties.getProperty('DEFAULT_CONDITION') || 'new'
  };
}

/**
 * Sets a default value for a column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to modify.
 * @param {number} columnIndex The index of the column to set.
 * @param {string} defaultValue The default value to set.
 */
function setColumnDefaultValue(sheet, columnIndex, defaultValue) {
  if (defaultValue) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, columnIndex, lastRow - 1, 1);
      const values = range.getValues();
      const updatedValues = values.map(row => [row[0] || defaultValue]);
      range.setValues(updatedValues);
    }
  }
}

/**
 * Sets default values for the selected product type.
 */
function setDefaultValuesForProductType() {
  const sheet = getOrCreateMainSheet();
  const defaultValues = getDefaultValues();
  const productType = defaultValues.product_type;

  const relevantColumns = PRODUCT_TYPE_COLUMNS[productType] || PRODUCT_TYPE_COLUMNS['default'];

  relevantColumns.forEach(columnName => {
    const columnIndex = getColumnIndexByHeader(columnName, sheet);
    if (columnIndex) {
      switch (columnName) {
        case 'currency':
          setColumnDefaultValue(sheet, columnIndex, defaultValues.currency);
          break;
        case 'category_id':
          setColumnDefaultValue(sheet, columnIndex, defaultValues.category);
          break;
        case 'availability':
          setColumnDefaultValue(sheet, columnIndex, defaultValues.availability);
          break;
        case 'condition':
          setColumnDefaultValue(sheet, columnIndex, defaultValues.condition);
          break;
      }
    }
  });

  hideIrrelevantColumns();

  Logger.log(`Default values set for product type: ${productType}`);
}

/**
 * Hides columns that are not relevant to the selected product type.
 */
function hideIrrelevantColumns() {
  const sheet = getOrCreateMainSheet();
  const selectedProductType = PropertiesService.getScriptProperties().getProperty('DEFAULT_PRODUCT_TYPE') || 'default';
  const diffHeaders = getHeaderDiff(HEADERS, PRODUCT_TYPE_COLUMNS[selectedProductType]);

  Logger.log("Hiding irrelevant columns: " + diffHeaders.join(","));
  const columnIndexArray = diffHeaders.map(headerName => getColumnIndexByHeader(headerName, sheet)).filter(index => index !== null);

  columnIndexArray.sort((a, b) => b - a);

  columnIndexArray.forEach(function(columnIndex) {
    if (columnIndex) {
      sheet.hideColumns(columnIndex);
    }
  });
}

/**
 * Processes the configuration form submission.
 * @param {Object} formObject The form data object.
 */
function processForm(formObject) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const {currency, availability, condition, category, product_type} = formObject;

    if (category) {
      scriptProperties.setProperty('DEFAULT_CATEGORY', category);
    }

    if (product_type) {
      scriptProperties.setProperty('DEFAULT_PRODUCT_TYPE', product_type);
      hideIrrelevantColumns();
    }

    if (currency) {
      scriptProperties.setProperty('DEFAULT_CURRENCY', currency);
      setValuesToColumn("currency", currency);
    }

    if (availability) {
      scriptProperties.setProperty('DEFAULT_AVAILABILITY', availability);
      setValuesToColumn("availability", availability);
    }

    if (condition) {
      scriptProperties.setProperty('DEFAULT_CONDITION', condition);
      setValuesToColumn("condition", condition);
    }

    setDefaultValuesForProductType();
    applyDataValidationToAllColumns();

    ErrorHandler.log('Configuration updated successfully', 'INFO');
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Configuration updated successfully"))
      .build();
  } catch (error) {
    ErrorHandler.handleError(error, "Error processing form");
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error updating configuration: " + error.message))
      .build();
  }
}

/**
 * Handles the edit event on the spreadsheet
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object
 */
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;

    if (range.getRow() === 1) {
      range.setValue(e.oldValue);
      SpreadsheetApp.getUi().alert("You cannot edit the header row.");
      return;
    }

    if (sheet.getName() !== "WhatsApp Catalog") {
      return;
    }

    const columnIndex = range.getColumn();
    const headerName = HEADERS[columnIndex - 1];

    if (headerName === 'image_url') {
      handleImageUrlEdit(sheet, range);
    } else if (headerName === 'thumbnail') {
      handleThumbnailEdit(range);
    }

    if (range.getRow() === sheet.getLastRow() && sheet.getLastRow() > sheet.getMaxRows() - 1) {
      applyDataValidationToAllColumns(sheet);
    }

    validateRow(range.getRow());
  } catch (error) {
    ErrorHandler.handleError(error, "Error in onEdit function");
  }
}

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
 * Saves the configuration from the card input.
 * @param {Object} e The event object from card interaction.
 * @return {CardService.ActionResponse} The action response after saving the configuration.
 */
function saveConfiguration(e) {
  const formInputs = e.commonEventObject.formInputs;

  const formObject = {
    product_type: formInputs.product_type.stringInputs.value[0],
    category: formInputs.category.stringInputs.value[0],
    currency: formInputs.currency.stringInputs.value[0],
    availability: formInputs.availability.stringInputs.value[0],
    condition: formInputs.condition.stringInputs.value[0]
  };

  return processForm(formObject);
}
