// Constants
const HEADERS = ["id", "thumbnail", "price", "name", "description", "image_url", "retailer_id", "brand", "variant_group_id", "url", "currency", "category_id", "availability", "condition", "sale_price", "is_hidden"];
const CUSTOM_COLUMNS = ["thumbnail"];
const PRODUCT_TYPE_COLUMNS = {
  'standard': ['id', 'thumbnail', 'description', 'name', 'price', 'currency', 'image_url', 'availability', 'condition', 'brand', 'category_id', 'url', 'retailer_id'],
  'service': ['id', 'thumbnail', 'description', 'name', 'price', 'currency', 'image_url', 'availability', 'category_id', 'url'],
  'default': HEADERS
};
const CURRENCY_LIST = ["ILS", "AED", "USD", "CAD", "EUR", "GBP", "INR", "MXN", "BRL", "IDR", "ZAR"];
const CATEGORY_LIST = [
  "AUTO_VEHICLES_PARTS_ACCESSORIES",
  "BEAUTY_HEALTH_HAIR",
  "BUSINESS_SERVICES",
  "BABY_KIDS_GOODS",
  "COMMERCIAL_EQUIPMENT",
  "ELECTRONICS",
  "FOOD_BEVERAGES",
  "FURNITURE_APPLIANCES",
  "HOME_GOODS_DECOR",
  "LUGGAGE_BAGS",
  "MEDIA_MUSIC_BOOKS",
  "MISC",
  "PERSONAL_ACCESSORIES",
  "PET_SUPPLIES",
  "SPORTING_GOODS",
  "TOYS_GAMES_COLLECTIBLES",
  "APPAREL_ACCESSORIES",
  "FOOTWEAR",
  "HAIR_EXTENSIONS_WIGS",
  "HAIR_STYLING_TOOLS",
  "MAKEUP_COSMETICS",
  "FRAGRANCES",
  "SKIN_CARE",
  "BATH_BODY",
  "NAIL_CARE",
  "VITAMINS_SUPPLEMENTS",
  "MEDICAL_SUPPLIES_EQUIPMENT",
  "TICKETS",
  "TRAVEL_SERVICES"
];
const PRODUCT_TYPE_LIST = ["standard", "service", "default"];
const AVAILABILITY_LIST = ["in stock", "out of stock"];
const CONDITION_LIST = ["new", "used"];

var CONFIG = {
  DEVELOPER_KEY: PropertiesService.getScriptProperties().getProperty('DEVELOPER_KEY'),
  CLIENT_ID: PropertiesService.getScriptProperties().getProperty('CLIENT_ID')
};

/**
 * Retrieves configuration dropdown lists and preselected values.
 * @return {Object} An object containing dropdown lists and preselected values.
 */
function getConfigurationDropdownLists() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const configData = {
      currencyList: CURRENCY_LIST,
      categoryList: CATEGORY_LIST,
      productTypeList: PRODUCT_TYPE_LIST,
      availabilityList: AVAILABILITY_LIST,
      conditionList: CONDITION_LIST,
      preselectedProductType: scriptProperties.getProperty('DEFAULT_PRODUCT_TYPE') || "",
      preselectedCurrency: scriptProperties.getProperty('DEFAULT_CURRENCY') || "",
      preselectedCategory: scriptProperties.getProperty('DEFAULT_CATEGORY') || "",
      preselectedAvailability: scriptProperties.getProperty('DEFAULT_AVAILABILITY') || "",
      preselectedCondition: scriptProperties.getProperty('DEFAULT_CONDITION') || ""
    };

    logEvent('Configuration dropdown lists retrieved', 'INFO');
    return configData;
  } catch (error) {
    logEvent('Error retrieving configuration dropdown lists: ' + error.message, 'ERROR');
    throw error;
  }
}

/**
 * Displays the add-on homepage.
 * @param {Object} e The event object.
 * @return {Card} The card to display.
 */
function onHomepage(e) {
  var card = CardService.newCardBuilder();

  var mainSection = CardService.newCardSection()
    .setHeader("WhatsApp Catalog Tools");
  
  mainSection.addWidget(CardService.newTextButton()
    .setText("Setup Spreadsheet")
    .setOnClickAction(CardService.newAction().setFunctionName("setupSpreadsheet")));

  mainSection.addWidget(CardService.newTextButton()
    .setText("Configuration")
    .setOnClickAction(CardService.newAction().setFunctionName("showConfigurationPrompt")));

  mainSection.addWidget(CardService.newTextButton()
    .setText("Validate All Data")
    .setOnClickAction(CardService.newAction().setFunctionName("validateAllData")));

  // Add other buttons similarly
  
  card.addSection(mainSection);

  return card.build();
}

/**
 * Runs when the add-on is installed.
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when the document is opened, creating the add-on's menu.
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  createMenu();
  setupSpreadsheet();
}

/**
 * Gets the developer key from script properties.
 * @return {string} The developer key.
 */
function getDeveloperKey() {
  return PropertiesService.getScriptProperties().getProperty('DEVELOPER_KEY');
}

/**
 * Sets the developer key in script properties.
 */
function setDeveloperKey() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'CLIENT_ID': '999543760535-9tlmf6n70b7p7jjr7o11itvgrqkb3blt.apps.googleusercontent.com',
    'DEVELOPER_KEY': 'AIzaSyBoPL7WmQF-YYoJtfeKQOTIuPwBR0oR8zQ',
  });
}

/**
 * Gets the ID of the specific spreadsheet to access.
 * @return {string} The spreadsheet ID.
 */
function getAccessSpecificSheet() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

/**
 * Pauses execution for a specified number of milliseconds.
 * @param {number} milliseconds The number of milliseconds to sleep.
 * @return {Promise} A promise that resolves after the specified time.
 */
function sleep(milliseconds) {
  return new Promise(resolve => Utilities.sleep(milliseconds));
}

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
  const sheet = SpreadsheetApp.getActiveSheet();
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
  const sheet = SpreadsheetApp.getActiveSheet();
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

    logEvent('Configuration updated successfully', 'INFO');
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Configuration updated successfully"))
      .build();
  } catch (error) {
    logEvent('Error processing form: ' + error.message, 'ERROR');
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error updating configuration: " + error.message))
      .build();
  }
}

/**
 * Handles edits to the spreadsheet.
 * @param {Object} e The event object.
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (range.getRow() === 1) {
    range.setValue(e.oldValue);
    SpreadsheetApp.getUi().alert("You cannot edit the header row.");
    return;
  }

  const imageUrlColumnIndex = getColumnIndexByHeader('image_url', sheet);
  if (range.getColumn() === imageUrlColumnIndex) {
    Logger.log("Image URL column edited");
    
    const imageUrl = range.getValue();
    if (imageUrl) {
      Logger.log("Image URL not empty");
      
      generateAndSetUniqueId(sheet, range.getRow());

      const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail', sheet);
      if (thumbnailColumnIndex) {
        const thumbnailCell = sheet.getRange(range.getRow(), thumbnailColumnIndex);
        thumbnailCell.setFormula(`=IMAGE("${imageUrl}",4,100,100)`);
      }
    } else {
      const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail', sheet);
      if (thumbnailColumnIndex) {
        const thumbnailCell = sheet.getRange(range.getRow(), thumbnailColumnIndex);
        thumbnailCell.clearContent();
      }
    }
  }
  
  const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail', sheet);
  if (range.getColumn() === thumbnailColumnIndex) {
    range.setValue(e.oldValue);
    SpreadsheetApp.getUi().alert("The thumbnail column is automatically generated and cannot be edited directly.");
  }
  
  if (range.getRow() === sheet.getLastRow() && sheet.getLastRow() > sheet.getMaxRows() - 1) {
    extendDataValidation();
  }

  validateRow(range.getRow());
}

/**
 * Exports relevant columns based on the selected product type.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The newly created export sheet.
 */
function exportRelevantColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const defaultProductType = PropertiesService.getScriptProperties().getProperty('DEFAULT_PRODUCT_TYPE') || 'standard';
  
  const relevantColumns = PRODUCT_TYPE_COLUMNS[defaultProductType] || PRODUCT_TYPE_COLUMNS['default'];
  
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const newSheetName = `Export_${defaultProductType}_${timestamp}`;
  const exportSheet = ss.insertSheet(newSheetName);
  
  const columnIndices = relevantColumns.map(header => getColumnIndexByHeader(header, sourceSheet)).filter(index => index !== null);
  
  const sourceData = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).getValues();
  const exportData = sourceData.map(row => columnIndices.map(index => row[index - 1]));
  
  if (exportData.length > 0) {
    exportSheet.getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);
  }
  
  exportData[0].forEach((_, index) => {
    exportSheet.autoResizeColumn(index + 1);
  });
  
  const headerRange = exportSheet.getRange(1, 1, 1, exportData[0].length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  
  Logger.log(`Exported relevant columns for product type '${defaultProductType}' to sheet '${newSheetName}'`);
  
  return exportSheet;
}

/**
 * Logs an event with a timestamp and severity level.
 * @param {string} message The message to log.
 * @param {string} severity The severity level of the log (default: 'INFO').
 */
function logEvent(message, severity = 'INFO') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Log');
  sheet.appendRow([new Date(), severity, message]);
  console.log(`${severity}: ${message}`);
}

/**
 * Handles errors by logging them and displaying an alert to the user.
 * @param {Error} error The error object.
 */
function handleError(error) {
  logEvent('Error: ' + error.toString(), 'ERROR');
  SpreadsheetApp.getUi().alert('An error occurred. Please check the Log sheet for details.');
}