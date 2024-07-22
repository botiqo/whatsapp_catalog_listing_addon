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

  card.addSection(mainSection);

  return card.build();
}

/**
 * Runs when the add-on is installed.
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  try {
    onOpen(e);
    logEvent('Add-on installed successfully', 'INFO');
  } catch (error) {
    logEvent('Error during add-on installation: ' + error.message, 'ERROR');
  }
}

/**
 * Runs when the document is opened, creating the add-on's menu.
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  try {
    createMenu();
    setupSpreadsheet();
    logEvent('Add-on opened successfully', 'INFO');
  } catch (error) {
    logEvent('Error during add-on opening: ' + error.message, 'ERROR');
  }
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
 * Validates product data.
 * @param {Object} product The product object to validate.
 * @return {Array} An array of error messages, empty if no errors.
 */
function validateProductData(product) {
  const errors = [];
  
  if (!product.id || product.id.length > 100) {
    errors.push("Invalid product ID");
  }
  if (!product.name || product.name.length > 150) {
    errors.push("Invalid product name");
  }
  if (product.description && product.description.length > 7000) {
    errors.push("Description exceeds 7000 characters");
  }
  if (!validatePrice(product.price, product.product_type)) {
    errors.push("Invalid price");
  }
  if (!CURRENCY_LIST.includes(product.currency)) {
    errors.push("Invalid currency");
  }
  if (!product.image_url || !/^https?:\/\/.+/.test(product.image_url)) {
    errors.push("Invalid image URL");
  }
  if (!AVAILABILITY_LIST.includes(product.availability)) {
    errors.push("Invalid availability");
  }
  if ((product.product_type === "standard" || product.product_type === "variable") && !CONDITION_LIST.includes(product.condition)) {
    errors.push("Invalid condition");
  }
  if (product.brand && product.brand.length > 64) {
    errors.push("Brand name exceeds 64 characters");
  }
  if (product.category_id && product.category_id.length > 250) {
    errors.push("Category exceeds 250 characters");
  }
  if (product.url && product.url.length > 2000) {
    errors.push("URL exceeds 2000 characters");
  }

  logEvent(`Product validation completed. Errors found: ${errors.length}`, 'INFO');
  return errors;
}

/**
 * Validates the price based on the product type.
 * @param {string} price The price to validate.
 * @param {string} productType The type of the product.
 * @return {boolean} True if the price is valid, false otherwise.
 */
function validatePrice(price, productType) {
  if (productType === "variable") {
    const priceRange = price.split("-");
    if (priceRange.length === 2) {
      return !isNaN(priceRange[0]) && !isNaN(priceRange[1]) && Number(priceRange[0]) > 0 && Number(priceRange[1]) > 0;
    }
  }
  return !isNaN(price) && Number(price) > 0;
}

/**
 * Validates a standard product.
 * @param {Object} product The product object to validate.
 * @return {Array} An array of error messages, empty if no errors.
 */
function validateStandardProduct(product) {
  const errors = validateProductData(product);
  if (product.product_type !== "standard") {
    errors.push("Invalid product type for standard product");
  }
  logEvent(`Standard product validation completed. Errors found: ${errors.length}`, 'INFO');
  return errors;
}

/**
 * Validates a service listing.
 * @param {Object} product The product object to validate.
 * @return {Array} An array of error messages, empty if no errors.
 */
function validateServiceListing(product) {
  const errors = validateProductData(product);
  if (product.product_type !== "service") {
    errors.push("Invalid product type for service listing");
  }
  if (product.condition) {
    errors.push("Condition should not be specified for services");
  }
  logEvent(`Service listing validation completed. Errors found: ${errors.length}`, 'INFO');
  return errors;
}

/**
 * Validates a variable product.
 * @param {Object} product The product object to validate.
 * @return {Array} An array of error messages, empty if no errors.
 */
function validateVariableProduct(product) {
  const errors = validateProductData(product);
  if (product.product_type !== "variable") {
    errors.push("Invalid product type for variable product");
  }
  if (!product.variant_group_id || product.variant_group_id.length > 100) {
    errors.push("Invalid variant group ID");
  }
  logEvent(`Variable product validation completed. Errors found: ${errors.length}`, 'INFO');
  return errors;
}

/**
 * Extends data validation to newly added rows.
 */
function extendDataValidation() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    applyDataValidationToAllColumns();
    
    logEvent(`Data validation extended to row ${lastRow}`, 'INFO');
  } catch (error) {
    logEvent(`Error extending data validation: ${error.message}`, 'ERROR');
    throw error;
  }
}

/**
 * Validates all products in the active sheet.
 * @return {Array} An array of error messages, empty if no errors.
 */
function validateAllProducts() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const errors = [];

  for (let i = 1; i < values.length; i++) {
    const product = {};
    headers.forEach((header, index) => {
      product[header] = values[i][index];
    });

    let productErrors;
    switch (product.product_type) {
      case 'standard':
        productErrors = validateStandardProduct(product);
        break;
      case 'service':
        productErrors = validateServiceListing(product);
        break;
      case 'variable':
        productErrors = validateVariableProduct(product);
        break;
      default:
        productErrors = [`Invalid product type: ${product.product_type}`];
    }

    if (productErrors.length > 0) {
      errors.push(`Row ${i + 1}: ${productErrors.join(', ')}`);
    }
  }

  if (errors.length > 0) {
    logEvent(`Validation completed. ${errors.length} errors found.`, 'WARNING');
    SpreadsheetApp.getUi().alert(`Validation Errors:\n\n${errors.join('\n')}`);
  } else {
    logEvent('Validation completed. No errors found.', 'INFO');
    SpreadsheetApp.getUi().alert('All products are valid!');
  }

  return errors;
}

/**
 * Validates a single row in the active sheet.
 * @param {number} rowNum The row number to validate.
 * @return {Array} An array of error messages, empty if no errors.
 */
function validateRow(rowNum) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const product = {};
    headers.forEach((header, index) => {
      product[header] = rowData[index];
    });

    let errors;
    switch (product.product_type) {
      case 'standard':
        errors = validateStandardProduct(product);
        break;
      case 'service':
        errors = validateServiceListing(product);
        break;
      case 'variable':
        errors = validateVariableProduct(product);
        break;
      default:
        errors = [`Invalid product type: ${product.product_type}`];
    }

    if (errors.length > 0) {
      logEvent(`Validation errors in row ${rowNum}: ${errors.join(', ')}`, 'WARNING');
    } else {
      logEvent(`Row ${rowNum} validated successfully`, 'INFO');
    }

    return errors;
  } catch (error) {
    logEvent(`Error validating row ${rowNum}: ${error.message}`, 'ERROR');
    throw error;
  }
}