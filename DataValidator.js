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