/**
 * Shows the image picker dialog.
 */
function showPicker() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Picker')
      .setWidth(600)
      .setHeight(425)
      .setTitle('Select an Image');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select an Image');
  logEvent('Image picker dialog displayed', 'INFO');
}

/**
 * Gets the OAuth token for the current user.
 * @return {string|null} The OAuth token, or null if an error occurs.
 */
function getOAuthToken() {
  try {
    return ScriptApp.getOAuthToken();
  } catch (e) {
    logEvent(`Error getting OAuth token: ${e.toString()}`, 'ERROR');
    return null;
  }
}

/**
 * Updates the image URL in the active cell of the 'image_url' column.
 * @param {string} url The URL of the selected image.
 */
function updateImageUrl(url) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const imageUrlColumnIndex = getColumnIndexByHeader('image_url');

  if (cell.getColumn() === imageUrlColumnIndex) {
    cell.setValue(url);
    logEvent(`Updated image URL in row ${cell.getRow()} to: ${url}`, 'INFO');
    
    // Update the thumbnail
    const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail');
    if (thumbnailColumnIndex) {
      const thumbnailCell = sheet.getRange(cell.getRow(), thumbnailColumnIndex);
      thumbnailCell.setFormula(`=IMAGE("${url}",4,100,100)`);
      logEvent(`Updated thumbnail formula in row ${cell.getRow()}`, 'INFO');
    }

    // Generate and set unique ID if not already present
    generateAndSetUniqueId(sheet, cell.getRow());

  } else {
    SpreadsheetApp.getUi().alert("Please select a cell in the image_url column before choosing an image.");
    logEvent("Attempted to update image URL in wrong column", 'WARNING');
  }
}

/**
 * Retrieves the developer key from script properties.
 * @return {string} The developer key.
 */
function getDeveloperKey() {
  const key = PropertiesService.getScriptProperties().getProperty('DEVELOPER_KEY');
  if (!key) {
    logEvent("Developer key not found in script properties", 'WARNING');
  }
  return key;
}

/**
 * Initializes the image picker.
 * This function should be called from the client-side JavaScript.
 */
function initializePicker() {
  const token = getOAuthToken();
  const key = getDeveloperKey();
  const folderId = getWhatsAppFolderId();

  if (!token || !key || !folderId) {
    logEvent("Failed to initialize picker: missing token, key, or folder ID", 'ERROR');
    return null;
  }

  return {
    token: token,
    key: key,
    folderId: folderId
  };
}