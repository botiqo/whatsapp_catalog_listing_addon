/**
 * Gets the OAuth token for the current user.
 * @return {string} The OAuth token.
 */
function getOAuthToken() {
  try {
    return ScriptApp.getOAuthToken();
  } catch (error) {
    ErrorHandler.handleError(error, "Failed to get OAuth token");
    throw error;
  }
}

/**
 * Gets or creates the WhatsApp folder in Google Drive.
 * @return {GoogleAppsScript.Drive.Folder} The WhatsApp folder.
 */
function getOrCreateWhatsAppFolder() {
  const userProperties = PropertiesService.getUserProperties();
  const folderName = userProperties.getProperty('WHATSAPP_FOLDER_NAME') || "WhatsApp Catalog Listing";

  try {
    const folders = DriveApp.getFoldersByName(folderName);

    if (folders.hasNext()) {
      ErrorHandler.log(`Existing folder "${folderName}" found.`, 'INFO');
      return folders.next();
    } else {
      ErrorHandler.log(`Creating new folder "${folderName}".`, 'INFO');
      const newFolder = DriveApp.createFolder(folderName);

      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      makeWhatsAppFolderFilesPublic(newFolder);

      return newFolder;
    }
  } catch (error) {
    ErrorHandler.handleError(error, "Error accessing or creating WhatsApp folder");
    throw error;
  }
}

/**
 * Sets the WhatsApp folder name in user properties.
 * @param {string} folderName The name to set for the WhatsApp folder.
 */
function setWhatsAppFolderName(folderName) {
  if (typeof folderName !== 'string' || folderName.trim() === '') {
    ErrorHandler.log('Invalid folder name provided.', 'ERROR');
    throw new Error('Invalid folder name provided.');
  }

  PropertiesService.getUserProperties().setProperty('WHATSAPP_FOLDER_NAME', folderName.trim());
  ErrorHandler.log(`WhatsApp folder name set to "${folderName}".`, 'INFO');
}

/**
 * Gets the ID of the WhatsApp folder.
 * @return {string} The ID of the WhatsApp folder.
 */
function getWhatsAppFolderId() {
  try {
    const folder = getOrCreateWhatsAppFolder();
    const folderId = folder.getId();
    ErrorHandler.log(`WhatsApp folder ID: ${folderId}`, 'INFO');
    return folderId;
  } catch (error) {
    ErrorHandler.handleError(error, "Error Please try again or contact support.");
    throw error;
  }
}

/**
 * Makes all files in the given folder publicly accessible.
 * @param {GoogleAppsScript.Drive.Folder} folder The folder to process.
 */
function makeWhatsAppFolderFilesPublic(folder) {
  const files = folder.getFiles();
  let fileCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileCount++;
  }

  ErrorHandler.log(`Made ${fileCount} file(s) in the WhatsApp folder public.`, 'INFO');
}

/**
 * Creates a thumbnail URL for a Google Drive file.
 * @param {string} url The Google Drive file URL.
 * @param {number} size The desired size of the thumbnail.
 * @return {string} The thumbnail URL.
 */
function DRIVETHUMBNAIL(url, size) {
  if (!url) return "";

  const fileId = url.match(/[-\w]{25,}/);

  if (!fileId) return url;

  return `https://drive.google.com/thumbnail?id=${fileId[0]}&sz=w${size}`;
}

/**
 * Gets image URLs from a Google Drive folder and sets them in the active sheet.
 * @param {string} directoryId The ID of the Google Drive folder.
 * @return {string[]} An array of image URLs.
 */
function getImageUrlsAndSetInSheet(directoryId) {
  ErrorHandler.log(`Starting getImageUrlsAndSetInSheet function with directory ID: ${directoryId}`, 'INFO');

  const imageUrls = [];

  try {
    const folder = DriveApp.getFolderById(directoryId);
    ErrorHandler.log(`Successfully accessed folder: ${folder.getName()}`, 'INFO');

    const imageMimeTypes = [MimeType.JPEG, MimeType.PNG, MimeType.GIF];

    for (const mimeType of imageMimeTypes) {
      ErrorHandler.log(`Searching for files of type: ${mimeType}`, 'INFO');
      const files = folder.getFilesByType(mimeType);

      while (files.hasNext()) {
        const file = files.next();
        const url = file.getUrl();
        imageUrls.push(url);
        ErrorHandler.log(`Found image: ${file.getName()} (${url})`, 'INFO');
      }
    }

    ErrorHandler.log(`Total images found: ${imageUrls.length}`, 'INFO');

    setImageUrlsInSheet(imageUrls);

  } catch (error) {
    ErrorHandler.handleError(error, "Error Please try again or contact support.");
    throw error;
  }

  return imageUrls;
}

/**
 * Sets image URLs in the active sheet.
 * @param {string[]} imageUrls An array of image URLs to set in the sheet.
 */
function setImageUrlsInSheet(imageUrls) {
  ErrorHandler.log("Starting to set image URLs in sheet", 'INFO');

  const sheet = getOrCreateMainSheet();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const imageUrlColIndex = headers.indexOf('image_url') + 1;

  if (imageUrlColIndex === 0) {
    ErrorHandler.log("'image_url' column not found in the sheet", 'ERROR');
    throw new Error("'image_url' column not found in the sheet");
  }

  ErrorHandler.log(`'image_url' column found at index: ${imageUrlColIndex}`, 'INFO');

  if (imageUrls.length > 0) {
    const range = sheet.getRange(2, imageUrlColIndex, imageUrls.length, 1);
    range.setValues(imageUrls.map(url => [url]));
    ErrorHandler.log(`Set ${imageUrls.length} image URLs in the sheet`, 'INFO');
  } else {
    ErrorHandler.log("No image URLs to set in the sheet", 'WARNING');
  }
}

/**
 * Checks and logs thumbnail formulas in the active sheet.
 */
function checkThumbnailFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail', sheet);

  if (!thumbnailColumnIndex) {
    ErrorHandler.log("Could not find 'thumbnail' column", 'ERROR');
    return;
  }

  const lastRow = sheet.getLastRow();
  const formulas = sheet.getRange(2, thumbnailColumnIndex, lastRow - 1, 1).getFormulas();

  formulas.forEach((formula, index) => {
    if (formula[0]) {
      ErrorHandler.log(`Row ${index + 2} formula: ${formula[0]}`, 'INFO');
    } else {
      ErrorHandler.log(`Row ${index + 2}: No formula`, 'WARNING');
    }
  });
}

/**
 * Tests the accessibility of image URLs in the active sheet.
 */
function testImageUrls() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const imageUrlColumnIndex = getColumnIndexByHeader('image_url', sheet);

  if (!imageUrlColumnIndex) {
    ErrorHandler.log("Could not find 'image_url' column", 'ERROR');
    return;
  }

  const lastRow = sheet.getLastRow();
  const imageUrls = sheet.getRange(2, imageUrlColumnIndex, lastRow - 1, 1).getValues();

  imageUrls.forEach((url, index) => {
    if (url[0]) {
      Utilities.sleep(1000); // Wait 1 second between requests to avoid rate limiting
      try {
        const response = UrlFetchApp.fetch(url[0], {muteHttpExceptions: true});
        const responseCode = response.getResponseCode();
        ErrorHandler.log(`Row ${index + 2}: URL ${url[0]} - Response code: ${responseCode}`, 'INFO');
      } catch (error) {
        ErrorHandler.handleError(error, `Row ${index + 2}: Error accessing URL ${url[0]} - ${error.message}`);
      }
    }
  });

  ErrorHandler.log("All image URL tests completed", 'INFO');
}

/**
 * Checks and logs the content of the thumbnail column in the active sheet.
 */
function checkThumbnailContent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail', sheet);

  if (!thumbnailColumnIndex) {
    ErrorHandler.log("Could not find 'thumbnail' column", 'ERROR');
    return;
  }

  const lastRow = sheet.getLastRow();
  const thumbnailRange = sheet.getRange(2, thumbnailColumnIndex, lastRow - 1, 1);
  const thumbnailValues = thumbnailRange.getValues();

  thumbnailValues.forEach((value, index) => {
    ErrorHandler.log(`Row ${index + 2} thumbnail content: ${value[0]}`, 'INFO');
  });
}

/**
 * Gets image URLs from the WhatsApp folder and sets them in the active sheet.
 * @return {string[]} An array of image URLs.
 */
function getListingImageUrlsAndSetInSheet() {
  ErrorHandler.log("Starting getListingImageUrlsAndSetInSheet function", 'INFO');

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const directoryId = getWhatsAppFolderId();

    const imageUrls = getImageUrlsAndSetInSheet(directoryId);

    if (imageUrls.length > 0) {
      ErrorHandler.log("Retrieved and set image URLs:", 'INFO');
      imageUrls.forEach((url, index) => {
        ErrorHandler.log(`${index + 1}: ${url}`, 'INFO');
      });

      thumbnailColumnInit(sheet);
      checkThumbnailFormulas();
      testImageUrls();
      checkThumbnailContent();
      generateAndSetUniqueIds(sheet);
      setDefaultValuesForProductType();

      SpreadsheetApp.flush();
      sheet.autoResizeColumn(1);
      ErrorHandler.log("Thumbnail column initialized and sheet recalculated", 'INFO');
    } else {
      ErrorHandler.log("No image URLs retrieved or set.", 'WARNING');
    }

    ErrorHandler.log("Finished getListingImageUrlsAndSetInSheet function", 'INFO');
    return imageUrls;
  } catch (error) {
    ErrorHandler.handleError(error, "Error Please try again or contact support.");
    throw error;
  }
}

/**
 * Shows the folder picker dialog.
 * @return {CardService.Card} The card with folder selection options.
 */
function showFolderPicker() {
  var cardBuilder = CardService.newCardBuilder();
  var section = CardService.newCardSection().setHeader("Select WhatsApp Images Folder");

  // Fetch folders
  var folders = DriveApp.getFolders();
  var folderList = [];

  // Collect folder information
  while (folders.hasNext()) {
    var folder = folders.next();
    folderList.push({
      id: folder.getId(),
      name: folder.getName()
    });
  }

  // Sort folders alphabetically
  folderList.sort((a, b) => a.name.localeCompare(b.name));

  // Create a selection input for folders
  var listItemBuilder = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle("Choose a folder")
    .setFieldName("selectedFolderId");

  // Add folders to the selection input
  folderList.forEach(function(folder) {
    listItemBuilder.addItem(folder.name, folder.id, false);
  });

  section.addWidget(listItemBuilder);

  // Add a button to confirm selection
  var confirmButton = CardService.newTextButton()
    .setText("Confirm Selection")
    .setOnClickAction(CardService.newAction().setFunctionName("processFolderSelection"));
  section.addWidget(confirmButton);

  cardBuilder.addSection(section);
  return cardBuilder.build();
}

/**
 * Processes the selected folder.
 * @param {Object} e The event object from the card action.
 * @return {CardService.ActionResponse} The action response after processing the selection.
 */
function processFolderSelection(e) {
  var folderId = e.formInput.selectedFolderId;
  var folder = DriveApp.getFolderById(folderId);
  var folderName = folder.getName();

  PropertiesService.getUserProperties().setProperties({
    'WHATSAPP_FOLDER_ID': folderId,
    'WHATSAPP_FOLDER_NAME': folderName
  });

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(`Folder "${folderName}" selected as WhatsApp Images Folder.`))
    .setNavigation(CardService.newNavigation().pushCard(createImportImagesCard()))
    .build();
}