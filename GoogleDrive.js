/**
 * Gets the OAuth token for the current user.
 * @return {string} The OAuth token.
 */
function getOAuthToken() {
  return ScriptApp.getOAuthToken();
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
      logEvent(`Existing folder "${folderName}" found.`, 'INFO');
      return folders.next();
    } else {
      logEvent(`Creating new folder "${folderName}".`, 'INFO');
      const newFolder = DriveApp.createFolder(folderName);

      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      makeWhatsAppFolderFilesPublic(newFolder);

      return newFolder;
    }
  } catch (error) {
    logEvent(`Error in getOrCreateWhatsAppFolder: ${error.message}`, 'ERROR');
    throw error;
  }
}

/**
 * Sets the WhatsApp folder name in user properties.
 * @param {string} folderName The name to set for the WhatsApp folder.
 */
function setWhatsAppFolderName(folderName) {
  if (typeof folderName !== 'string' || folderName.trim() === '') {
    logEvent('Invalid folder name provided.', 'ERROR');
    throw new Error('Invalid folder name provided.');
  }

  PropertiesService.getUserProperties().setProperty('WHATSAPP_FOLDER_NAME', folderName.trim());
  logEvent(`WhatsApp folder name set to "${folderName}".`, 'INFO');
}

/**
 * Gets the ID of the WhatsApp folder.
 * @return {string} The ID of the WhatsApp folder.
 */
function getWhatsAppFolderId() {
  try {
    const folder = getOrCreateWhatsAppFolder();
    const folderId = folder.getId();
    logEvent(`WhatsApp folder ID: ${folderId}`, 'INFO');
    return folderId;
  } catch (error) {
    logEvent(`Error in getWhatsAppFolderId: ${error.message}`, 'ERROR');
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

  logEvent(`Made ${fileCount} file(s) in the WhatsApp folder public.`, 'INFO');
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
  logEvent(`Starting getImageUrlsAndSetInSheet function with directory ID: ${directoryId}`, 'INFO');

  const imageUrls = [];

  try {
    const folder = DriveApp.getFolderById(directoryId);
    logEvent(`Successfully accessed folder: ${folder.getName()}`, 'INFO');

    const imageMimeTypes = [MimeType.JPEG, MimeType.PNG, MimeType.GIF];

    for (const mimeType of imageMimeTypes) {
      logEvent(`Searching for files of type: ${mimeType}`, 'INFO');
      const files = folder.getFilesByType(mimeType);

      while (files.hasNext()) {
        const file = files.next();
        const url = file.getUrl();
        imageUrls.push(url);
        logEvent(`Found image: ${file.getName()} (${url})`, 'INFO');
      }
    }

    logEvent(`Total images found: ${imageUrls.length}`, 'INFO');

    setImageUrlsInSheet(imageUrls);

  } catch (error) {
    logEvent(`Error in getImageUrlsAndSetInSheet: ${error.message}`, 'ERROR');
    throw error;
  }

  return imageUrls;
}

/**
 * Sets image URLs in the active sheet.
 * @param {string[]} imageUrls An array of image URLs to set in the sheet.
 */
function setImageUrlsInSheet(imageUrls) {
  logEvent("Starting to set image URLs in sheet", 'INFO');

  const sheet = SpreadsheetApp.getActiveSheet();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const imageUrlColIndex = headers.indexOf('image_url') + 1;

  if (imageUrlColIndex === 0) {
    logEvent("'image_url' column not found in the sheet", 'ERROR');
    throw new Error("'image_url' column not found in the sheet");
  }

  logEvent(`'image_url' column found at index: ${imageUrlColIndex}`, 'INFO');

  if (imageUrls.length > 0) {
    const range = sheet.getRange(2, imageUrlColIndex, imageUrls.length, 1);
    range.setValues(imageUrls.map(url => [url]));
    logEvent(`Set ${imageUrls.length} image URLs in the sheet`, 'INFO');
  } else {
    logEvent("No image URLs to set in the sheet", 'WARNING');
  }
}

/**
 * Checks and logs thumbnail formulas in the active sheet.
 */
function checkThumbnailFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail', sheet);

  if (!thumbnailColumnIndex) {
    logEvent("Could not find 'thumbnail' column", 'ERROR');
    return;
  }

  const lastRow = sheet.getLastRow();
  const formulas = sheet.getRange(2, thumbnailColumnIndex, lastRow - 1, 1).getFormulas();

  formulas.forEach((formula, index) => {
    if (formula[0]) {
      logEvent(`Row ${index + 2} formula: ${formula[0]}`, 'INFO');
    } else {
      logEvent(`Row ${index + 2}: No formula`, 'WARNING');
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
    logEvent("Could not find 'image_url' column", 'ERROR');
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
        logEvent(`Row ${index + 2}: URL ${url[0]} - Response code: ${responseCode}`, 'INFO');
      } catch (error) {
        logEvent(`Row ${index + 2}: Error accessing URL ${url[0]} - ${error.message}`, 'ERROR');
      }
    }
  });

  logEvent("All image URL tests completed", 'INFO');
}

/**
 * Checks and logs the content of the thumbnail column in the active sheet.
 */
function checkThumbnailContent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail', sheet);

  if (!thumbnailColumnIndex) {
    logEvent("Could not find 'thumbnail' column", 'ERROR');
    return;
  }

  const lastRow = sheet.getLastRow();
  const thumbnailRange = sheet.getRange(2, thumbnailColumnIndex, lastRow - 1, 1);
  const thumbnailValues = thumbnailRange.getValues();

  thumbnailValues.forEach((value, index) => {
    logEvent(`Row ${index + 2} thumbnail content: ${value[0]}`, 'INFO');
  });
}

/**
 * Gets image URLs from the WhatsApp folder and sets them in the active sheet.
 * @return {string[]} An array of image URLs.
 */
function getListingImageUrlsAndSetInSheet() {
  logEvent("Starting getListingImageUrlsAndSetInSheet function", 'INFO');

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const directoryId = getWhatsAppFolderId();

    const imageUrls = getImageUrlsAndSetInSheet(directoryId);

    if (imageUrls.length > 0) {
      logEvent("Retrieved and set image URLs:", 'INFO');
      imageUrls.forEach((url, index) => {
        logEvent(`${index + 1}: ${url}`, 'INFO');
      });

      thumbnailColumnInit(sheet);
      checkThumbnailFormulas();
      testImageUrls();
      checkThumbnailContent();
      generateAndSetUniqueIds(sheet);
      setDefaultValuesForProductType();

      SpreadsheetApp.flush();
      sheet.autoResizeColumn(1);
      logEvent("Thumbnail column initialized and sheet recalculated", 'INFO');
    } else {
      logEvent("No image URLs retrieved or set.", 'WARNING');
    }

    logEvent("Finished getListingImageUrlsAndSetInSheet function", 'INFO');
    return imageUrls;
  } catch (error) {
    logEvent(`Error in getListingImageUrlsAndSetInSheet: ${error.message}`, 'ERROR');
    throw error;
  }
}