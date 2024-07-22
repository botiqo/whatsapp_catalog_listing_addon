/**
 * Creates and returns a card for the image picker.
 * @return {Card} The image picker card.
 */
function createImagePickerCard() {
  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader().setTitle("Select an Image"));
  
  const section = CardService.newCardSection()
    .setHeader("Choose an image from your WhatsApp Catalog Listing folder:");
  
  const action = CardService.newAction().setFunctionName("selectImage");
  const button = CardService.newTextButton()
    .setText("Select Image")
    .setOnClickAction(action);
  
  section.addWidget(button);
  
  card.addSection(section);
  
  return card.build();
}

/**
 * Handles the image selection process.
 * @return {ActionResponse} The action response after image selection.
 */
function selectImage() {
  const folderId = getWhatsAppFolderId();
  
  if (!folderId) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("WhatsApp folder not found. Please set up the folder first."))
      .build();
  }
  
  const picker = createPicker(folderId);
  
  if (!picker) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Unable to create image picker. Please try again."))
      .build();
  }
  
  return CardService.newActionResponseBuilder()
    .setOpenDynamicLinkAction(picker)
    .build();
}

/**
 * Creates a Google Picker for image selection.
 * @param {string} folderId The ID of the WhatsApp folder.
 * @return {DynamicLinkAction|null} The picker action or null if unable to create.
 */
function createPicker(folderId) {
  const token = ScriptApp.getOAuthToken();
  const pickerCallback = "pickerCallback";
  const pickerBuilder = DocumentApp.newPickerBuilder()
    .addView(DocumentApp.PickerView.FOLDERS)
    .setOrigin(ScriptApp.getService().getUrl())
    .setOAuthToken(token)
    .setCallback(pickerCallback)
    .setSelectableMimeTypes("image/png,image/jpeg,image/gif")
    .setTitle("Select an Image");
  
  if (folderId) {
    pickerBuilder.setParent(folderId);
  }
  
  return pickerBuilder.build();
}

/**
 * Callback function for the Google Picker.
 * @param {Object} params The parameters returned by the Picker.
 * @return {ActionResponse} The action response after image selection.
 */
function pickerCallback(params) {
  if (params.action === "picked") {
    const doc = params.docs[0];
    const url = doc.url;
    const name = doc.name;
    
    updateImageUrl(url);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Image "${name}" selected successfully.`))
      .build();
  } else if (params.action === "cancel") {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Image selection cancelled."))
      .build();
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
    logEvent("Attempted to update image URL in wrong column", 'WARNING');
    throw new Error("Please select a cell in the image_url column before choosing an image.");
  }
}

/**
 * Shows the image picker card.
 * @return {CardService.Card} The image picker card.
 */
function showImagePicker() {
  return createImagePickerCard();
}