/**
 * Shows the image picker card.
 * @return {CardService.Card} The image picker card.
 */
function showImagePickerCard() {
  const card = createBaseCard(CARD_TITLES.IMAGE_PICKER);

  const widgets = [
    CardService.newTextParagraph().setText("Select an image from your WhatsApp Images Folder:"),
    CardService.newTextButton()
      .setText("Choose Image")
      .setOnClickAction(CardService.newAction().setFunctionName("showImagePicker"))
  ];

  card.addSection(createSection("", widgets));
  return card.build();
}

/**
 * Creates and returns a Google Picker for image selection.
 * @return {CardService.ActionResponse} The action response with the picker.
 */
function showImagePicker() {
  const token = ScriptApp.getOAuthToken();
  const pickerCallback = "processImagePicker";
  const folderId = PropertiesService.getUserProperties().getProperty('WHATSAPP_FOLDER_ID');

  const picker = DocumentApp.newPickerBuilder()
    .addView(DocumentApp.PickerView.DOCS_IMAGES)
    .setOAuthToken(token)
    .setCallback(pickerCallback)
    .setOrigin(ScriptApp.getService().getUrl())
    .setTitle("Select an Image")
    .setSelectableMimeTypes("image/png,image/jpeg,image/gif")
    .setParent(folderId)
    .build();

  return CardService.newActionResponseBuilder()
    .setOpenDynamicLinkAction(picker)
    .build();
}

/**
 * Processes the selected image from the Google Picker.
 * @param {Object} params The parameters returned by the Picker.
 * @return {CardService.ActionResponse} The action response after image selection.
 */
function processImagePicker(params) {
  if (params.action === "picked") {
    const image = params.docs[0];
    const imageUrl = image.url;
    const imageName = image.name;

    updateImageUrl(imageUrl);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Image "${imageName}" selected successfully.`))
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
  const sheet = getOrCreateMainSheet();
  const cell = sheet.getActiveCell();
  const imageUrlColumnIndex = getColumnIndexByHeader('image_url');

  if (cell.getColumn() === imageUrlColumnIndex) {
    cell.setValue(url);

    // Update the thumbnail
    const thumbnailColumnIndex = getColumnIndexByHeader('thumbnail');
    if (thumbnailColumnIndex) {
      const thumbnailCell = sheet.getRange(cell.getRow(), thumbnailColumnIndex);
      thumbnailCell.setFormula(`=IMAGE("${url}",4,100,100)`);
    }

    // Generate and set unique ID if not already present
    generateAndSetUniqueId(sheet, cell.getRow());
  } else {
    throw new Error("Please select a cell in the image_url column before choosing an image.");
  }
}