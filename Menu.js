/**
 * Creates the add-on menu in the Google Sheets UI.
 */
function createMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createAddonMenu();

    menu.addItem("Setup Spreadsheet", "showSetupSpreadsheetCard")
        .addItem("Configuration", "showConfigurationCard")
        .addItem("Validate All Data", "showValidateAllDataCard")
        .addItem("Select Image from Drive", "showImagePickerCard")
        .addItem("Import images from Drive", "showImportImagesCard")
        .addItem("Set WhatsApp Images Folder", "showSetFolderNameCard")
        .addItem("Export Relevant Columns", "showExportColumnsCard")
        .addItem("Instructions", "showInstructionsCard")
        .addToUi();

    logEvent('Menu created', 'INFO');
  } catch (error) {
    logEvent('Error creating menu: ' + error.message, 'ERROR');
    console.error('Error creating menu:', error);
  }
}

/**
 * Shows a card to set the WhatsApp images folder name.
 * @return {CardService.Card} The folder name input card.
 */
function showSetFolderNameCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Set WhatsApp Images Folder"));

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextInput()
      .setFieldName("folderName")
      .setTitle("Enter the name for the WhatsApp images folder:")
    )
    .addWidget(CardService.newTextButton()
      .setText("Set Folder Name")
      .setOnClickAction(CardService.newAction().setFunctionName("setFolderNameFromCard"))
    );

  card.addSection(section);
  return card.build();
}

/**
 * Sets the folder name from the card input.
 * @param {Object} e The event object from card interaction.
 * @return {CardService.ActionResponse} The action response after setting the folder name.
 */
function setFolderNameFromCard(e) {
  const folderName = e.commonEventObject.formInputs.folderName;

  try {
    setWhatsAppFolderName(folderName);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Folder name set to: " + folderName))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error: " + error.message))
      .build();
  }
}

/**
 * Shows the configuration card using CardService.
 * @return {CardService.Card} The configuration card.
 */
function showConfigurationCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("WhatsApp Catalog Configuration"));

  const section = CardService.newCardSection();

  // Get the configuration data
  const config = getConfigurationDropdownLists();

  // WhatsApp Catalog Listing Type
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("WhatsApp Catalog Listing Type")
    .setFieldName("product_type")
    .addItems(config.productTypeList.map(type => CardService.newOption(type, type, type === config.preselectedProductType)))
  );

  // Default Currency
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Currency")
    .setFieldName("currency")
    .addItems(config.currencyList.map(currency => CardService.newOption(currency, currency, currency === config.preselectedCurrency)))
  );

  // Default Category
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Category")
    .setFieldName("category")
    .addItems(config.categoryList.map(category => CardService.newOption(category, category, category === config.preselectedCategory)))
  );

  // Default Availability
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Availability")
    .setFieldName("availability")
    .addItems(config.availabilityList.map(availability => CardService.newOption(availability, availability, availability === config.preselectedAvailability)))
  );

  // Default Condition
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Condition")
    .setFieldName("condition")
    .addItems(config.conditionList.map(condition => CardService.newOption(condition, condition, condition === config.preselectedCondition)))
  );

  // Save button
  section.addWidget(CardService.newTextButton()
    .setText("Save Configuration")
    .setOnClickAction(CardService.newAction().setFunctionName("saveConfiguration"))
  );

  card.addSection(section);

  return card.build();
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

/**
 * Shows the instructions card.
 * @return {CardService.Card} The instructions card.
 */
function showInstructionsCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Instructions"));

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText("1. Click on 'Setup Spreadsheet' to initialize your sheet with the correct headers and formatting."))
    .addWidget(CardService.newTextParagraph().setText("2. Use 'Configuration' to set default values for product type, currency, category, availability, and condition."))
    .addWidget(CardService.newTextParagraph().setText("3. Use 'Set WhatsApp Images Folder' to specify the Google Drive folder for your catalog images."))
    .addWidget(CardService.newTextParagraph().setText("4. 'Import images from Drive' will populate your sheet with image URLs from the specified folder."))
    .addWidget(CardService.newTextParagraph().setText("5. Use 'Select Image from Drive' to choose individual images for each product."))
    .addWidget(CardService.newTextParagraph().setText("6. 'Validate All Data' checks your catalog for any errors or missing information."))
    .addWidget(CardService.newTextParagraph().setText("7. When your catalog is ready, use 'Export Relevant Columns' to create a new sheet with only the required data for WhatsApp."));

  card.addSection(section);
  return card.build();
}

/**
 * Shows the image picker card.
 * @return {CardService.ActionResponse} The action response to show the image picker.
 */
function showImagePickerCard() {
  const card = showImagePicker();
  const navigation = CardService.newNavigation().pushCard(card);
  return CardService.newActionResponseBuilder()
    .setNavigation(navigation)
    .build();
}

/**
 * Shows a card to validate all data in the spreadsheet.
 * @return {CardService.Card} The validate data card.
 */
function showValidateAllDataCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Validate All Data"));

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText("This will check all product data in your spreadsheet for errors or missing information."))
    .addWidget(CardService.newTextButton()
      .setText("Start Validation")
      .setOnClickAction(CardService.newAction().setFunctionName("validateAllDataFromCard"))
    );

  card.addSection(section);
  return card.build();
}

/**
 * Performs data validation from the card confirmation.
 * @return {CardService.ActionResponse} The action response after validating the data.
 */
function validateAllDataFromCard() {
  try {
    const errors = validateAllProducts();
    if (errors.length === 0) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("All data is valid!"))
        .build();
    } else {
      // Create a card to display errors
      const errorCard = CardService.newCardBuilder();
      errorCard.setHeader(CardService.newCardHeader().setTitle("Validation Errors"));

      const errorSection = CardService.newCardSection();
      errors.forEach(error => {
        errorSection.addWidget(CardService.newTextParagraph().setText(error));
      });

      errorCard.addSection(errorSection);

      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().pushCard(errorCard.build()))
        .build();
    }
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error validating data: " + error.message))
      .build();
  }
}

/**
 * Shows a card to import images from Google Drive.
 * @return {CardService.Card} The import images card.
 */
function showImportImagesCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Import Images from Drive"));

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText("This will import image URLs from your WhatsApp Catalog Listing folder in Google Drive."))
    .addWidget(CardService.newTextButton()
      .setText("Start Import")
      .setOnClickAction(CardService.newAction().setFunctionName("importImagesFromCard"))
    );

  card.addSection(section);
  return card.build();
}

/**
 * Performs image import from the card confirmation.
 * @return {CardService.ActionResponse} The action response after importing images.
 */
function importImagesFromCard() {
  try {
    const imageUrls = getListingImageUrlsAndSetInSheet();
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Successfully imported ${imageUrls.length} image URLs.`))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error importing images: " + error.message))
      .build();
  }
}

/**
 * Shows a card to export relevant columns.
 * @return {CardService.Card} The export columns card.
 */
function showExportColumnsCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Export Relevant Columns"));

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText("This will create a new sheet with only the columns required for your WhatsApp Catalog."))
    .addWidget(CardService.newTextButton()
      .setText("Start Export")
      .setOnClickAction(CardService.newAction().setFunctionName("exportColumnsFromCard"))
    );

  card.addSection(section);
  return card.build();
}

/**
 * Performs column export from the card confirmation.
 * @return {CardService.ActionResponse} The action response after exporting columns.
 */
function exportColumnsFromCard() {
  try {
    const exportSheet = exportRelevantColumns();
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Relevant columns exported to sheet: ${exportSheet.getName()}`))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error exporting columns: " + error.message))
      .build();
  }
}

/**
 * Shows a card to confirm setup of the spreadsheet.
 * @return {CardService.Card} The setup confirmation card.
 */
function showSetupSpreadsheetCard() {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Setup Spreadsheet"));

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText("This will set up your spreadsheet with the correct headers and formatting. Any existing data will be preserved."))
    .addWidget(CardService.newTextButton()
      .setText("Confirm Setup")
      .setOnClickAction(CardService.newAction().setFunctionName("setupSpreadsheetFromCard"))
    );

  card.addSection(section);
  return card.build();
}

/**
 * Performs the spreadsheet setup from the card confirmation.
 * @return {CardService.ActionResponse} The action response after setting up the spreadsheet.
 */
function setupSpreadsheetFromCard() {
  try {
    setupSpreadsheet();
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Spreadsheet set up successfully."))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Error setting up spreadsheet: " + error.message))
      .build();
  }
}