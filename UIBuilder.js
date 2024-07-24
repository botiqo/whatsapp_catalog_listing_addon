const CARD_TITLES = {
    HOMEPAGE: "WhatsApp Catalog Tools",
    CONFIGURATION: "WhatsApp Catalog Configuration",
    IMAGE_PICKER: "Select an Image",
    IMPORT_IMAGES: "Import Images from Drive",
    EXPORT_COLUMNS: "Export Relevant Columns",
    INSTRUCTIONS: "Instructions",
    VALIDATE_DATA: "Validate All Data"
};

/**
 * Creates a standard card with a header.
 * @param {string} title The title of the card.
 * @return {CardService.CardBuilder} The card builder.
 */
function createBaseCard(title) {
    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle(title));
}

/**
 * Creates a button with an action.
 * @param {string} text The button text.
 * @param {string} functionName The function to call when clicked.
 * @return {CardService.TextButton} The button.
 */
function createActionButton(text, functionName) {
    return CardService.newTextButton()
        .setText(text)
        .setOnClickAction(CardService.newAction().setFunctionName(functionName));
}

/**
 * Creates a section with a header and widgets.
 * @param {string} header The section header.
 * @param {CardService.Widget[]} widgets The widgets to add to the section.
 * @return {CardService.CardSection} The card section.
 */
function createSection(header, widgets) {
    const section = CardService.newCardSection().setHeader(header);
    widgets.forEach(widget => section.addWidget(widget));
    return section;
}

/**
 * Creates a selection input.
 * @param {string} title The input title.
 * @param {string} fieldName The field name.
 * @param {string[]} items The list of items.
 * @param {string} selectedItem The preselected item.
 * @return {CardService.SelectionInput} The selection input.
 */
function createSelectionInput(title, fieldName, items, selectedItem) {
    const input = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle(title)
        .setFieldName(fieldName);

    items.forEach(item => {
        input.addItem(item, item, item === selectedItem);
    });

    return input;
}

/**
 * Creates and returns a card for the add-on homepage.
 * @return {CardService.Card} The homepage card.
 */
function createHomepageCard() {
    const card = createBaseCard(CARD_TITLES.HOMEPAGE);
    const buttons = [
      createActionButton("Setup Spreadsheet", "setupSpreadsheet"),
      createActionButton("Configuration", "createConfigurationCard"),
      createActionButton("Validate All Data", "validateAllProducts"),
      createActionButton("Import images from Drive", "createImportImagesCard"),
      createActionButton("Export Relevant Columns", "createExportColumnsCard"),
      createActionButton("Instructions", "createInstructionsCard"),
      createActionButton("Provide Feedback", "showFeedbackCard") // Add this line
    ];

    card.addSection(createSection(CARD_TITLES.HOMEPAGE, buttons));
    return card.build();
  }

/**
 * Creates and returns a card for the configuration settings.
 * @return {CardService.Card} The configuration card.
 */
function createConfigurationCard() {
    const card = createBaseCard(CARD_TITLES.CONFIGURATION);
    const config = getConfigurationDropdownLists();

    const inputs = [
        createSelectionInput("WhatsApp Catalog Listing Type", "product_type", config.productTypeList, config.preselectedProductType),
        createSelectionInput("Default Currency", "currency", config.currencyList, config.preselectedCurrency),
        createSelectionInput("Default Category", "category", config.categoryList, config.preselectedCategory),
        createSelectionInput("Default Availability", "availability", config.availabilityList, config.preselectedAvailability),
        createSelectionInput("Default Condition", "condition", config.conditionList, config.preselectedCondition)
    ];

    const saveButton = createActionButton("Save Configuration", "saveConfiguration");

    card.addSection(createSection("", [...inputs, saveButton]));
    return card.build();
}

/**
 * Creates and returns a card for managing WhatsApp image folder and importing images.
 * @return {CardService.Card} The image management card.
 */
function createImportImagesCard() {
    const card = createBaseCard(CARD_TITLES.IMPORT_IMAGES);
    const folderName = PropertiesService.getUserProperties().getProperty('WHATSAPP_FOLDER_NAME') || "Not set";
    const folderId = PropertiesService.getUserProperties().getProperty('WHATSAPP_FOLDER_ID');

    ErrorHandler.log(`createImportImagesCard: Current folder - ${folderName}`, 'INFO');

    const folderSection = CardService.newCardSection()
        .setHeader("WhatsApp Images Folder")
        .addWidget(CardService.newTextParagraph().setText(`Current folder: ${folderName}`))
        .addWidget(CardService.newTextButton()
            .setText("Select Folder")
            .setOnClickAction(CardService.newAction().setFunctionName("showFolderPicker")));

    card.addSection(folderSection);

    if (folderId) {
        const importSection = CardService.newCardSection()
            .setHeader("Import Images")
            .addWidget(CardService.newTextParagraph().setText("Import image URLs from your WhatsApp Catalog Listing folder in Google Drive."))
            .addWidget(createActionButton("Start Import", "importImagesFromCard"));

        card.addSection(importSection);

        const selectImageSection = CardService.newCardSection()
            .setHeader("Select Individual Image")
            .addWidget(CardService.newTextParagraph().setText("Choose a single image to add to your catalog."))
            .addWidget(CardService.newTextButton()
                .setText("Select Image")
                .setOnClickAction(CardService.newAction().setFunctionName("showImagePicker")));

        card.addSection(selectImageSection);
    } else {
        const warningSection = CardService.newCardSection()
            .addWidget(CardService.newTextParagraph()
                .setText("Please select a WhatsApp Images Folder to enable importing and image selection."));

        card.addSection(warningSection);
    }

    return card.build();
}

/**
 * Create a card to export relevant columns.
 * @return {CardService.Card} The export columns card.
 */
function createExportColumnsCard() {
    const card = createBaseCard(CARD_TITLES.EXPORT_COLUMNS);

    const widgets = [
        CardService.newTextParagraph().setText("This will create a new sheet with only the columns required for your WhatsApp Catalog."),
        createActionButton("Export", "exportColumnsFromCard")
    ];

    card.addSection(createSection("", widgets));
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
 * Creates and returns a card with instructions.
 * @return {CardService.Card} The instructions card.
 */
function createInstructionsCard() {
    const card = createBaseCard(CARD_TITLES.INSTRUCTIONS);
    const instructions = [
        "1. Click on 'Setup Spreadsheet' to initialize your sheet with the correct headers and formatting.",
        "2. Use 'Configuration' to set default values for product type, currency, category, availability, and condition.",
        "3. Use 'Set WhatsApp Images Folder' to specify the Google Drive folder for your catalog images.",
        "4. 'Import images from Drive' will populate your sheet with image URLs from the specified folder.",
        "5. Use 'Select Image from Drive' to choose individual images for each product.",
        "6. 'Validate All Data' checks your catalog for any errors or missing information.",
        "7. When your catalog is ready, use 'Export Relevant Columns' to create a new sheet with only the required data for WhatsApp."
    ];

    const widgets = instructions.map(instruction => CardService.newTextParagraph().setText(instruction));
    card.addSection(createSection("", widgets));
    return card.build();
}

function showLoadingCard(message) {
    const card = CardService.newCardBuilder();
    const section = CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(message))
        .addWidget(CardService.newImage().setImageUrl(LOADER_IMG_URL));
    card.addSection(section);
    return card.build();
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

    // Clear the cache when a new folder is selected
    CacheManager.clear();

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`Folder "${folderName}" selected as WhatsApp Images Folder.`))
        .setNavigation(CardService.newNavigation().pushCard(createImportImagesCard()))
        .build();
}

function showProgressCard(message, progress) {
    const card = CardService.newCardBuilder();
    const section = CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(message))
        .addWidget(CardService.newProgressBar().setProgress(progress));
    card.addSection(section);
    return card.build();
}

function createValidationResultsCard(errors) {
    const card = CardService.newCardBuilder();
    const section = CardService.newCardSection();

    if (errors.length === 0) {
      section.addWidget(CardService.newTextParagraph().setText("All products are valid!"));
    } else {
      section.addWidget(CardService.newTextParagraph().setText(`Found ${errors.length} validation errors:`));

      errors.forEach((error, index) => {
        if (index < 10) { // Limit to first 10 errors to avoid card size limits
          section.addWidget(CardService.newTextParagraph().setText(error));
        }
      });

      if (errors.length > 10) {
        section.addWidget(CardService.newTextParagraph().setText(`... and ${errors.length - 10} more errors.`));
      }
    }

    card.addSection(section);
    return card.build();
  }