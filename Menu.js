/**
 * Creates the add-on menu in the Google Sheets UI.
 */
function createMenu() {
  try {
    const menu = SpreadsheetApp.getUi().createAddonMenu();
    menu.addItem("Setup Spreadsheet", "setupSpreadsheet");
    menu.addItem("Configuration", "showConfigurationPrompt");
    menu.addItem("Validate All Data", "validateAllData");
    menu.addItem("Select Image from Drive", "showPicker");
    menu.addItem("Import images from Drive", "getListingImageUrlsAndSetInSheet");
    menu.addItem("Set WhatsApp Images Folder", "showFolderNamePrompt");
    menu.addItem("Export Relevant Columns", "exportRelevantColumns");
    menu.addItem("Instructions", "showInstructions");
    menu.addToUi();

    // const ui = SpreadsheetApp.getUi();
    // ui.createAddonMenu()
    //   .addItem('Setup Spreadsheet', 'setupSpreadsheet')
    //   .addItem('Configuration', 'showConfigurationPrompt')
    //   .addItem('Validate All Data', 'validateAllData')
    //   .addItem('Select Image from Drive', 'showPicker')
    //   .addItem('Import images from Drive', 'getListingImageUrlsAndSetInSheet')
    //   .addItem('Set WhatsApp Images Folder', 'showFolderNamePrompt')
    //   .addItem('Export Relevant Columns', 'exportRelevantColumns')
    //   .addItem('Instructions', 'showInstructions')
    //   .addToUi();
    
    logEvent('Menu created', 'INFO');
  } catch (error) {
    logEvent('Error creating menu: ' + error.message, 'ERROR');
    console.error('Error creating menu:', error);
  }
}

/**
 * Shows a prompt to set the WhatsApp images folder name.
 */
function showFolderNamePrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Set WhatsApp Images Folder', 'Enter the name for the WhatsApp images folder:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const folderName = response.getResponseText().trim();
    try {
      setWhatsAppFolderName(folderName);
      ui.alert('Folder name set to: ' + folderName);
      logEvent(`WhatsApp Images folder name set to: ${folderName}`, 'INFO');
    } catch (error) {
      ui.alert('Error: ' + error.message);
      logEvent(`Error setting WhatsApp Images folder name: ${error.message}`, 'ERROR');
    }
  } else {
    logEvent('Folder name setting cancelled by user', 'INFO');
  }
}

/**
 * Shows the configuration prompt using CardService.
 * @return {Card} The card to display.
 */
function showConfigurationPrompt() {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("WhatsApp Catalog Configuration"));

  var section = CardService.newCardSection();

  // Get the configuration data
  var config = getConfigurationDropdownLists();

  // WhatsApp Catalog Listing Type
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("WhatsApp Catalog Listing Type")
    .setFieldName("product_type")
    .addItems(config.productTypeList.map(function(type) {
      return CardService.newOption(type, type, type === config.preselectedProductType);
    }))
  );

  // Default Currency
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Currency")
    .setFieldName("currency")
    .addItems(config.currencyList.map(function(currency) {
      return CardService.newOption(currency, currency, currency === config.preselectedCurrency);
    }))
  );

  // Default Category
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Category")
    .setFieldName("category")
    .addItems(config.categoryList.map(function(category) {
      return CardService.newOption(category, category, category === config.preselectedCategory);
    }))
  );

  // Default Availability
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Availability")
    .setFieldName("availability")
    .addItems(config.availabilityList.map(function(availability) {
      return CardService.newOption(availability, availability, availability === config.preselectedAvailability);
    }))
  );

  // Default Condition
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Default Condition")
    .setFieldName("condition")
    .addItems(config.conditionList.map(function(condition) {
      return CardService.newOption(condition, condition, condition === config.preselectedCondition);
    }))
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
 * Saves the configuration.
 * @param {Object} e The event object from the card submission.
 * @return {ActionResponse} The action response to acknowledge the save.
 */
function saveConfiguration(e) {
  var formInputs = e.commonEventObject.formInputs;
  
  var formObject = {
    product_type: formInputs.product_type.stringInputs.value[0],
    category: formInputs.category.stringInputs.value[0],
    currency: formInputs.currency.stringInputs.value[0],
    availability: formInputs.availability.stringInputs.value[0],
    condition: formInputs.condition.stringInputs.value[0]
  };

  processForm(formObject);
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Configuration saved successfully"))
    .build();
}

/**
 * Shows the instructions dialog.
 */
function showInstructions() {
  const html = HtmlService.createHtmlOutputFromFile('Instructions')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Instructions');
  logEvent('Instructions displayed', 'INFO');
}

/**
 * Shows the configuration form.
 */
function showConfigurationForm() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile('Dropdown')
      .setWidth(800)
      .setHeight(600);
  ui.showModalDialog(html, 'Configuration');
  logEvent('Configuration form displayed', 'INFO');
}

/**
 * Shows the image picker dialog.
 */
function showPicker() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Picker')
      .setWidth(600)
      .setHeight(425)
      .setTitle('Select an Image');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select an Image');
  logEvent('Image picker displayed', 'INFO');
}
