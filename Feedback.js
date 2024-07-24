// Store this in Script Properties
const FEEDBACK_SHEET_ID = '1o2tbUlV30rVAPqSWTt5kPziipeNjrMWXwB-JZ8mjYKM';

/**
 * Shows the feedback card.
 * @return {CardService.Card} The feedback card.
 */
function showFeedbackCard() {
  const card = CardService.newCardBuilder();
  const section = CardService.newCardSection()
    .setHeader("Provide Feedback");

  const feedbackInput = CardService.newTextInput()
    .setFieldName("feedback")
    .setTitle("Your Feedback")
    .setMultiline(true);

  const ratingInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Rating")
    .setFieldName("rating");

  for (let i = 1; i <= 5; i++) {
    ratingInput.addItem(i.toString(), i.toString(), false);
  }

  const submitButton = CardService.newTextButton()
    .setText("Submit Feedback")
    .setOnClickAction(CardService.newAction().setFunctionName("submitFeedback"));

  section.addWidget(feedbackInput)
    .addWidget(ratingInput)
    .addWidget(submitButton);

  card.addSection(section);
  return card.build();
}

/**
 * Submits the feedback.
 * @param {Object} e The event object from the card action.
 * @return {CardService.ActionResponse} The action response after submitting feedback.
 */
function submitFeedback(e) {
    const feedback = e.formInput.feedback;
    const rating = e.formInput.rating;

    if (!feedback || !rating) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("Please provide both feedback and rating."))
        .build();
    }

    try {
      saveFeedback(feedback, rating);
      updateLastFeedbackTime();
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("Thank you for your feedback!"))
        .setNavigation(CardService.newNavigation().popToRoot())
        .build();
    } catch (error) {
      Logger.log("Error in submitFeedback: " + error.message);
      Logger.log("Stack trace: " + error.stack);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("Error submitting feedback: " + error.message))
        .build();
    }
  }

/**
 * Saves the feedback to the developer's Google Sheet.
 * @param {string} feedback The user's feedback.
 * @param {string} rating The user's rating.
 */
function saveFeedback(feedback, rating) {
    try {
      const feedbackSheetId = FEEDBACK_SHEET_ID;
      Logger.log("Feedback Sheet ID from properties: " + feedbackSheetId);

      if (!feedbackSheetId) {
        throw new Error("Feedback sheet ID not set in Script Properties");
      }

      Logger.log("Attempting to open spreadsheet with ID: " + feedbackSheetId);
      const ss = SpreadsheetApp.openById(feedbackSheetId);
      if (!ss) {
        throw new Error("Could not open the feedback spreadsheet");
      }

      Logger.log("Spreadsheet opened successfully: " + ss.getName());
      let sheet = ss.getSheetByName("Feedback");

      if (!sheet) {
        Logger.log("Feedback sheet not found, creating new sheet");
        sheet = ss.insertSheet("Feedback");
        sheet.appendRow(["Timestamp", "User", "Feedback", "Rating"]);
        Logger.log("New 'Feedback' sheet created and header row added");
      } else {
        Logger.log("Existing 'Feedback' sheet found");
      }

      const timestamp = new Date();
      const user = Session.getActiveUser().getEmail();

      Logger.log("Appending new row with feedback");
      Logger.log("Timestamp: " + timestamp);
      Logger.log("User: " + user);
      Logger.log("Feedback: " + feedback);
      Logger.log("Rating: " + rating);

      sheet.appendRow([timestamp, user, feedback, rating]);
      Logger.log("New row appended to 'Feedback' sheet");

      // Verify the append
      const lastRow = sheet.getLastRow();
      const lastRowData = sheet.getRange(lastRow, 1, 1, 4).getValues()[0];
      Logger.log("Last row data: " + JSON.stringify(lastRowData));

    } catch (error) {
      Logger.log("Error in saveFeedback: " + error.message);
      Logger.log("Stack trace: " + error.stack);
      throw error;  // Re-throw the error so it can be caught in the calling function
    }
  }

/**
 * Checks if the user has recently provided feedback.
 * @return {boolean} True if the user has provided feedback in the last 7 days.
 */
function hasRecentlyProvidedFeedback() {
  const userProperties = PropertiesService.getUserProperties();
  const lastFeedbackTime = userProperties.getProperty('lastFeedbackTime');

  if (!lastFeedbackTime) {
    return false;
  }

  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  return new Date(lastFeedbackTime) > sevenDaysAgo;
}

/**
 * Updates the last feedback time for the user.
 */
function updateLastFeedbackTime() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('lastFeedbackTime', new Date().toISOString());
}

function testFeedbackSheetAccess() {
    try {
      const feedbackSheetId = FEEDBACK_SHEET_ID;
      Logger.log("Feedback Sheet ID from properties: " + feedbackSheetId);

      if (!feedbackSheetId) {
        throw new Error("FEEDBACK_SHEET_ID is not set in Script Properties");
      }

      const ss = SpreadsheetApp.openById(feedbackSheetId);
      Logger.log("Spreadsheet name: " + ss.getName());
      Logger.log("Sheets: " + ss.getSheets().map(sheet => sheet.getName()).join(", "));

      // Test writing to the sheet
      let sheet = ss.getSheetByName("Feedback");
      if (!sheet) {
        Logger.log("Feedback sheet not found, creating new sheet");
        sheet = ss.insertSheet("Feedback");
        sheet.appendRow(["Timestamp", "User", "Feedback", "Rating"]);
      }

      saveFeedback("Test Feedback", "5");
      Logger.log("Test row appended successfully");

    } catch (error) {
      Logger.log("Error in testFeedbackSheetAccess: " + error.message);
      Logger.log("Stack trace: " + error.stack);
    }
  }