/**
 * Code.gs — Main entry points for Mail Merge add-on.
 * Handles menu creation, sidebar display, routing, and setup.
 */

/**
 * Runs when the spreadsheet opens. Adds the "Mail Merge" custom menu.
 */
const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('Mail Merge')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Setup Sheets', 'setupSheets')
    .addToUi();
};

/**
 * Opens the sidebar UI from index.html.
 */
const showSidebar = () => {
  // Store spreadsheet ID for use by web app and triggers (which lack active context)
  PropertiesService.getScriptProperties().setProperty(
    'spreadsheetId',
    SpreadsheetApp.getActiveSpreadsheet().getId()
  );

  const html = HtmlService.createTemplateFromFile('sidebar/index')
    .evaluate()
    .setWidth(300)
    .setTitle('Mail Merge');
  SpreadsheetApp.getUi().showSidebar(html);
};

/**
 * Helper to include HTML partials (CSS/JS files) inside templates.
 * Usage in HTML: <?!= include('styles') ?>
 * @param {string} filename — Name of the HTML file to include (without .html extension).
 * @returns {string} Raw HTML content of the file.
 */
const include = (filename) => {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
};

/**
 * Web app GET handler. Routes tracking requests (open/click) to Tracker.gs.
 * @param {Object} e — Event object with query parameters.
 * @returns {HtmlOutput|TextOutput} Tracking response (pixel, redirect, or error).
 */
const doGet = (e) => {
  return handleTracking(e);
};

/**
 * Creates the required helper sheets (Mail Merge Log and Templates)
 * with their respective column headers.
 */
const setupSheets = () => {
  getOrCreateSheet('Mail Merge Log', LOG_HEADERS);
  getOrCreateSheet('Templates', TEMPLATE_HEADERS);
  SpreadsheetApp.getUi().alert('Setup complete. "Mail Merge Log" and "Templates" sheets are ready.');
};

/**
 * Returns the email address of the current user.
 * @returns {string} Active user's email.
 */
const getUserEmail = () => {
  return Session.getActiveUser().getEmail();
};

/**
 * Counts how many rows in the active sheet have a valid email in the "Email" column.
 * @returns {number} Number of rows with valid emails.
 */
const getRecipientCount = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return 0;

  const headers = data[0].map(h => String(h).trim());
  const emailColIndex = headers.indexOf('Email');

  if (emailColIndex === -1) return 0;

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][emailColIndex]).trim();
    if (validateEmail(email)) {
      count++;
    }
  }

  return count;
};
