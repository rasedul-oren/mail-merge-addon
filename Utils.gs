/**
 * Utils.gs — Shared utility functions, constants, and data helpers.
 */

/** Column headers for the Mail Merge Log sheet. */
const LOG_HEADERS = [
  'Campaign ID',
  'Campaign Name',
  'Email',
  'Send Status',
  'Send Time',
  'Opens',
  'Last Opened',
  'Clicks',
  'Clicked Links',
  'Bounce Detected',
  'Tracking ID'
];

/** Column headers for the Templates sheet. */
const TEMPLATE_HEADERS = [
  'Template Name',
  'Subject',
  'Body HTML',
  'Created Date'
];

/**
 * Finds or creates a sheet tab with the given name and column headers.
 * If the sheet exists, it is returned as-is. If not, a new sheet is created
 * with the headers written to the first row.
 * @param {string} name — Sheet tab name.
 * @param {string[]} headers — Array of column header strings.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The found or created sheet.
 */
const getOrCreateSheet = (name, headers) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }

  return sheet;
};

/**
 * Generates a unique tracking ID using Apps Script's built-in UUID generator.
 * @returns {string} A UUID string.
 */
const generateTrackingId = () => {
  return Utilities.getUuid();
};

/**
 * Reads the first row of a given sheet (or the active sheet) and returns
 * non-empty header strings.
 * @param {string|null} sheetName — Name of the sheet, or null for active sheet.
 * @returns {string[]} Array of column header strings.
 */
const getColumnHeaders = (sheetName = null) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();

  if (!sheet) return [];

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headers.filter(h => String(h).trim() !== '').map(h => String(h).trim());
};

/**
 * Reads the active sheet and returns an array of row objects keyed by column headers.
 * Skips rows where "Do Not Email" column is truthy.
 * @param {boolean} selectedOnly — If true, only return rows currently selected by the user.
 * @returns {Object[]} Array of contact objects.
 */
const getContactsData = (selectedOnly = false) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const doNotEmailIndex = headers.indexOf('Do Not Email');

  // Determine which row numbers are selected (1-based, data rows start at index 1)
  let selectedRows = null;
  if (selectedOnly) {
    selectedRows = new Set();
    const selection = sheet.getActiveRange();
    if (selection) {
      const startRow = selection.getRow();
      const numRows = selection.getNumRows();
      for (let r = startRow; r < startRow + numRows; r++) {
        // Convert sheet row (1-based, header=1) to data array index
        selectedRows.add(r - 1);
      }
    }
  }

  const contacts = [];
  for (let i = 1; i < data.length; i++) {
    // Skip if "Do Not Email" is truthy
    if (doNotEmailIndex !== -1 && data[i][doNotEmailIndex]) continue;

    // Skip if selectedOnly and this row is not in the selection
    if (selectedOnly && selectedRows && !selectedRows.has(i)) continue;

    const contact = {};
    headers.forEach((header, colIndex) => {
      if (header) {
        contact[header] = data[i][colIndex];
      }
    });

    // Store the original row number (1-based, for reference)
    contact._rowIndex = i + 1;
    contacts.push(contact);
  }

  return contacts;
};

/**
 * Validates an email address using a regex pattern.
 * @param {string} email — Email string to validate.
 * @returns {boolean} True if the email is valid.
 */
const validateEmail = (email) => {
  if (!email) return false;
  const pattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return pattern.test(String(email).trim());
};

/**
 * Finds duplicate email addresses in an array of contact objects.
 * @param {Object[]} contacts — Array of contact objects with an "Email" property.
 * @returns {string[]} Array of duplicate email strings.
 */
const findDuplicateEmails = (contacts) => {
  const emailCounts = {};
  const duplicates = [];

  contacts.forEach(contact => {
    const email = String(contact.Email || '').trim().toLowerCase();
    if (email) {
      emailCounts[email] = (emailCounts[email] || 0) + 1;
    }
  });

  Object.keys(emailCounts).forEach(email => {
    if (emailCounts[email] > 1) {
      duplicates.push(email);
    }
  });

  return duplicates;
};

/**
 * Returns the deployed web app URL for tracking purposes.
 * @returns {string} The web app URL.
 */
const getWebAppUrl = () => {
  return ScriptApp.getService().getUrl();
};
