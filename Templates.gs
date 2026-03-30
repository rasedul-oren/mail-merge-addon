/**
 * Templates.gs — CRUD operations for reusable email templates
 * stored in the "Templates" sheet.
 */

/**
 * Saves a new email template to the Templates sheet.
 * @param {string} name — Template name (identifier).
 * @param {string} subject — Email subject line.
 * @param {string} bodyHtml — Email body HTML content.
 * @returns {Object} Confirmation with template name and timestamp.
 */
const saveTemplate = (name, subject, bodyHtml) => {
  const sheet = getOrCreateSheet('Templates', TEMPLATE_HEADERS);
  const timestamp = new Date();

  sheet.appendRow([name, subject, bodyHtml, timestamp]);

  return {
    success: true,
    message: `Template "${name}" saved.`,
    createdDate: timestamp
  };
};

/**
 * Returns all saved templates from the Templates sheet.
 * @returns {Object[]} Array of {name, subject, bodyHtml, createdDate}.
 */
const getTemplates = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');

  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const templates = [];

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    if (!name) continue;

    templates.push({
      name,
      subject: String(data[i][1] || ''),
      bodyHtml: String(data[i][2] || ''),
      createdDate: data[i][3]
    });
  }

  return templates;
};

/**
 * Loads a specific template by name.
 * @param {string} name — Template name to find.
 * @returns {Object|null} Template {name, subject, bodyHtml} or null if not found.
 */
const loadTemplate = (name) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');

  if (!sheet || sheet.getLastRow() < 2) return null;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name) {
      return {
        name: String(data[i][0]).trim(),
        subject: String(data[i][1] || ''),
        bodyHtml: String(data[i][2] || '')
      };
    }
  }

  return null;
};

/**
 * Deletes a template by name from the Templates sheet.
 * @param {string} name — Template name to delete.
 * @returns {Object} Result with success status.
 */
const deleteTemplate = (name) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'Template not found.' };
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name) {
      sheet.deleteRow(i + 1); // Convert to 1-based row number
      return { success: true, message: `Template "${name}" deleted.` };
    }
  }

  return { success: false, message: `Template "${name}" not found.` };
};
