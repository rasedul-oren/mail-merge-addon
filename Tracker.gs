/**
 * Tracker.gs — Open/click tracking via web app GET requests.
 * Handles tracking pixel serves, link redirects, and event logging.
 */

/**
 * Routes incoming GET requests based on the "type" query parameter.
 * Called from doGet() in Code.gs.
 * @param {Object} e — Event object with e.parameter containing query params.
 * @returns {HtmlOutput|TextOutput} Appropriate tracking response.
 */
const handleTracking = (e) => {
  const params = e.parameter || {};
  const type = params.type;
  const trackingId = params.id;

  if (type === 'open' && trackingId) {
    // Log the open event
    logTrackingEvent(trackingId, 'open');
    // Return a transparent 1x1 pixel
    return getTransparentPixel();
  }

  if (type === 'click' && trackingId) {
    const originalUrl = params.url || '';
    // Log the click event with the clicked URL
    logTrackingEvent(trackingId, 'click', { url: decodeURIComponent(originalUrl) });
    // Redirect user to the original destination
    return getRedirectPage(decodeURIComponent(originalUrl));
  }

  // Unknown or missing type — return error
  return ContentService.createTextOutput('Invalid tracking request.')
    .setMimeType(ContentService.MimeType.TEXT);
};

/**
 * Logs a tracking event (open or click) to the Mail Merge Log sheet.
 * Finds the row by Tracking ID and updates the relevant columns.
 * @param {string} trackingId — The unique tracking ID for the recipient.
 * @param {string} eventType — Either 'open' or 'click'.
 * @param {Object} metadata — Additional data (e.g., {url: '...'} for clicks).
 */
const logTrackingEvent = (trackingId, eventType, metadata = {}) => {
  try {
    // Web app runs outside spreadsheet context — use stored ID
    const ssId = PropertiesService.getScriptProperties().getProperty('spreadsheetId');
    if (!ssId) return;
    const ss = SpreadsheetApp.openById(ssId);
    const logSheet = ss.getSheetByName('Mail Merge Log');

    if (!logSheet) return;

    const data = logSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    // Find column indices
    const trackingIdCol = headers.indexOf('Tracking ID');
    const opensCol = headers.indexOf('Opens');
    const lastOpenedCol = headers.indexOf('Last Opened');
    const clicksCol = headers.indexOf('Clicks');
    const clickedLinksCol = headers.indexOf('Clicked Links');

    if (trackingIdCol === -1) return;

    // Find the row matching this tracking ID
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][trackingIdCol]).trim() === trackingId) {
        const rowNum = i + 1; // Convert to 1-based sheet row

        if (eventType === 'open') {
          // Increment opens count
          const currentOpens = parseInt(data[i][opensCol]) || 0;
          logSheet.getRange(rowNum, opensCol + 1).setValue(currentOpens + 1);
          // Update last opened timestamp
          logSheet.getRange(rowNum, lastOpenedCol + 1).setValue(new Date());
        }

        if (eventType === 'click') {
          // Increment clicks count
          const currentClicks = parseInt(data[i][clicksCol]) || 0;
          logSheet.getRange(rowNum, clicksCol + 1).setValue(currentClicks + 1);
          // Append URL to clicked links (comma-separated)
          const currentLinks = String(data[i][clickedLinksCol] || '').trim();
          const clickedUrl = metadata.url || '';
          const updatedLinks = currentLinks ? `${currentLinks}, ${clickedUrl}` : clickedUrl;
          logSheet.getRange(rowNum, clickedLinksCol + 1).setValue(updatedLinks);
        }

        break;
      }
    }
  } catch (error) {
    console.error('Tracking event logging failed:', error.message);
  }
};

/**
 * Returns a ContentService response containing a transparent 1x1 GIF pixel.
 * Used as the response for open-tracking image requests.
 * @returns {TextOutput} Binary GIF content with correct MIME type.
 */
const getTransparentPixel = () => {
  // ContentService cannot serve raw binary images reliably.
  // Return an HTML page with an inline base64 image instead.
  const html = '<html><body><img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" width="1" height="1" /></body></html>';
  return HtmlService.createHtmlOutput(html);
};

/**
 * Returns an HTML page that immediately redirects the user to the target URL.
 * Uses both a meta refresh tag and JavaScript for maximum compatibility.
 * @param {string} url — The original destination URL to redirect to.
 * @returns {HtmlOutput} HTML page with redirect.
 */
const getRedirectPage = (url) => {
  const safeUrl = url.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="refresh" content="0; url=${safeUrl}" />
  <script>window.location.href = "${safeUrl}";</script>
</head>
<body>
  <p>Redirecting... <a href="${safeUrl}">Click here</a> if not redirected.</p>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html);
};
