/**
 * Sender.gs — Email composition, merging, sending, and scheduling logic.
 */

/**
 * Replaces all {{Field Name}} placeholders in a template string with
 * corresponding values from rowData. Missing fields become empty strings.
 * @param {string} template — Template string containing {{placeholders}}.
 * @param {Object} rowData — Key-value pairs of field names and their values.
 * @returns {string} Merged string with placeholders replaced.
 */
const mergeFields = (template, rowData) => {
  if (!template) return '';

  return template.replace(/\{\{(.+?)\}\}/g, (match, fieldName) => {
    const key = fieldName.trim();
    const value = rowData[key];
    return value !== undefined && value !== null ? String(value) : '';
  });
};

/**
 * Appends an invisible 1x1 tracking pixel image tag to the HTML body.
 * Inserted before </body> if present, otherwise appended to the end.
 * @param {string} html — Email HTML body.
 * @param {string} trackingId — Unique tracking identifier.
 * @param {string} webAppUrl — Deployed web app URL for the tracking endpoint.
 * @returns {string} HTML with the tracking pixel injected.
 */
const injectTrackingPixel = (html, trackingId, webAppUrl) => {
  const pixelUrl = `${webAppUrl}?type=open&id=${trackingId}`;
  const pixelTag = `<img src="${pixelUrl}" width="1" height="1" style="display:none;" alt="" />`;

  if (html.includes('</body>')) {
    return html.replace('</body>', `${pixelTag}</body>`);
  }

  return html + pixelTag;
};

/**
 * Rewrites all anchor tag href attributes to route through the tracking
 * redirect endpoint. Skips mailto: links.
 * @param {string} html — Email HTML body.
 * @param {string} trackingId — Unique tracking identifier.
 * @param {string} webAppUrl — Deployed web app URL for the tracking endpoint.
 * @returns {string} HTML with links rewritten for click tracking.
 */
const rewriteLinks = (html, trackingId, webAppUrl) => {
  // Match href="..." in anchor tags, skip mailto: links
  return html.replace(/<a\s([^>]*?)href=["']([^"']+)["']([^>]*?)>/gi, (match, before, url, after) => {
    // Skip mailto: links
    if (url.trim().toLowerCase().startsWith('mailto:')) {
      return match;
    }

    const trackingUrl = `${webAppUrl}?type=click&id=${trackingId}&url=${encodeURIComponent(url)}`;
    return `<a ${before}href="${trackingUrl}"${after}>`;
  });
};

/**
 * Main campaign send function. Iterates over contacts, merges fields,
 * optionally injects tracking, sends via GmailApp, and logs results.
 * @param {Object} payload — {campaignName, subject, body, attachments: [{name, data, type}]}.
 * @param {boolean} selectedOnly — Only send to selected rows if true.
 * @returns {Object} Result: {sent, failed, skipped, warnings}.
 */
const sendCampaign = (payload, selectedOnly) => {
  const campaignName = payload.campaignName || '';
  const subject = payload.subject || '';
  const bodyHtml = payload.body || '';
  const options = {
    selectedOnly: selectedOnly || false,
    attachments: payload.attachments || []
  };
  const settings = getSettings();
  const contacts = getContactsData(options.selectedOnly);
  const warnings = [];

  // Validate emails and check for duplicates
  const duplicates = findDuplicateEmails(contacts);
  if (duplicates.length > 0) {
    warnings.push(`Duplicate emails found: ${duplicates.join(', ')}`);
  }

  const invalidEmails = contacts.filter(c => !validateEmail(c.Email));
  if (invalidEmails.length > 0) {
    warnings.push(`${invalidEmails.length} contact(s) with invalid or missing emails will be skipped.`);
  }

  // Generate a unique campaign ID
  const campaignId = generateTrackingId();
  const webAppUrl = getWebAppUrl();
  const logSheet = getOrCreateSheet('Mail Merge Log', LOG_HEADERS);
  const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];

  // Prepare attachments from base64 if provided
  let blobAttachments = [];
  if (options.attachments && options.attachments.length > 0) {
    blobAttachments = options.attachments.map(att => {
      const decoded = Utilities.base64Decode(att.data);
      return Utilities.newBlob(decoded, att.type || att.mimeType, att.name);
    });
  }

  let sent = 0;
  let failed = 0;
  let skipped = 0;

  // Track emails already sent to (for deduplication within campaign)
  const sentEmails = new Set();

  contacts.forEach(contact => {
    const email = String(contact.Email || '').trim();

    // Skip invalid emails
    if (!validateEmail(email)) {
      skipped++;
      return;
    }

    // Skip duplicate sends within same campaign
    if (sentEmails.has(email.toLowerCase())) {
      skipped++;
      warnings.push(`Skipped duplicate: ${email}`);
      return;
    }

    try {
      // Merge placeholders in subject and body
      let mergedSubject = mergeFields(subject, contact);
      let mergedBody = mergeFields(bodyHtml, contact);

      // Generate tracking ID for this recipient
      const trackingId = generateTrackingId();

      // Inject tracking if enabled
      if (settings.trackOpens) {
        mergedBody = injectTrackingPixel(mergedBody, trackingId, webAppUrl);
      }
      if (settings.trackClicks) {
        mergedBody = rewriteLinks(mergedBody, trackingId, webAppUrl);
      }

      // Build send options
      const sendOptions = {
        htmlBody: mergedBody
      };

      // CC / BCC from options or settings
      const cc = options.cc || settings.defaultCc || '';
      const bcc = options.bcc || settings.defaultBcc || '';
      if (cc) sendOptions.cc = cc;
      if (bcc) sendOptions.bcc = bcc;

      // Reply-to from settings
      if (settings.replyTo) {
        sendOptions.replyTo = settings.replyTo;
      }

      // Attachments
      if (blobAttachments.length > 0) {
        sendOptions.attachments = blobAttachments;
      }

      // Send the email
      GmailApp.sendEmail(email, mergedSubject, '', sendOptions);

      // Log the send to the Mail Merge Log sheet
      const logRow = [
        campaignId,
        campaignName,
        email,
        'Sent',
        new Date(),
        0,          // Opens
        '',         // Last Opened
        0,          // Clicks
        '',         // Clicked Links
        'No',       // Bounce Detected
        trackingId
      ];
      logSheet.appendRow(logRow);

      sentEmails.add(email.toLowerCase());
      sent++;

      // Throttle between sends
      const delay = (settings.throttleDelay || 2) * 1000;
      if (delay > 0) {
        Utilities.sleep(delay);
      }

    } catch (error) {
      failed++;
      warnings.push(`Failed to send to ${email}: ${error.message}`);

      // Log the failure
      const logRow = [
        campaignId,
        campaignName,
        email,
        'Failed',
        new Date(),
        0,
        '',
        0,
        '',
        'No',
        ''
      ];
      logSheet.appendRow(logRow);
    }
  });

  // Set up bounce checking after campaign completes
  if (sent > 0) {
    createBounceTrigger();
  }

  return { sent, failed, skipped, warnings };
};

/**
 * Sends a test email to the current user using the first contact row's data
 * for merge field preview.
 * @param {string} subject — Email subject template.
 * @param {string} bodyHtml — Email body HTML template.
 * @param {Array} attachments — Optional [{name, data, mimeType}] attachments.
 * @returns {Object} Result with success status and message.
 */
const sendTestEmail = (payload) => {
  const subject = payload.subject || '';
  const bodyHtml = payload.body || '';
  const attachments = payload.attachments || [];
  const userEmail = getUserEmail();
  const contacts = getContactsData(false);

  // Use first contact's data for merge preview, or empty object
  const sampleData = contacts.length > 0 ? contacts[0] : {};

  const mergedSubject = mergeFields(subject, sampleData);
  let mergedBody = mergeFields(bodyHtml, sampleData);

  const sendOptions = { htmlBody: mergedBody };

  // Process attachments if provided
  if (attachments && attachments.length > 0) {
    sendOptions.attachments = attachments.map(att => {
      const decoded = Utilities.base64Decode(att.data);
      return Utilities.newBlob(decoded, att.type || att.mimeType, att.name);
    });
  }

  try {
    GmailApp.sendEmail(userEmail, `[TEST] ${mergedSubject}`, '', sendOptions);
    return { success: true, message: `Test email sent to ${userEmail}` };
  } catch (error) {
    return { success: false, message: `Test send failed: ${error.message}` };
  }
};

/**
 * Schedules a campaign to be sent at a future time.
 * Stores parameters in PropertiesService and creates a time-driven trigger.
 * @param {string} campaignName — Campaign name.
 * @param {string} subject — Subject template.
 * @param {string} bodyHtml — Body HTML template.
 * @param {Object} options — Send options (selectedOnly, attachments, cc, bcc).
 * @param {string} sendTime — ISO 8601 datetime string for when to send.
 * @returns {Object} Confirmation with scheduled time.
 */
const scheduleSend = (payload, sendTime) => {
  const props = PropertiesService.getUserProperties();

  // Store campaign parameters and spreadsheet ID for later retrieval
  const scheduledData = {
    payload,
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId()
  };
  props.setProperty('scheduledCampaign', JSON.stringify(scheduledData));

  // Create time-driven trigger at the specified send time
  const triggerTime = new Date(sendTime);
  ScriptApp.newTrigger('executeSendScheduled')
    .timeBased()
    .at(triggerTime)
    .create();

  return {
    success: true,
    message: `Campaign "${payload.campaignName}" scheduled for ${triggerTime.toLocaleString()}`
  };
};

/**
 * Executed by the scheduled trigger. Reads stored parameters from
 * PropertiesService, runs the campaign, and cleans up.
 */
const executeSendScheduled = () => {
  const props = PropertiesService.getUserProperties();
  const rawData = props.getProperty('scheduledCampaign');

  if (!rawData) {
    console.error('No scheduled campaign data found in PropertiesService.');
    return;
  }

  const { payload } = JSON.parse(rawData);

  // Execute the campaign
  const result = sendCampaign(payload, false);
  console.log(`Scheduled campaign completed:`, JSON.stringify(result));

  // Clean up stored data
  props.deleteProperty('scheduledCampaign');

  // Delete the trigger that fired this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'executeSendScheduled') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
};

/**
 * Returns a preview of the merged email using the first contact row's data.
 * @param {Object} payload — {subject, body} from the compose view.
 * @returns {Object} {subject, body} with merge fields replaced.
 */
const previewMergedEmail = (payload) => {
  const contacts = getContactsData(false);
  const sampleData = contacts.length > 0 ? contacts[0] : {};

  return {
    subject: mergeFields(payload.subject || '', sampleData),
    body: mergeFields(payload.body || '', sampleData)
  };
};
