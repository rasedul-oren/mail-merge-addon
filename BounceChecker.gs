/**
 * BounceChecker.gs — Detects email bounces by scanning Gmail for
 * mailer-daemon/postmaster messages and updates the Mail Merge Log.
 */

/**
 * Checks Gmail for bounce notifications and updates the Mail Merge Log
 * for any matched recipient emails. Called periodically by a time trigger.
 */
const checkBounces = () => {
  // Time-driven triggers run without UI context — use stored spreadsheet ID
  const ssId = PropertiesService.getScriptProperties().getProperty('spreadsheetId');
  if (!ssId) return;
  const ss = SpreadsheetApp.openById(ssId);
  const logSheet = ss.getSheetByName('Mail Merge Log');

  if (!logSheet || logSheet.getLastRow() < 2) return;

  const data = logSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  // Find column indices
  const emailCol = headers.indexOf('Email');
  const sendStatusCol = headers.indexOf('Send Status');
  const bounceCol = headers.indexOf('Bounce Detected');
  const sendTimeCol = headers.indexOf('Send Time');

  if (emailCol === -1 || sendStatusCol === -1 || bounceCol === -1) return;

  // Collect recipient emails that haven't been marked as bounced yet
  const recipientEmails = new Map(); // email -> [row indices]
  let oldestSendTime = null;

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][sendStatusCol]).trim();
    const bounced = String(data[i][bounceCol]).trim();

    if (status === 'Sent' && bounced !== 'Yes') {
      const email = String(data[i][emailCol]).trim().toLowerCase();
      if (email) {
        if (!recipientEmails.has(email)) {
          recipientEmails.set(email, []);
        }
        recipientEmails.get(email).push(i);
      }
    }

    // Track the oldest send time for trigger cleanup
    const sendTime = data[i][sendTimeCol];
    if (sendTime instanceof Date) {
      if (!oldestSendTime || sendTime < oldestSendTime) {
        oldestSendTime = sendTime;
      }
    }
  }

  if (recipientEmails.size === 0) {
    cleanupBounceTriggers();
    return;
  }

  // Search Gmail for bounce messages from the last day
  try {
    const threads = GmailApp.search(
      'from:mailer-daemon OR from:postmaster subject:"delivery" newer_than:1d'
    );

    const bouncedEmails = new Set();

    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const body = message.getPlainBody().toLowerCase();

        // Check each recipient email against the bounce message body
        recipientEmails.forEach((rowIndices, email) => {
          if (body.includes(email)) {
            bouncedEmails.add(email);
          }
        });
      });
    });

    // Update the log sheet for bounced emails
    bouncedEmails.forEach(email => {
      const rowIndices = recipientEmails.get(email);
      if (rowIndices) {
        rowIndices.forEach(i => {
          const rowNum = i + 1;
          logSheet.getRange(rowNum, bounceCol + 1).setValue('Yes');
          logSheet.getRange(rowNum, sendStatusCol + 1).setValue('Bounced');
        });
      }
    });

  } catch (error) {
    console.error('Bounce check failed:', error.message);
  }

  // Clean up trigger if oldest campaign is more than 24 hours old
  if (oldestSendTime) {
    const hoursSinceOldest = (new Date() - oldestSendTime) / (1000 * 60 * 60);
    if (hoursSinceOldest > 24) {
      cleanupBounceTriggers();
    }
  }
};

/**
 * Creates a time-driven trigger to run checkBounces every 30 minutes.
 * Deletes any existing bounce triggers first to prevent duplicates.
 */
const createBounceTrigger = () => {
  // Remove existing bounce triggers to avoid duplicates
  cleanupBounceTriggers();

  ScriptApp.newTrigger('checkBounces')
    .timeBased()
    .everyMinutes(30)
    .create();
};

/**
 * Finds and deletes all project triggers with the handler function 'checkBounces'.
 */
const cleanupBounceTriggers = () => {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkBounces') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
};
