/**
 * Dashboard.gs — Campaign analytics: listing, stats, per-recipient data, and CSV export.
 */

/**
 * Reads the Mail Merge Log and returns a list of unique campaigns,
 * sorted by date descending.
 * @returns {Object[]} Array of {campaignId, campaignName, date, totalSent}.
 */
const getCampaignList = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Mail Merge Log');

  if (!logSheet || logSheet.getLastRow() < 2) return [];

  const data = logSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const campaignIdCol = headers.indexOf('Campaign ID');
  const campaignNameCol = headers.indexOf('Campaign Name');
  const sendTimeCol = headers.indexOf('Send Time');
  const sendStatusCol = headers.indexOf('Send Status');

  if (campaignIdCol === -1) return [];

  // Aggregate by campaign ID
  const campaigns = new Map();

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][campaignIdCol]).trim();
    if (!id) continue;

    if (!campaigns.has(id)) {
      campaigns.set(id, {
        campaignId: id,
        campaignName: String(data[i][campaignNameCol] || '').trim(),
        date: data[i][sendTimeCol],
        totalSent: 0
      });
    }

    const campaign = campaigns.get(id);

    // Count sent emails (including bounced, which were initially sent)
    const status = String(data[i][sendStatusCol]).trim();
    if (status === 'Sent' || status === 'Bounced') {
      campaign.totalSent++;
    }

    // Use the earliest send time as the campaign date
    const rowTime = data[i][sendTimeCol];
    if (rowTime instanceof Date && (!(campaign.date instanceof Date) || rowTime < campaign.date)) {
      campaign.date = rowTime;
    }
  }

  // Convert to array and sort by date descending
  return Array.from(campaigns.values()).sort((a, b) => {
    const dateA = a.date instanceof Date ? a.date.getTime() : 0;
    const dateB = b.date instanceof Date ? b.date.getTime() : 0;
    return dateB - dateA;
  });
};

/**
 * Returns aggregate statistics for a specific campaign.
 * @param {string} campaignId — The campaign UUID to filter by.
 * @returns {Object} Stats: {sent, opens, uniqueOpens, openRate, clicks, uniqueClicks, clickRate, bounces, bounceRate}.
 */
const getCampaignStats = (campaignId) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Mail Merge Log');

  if (!logSheet || logSheet.getLastRow() < 2) {
    return { sent: 0, opens: 0, uniqueOpens: 0, openRate: 0, clicks: 0, uniqueClicks: 0, clickRate: 0, bounces: 0, bounceRate: 0 };
  }

  const data = logSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const idCol = headers.indexOf('Campaign ID');
  const opensCol = headers.indexOf('Opens');
  const clicksCol = headers.indexOf('Clicks');
  const bounceCol = headers.indexOf('Bounce Detected');
  const sendStatusCol = headers.indexOf('Send Status');

  let sent = 0;
  let totalOpens = 0;
  let uniqueOpens = 0;
  let totalClicks = 0;
  let uniqueClicks = 0;
  let bounces = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() !== campaignId) continue;

    const status = String(data[i][sendStatusCol]).trim();
    if (status === 'Sent' || status === 'Bounced') sent++;

    const opens = parseInt(data[i][opensCol]) || 0;
    totalOpens += opens;
    if (opens > 0) uniqueOpens++;

    const clicks = parseInt(data[i][clicksCol]) || 0;
    totalClicks += clicks;
    if (clicks > 0) uniqueClicks++;

    if (String(data[i][bounceCol]).trim() === 'Yes') bounces++;
  }

  return {
    sent,
    opens: totalOpens,
    uniqueOpens,
    openRate: sent > 0 ? Math.round((uniqueOpens / sent) * 100) : 0,
    clicks: totalClicks,
    uniqueClicks,
    clickRate: sent > 0 ? Math.round((uniqueClicks / sent) * 100) : 0,
    bounces,
    bounceRate: sent > 0 ? Math.round((bounces / sent) * 100) : 0
  };
};

/**
 * Returns per-recipient details for a specific campaign.
 * @param {string} campaignId — The campaign UUID to filter by.
 * @returns {Object[]} Array of {email, sendStatus, opens, clicks, clickedLinks, lastActivity, bounced}.
 */
const getCampaignRecipients = (campaignId) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Mail Merge Log');

  if (!logSheet || logSheet.getLastRow() < 2) return [];

  const data = logSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const idCol = headers.indexOf('Campaign ID');
  const emailCol = headers.indexOf('Email');
  const sendStatusCol = headers.indexOf('Send Status');
  const opensCol = headers.indexOf('Opens');
  const clicksCol = headers.indexOf('Clicks');
  const clickedLinksCol = headers.indexOf('Clicked Links');
  const lastOpenedCol = headers.indexOf('Last Opened');
  const bounceCol = headers.indexOf('Bounce Detected');

  const recipients = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() !== campaignId) continue;

    recipients.push({
      email: String(data[i][emailCol]).trim(),
      sendStatus: String(data[i][sendStatusCol]).trim(),
      opens: parseInt(data[i][opensCol]) || 0,
      clicks: parseInt(data[i][clicksCol]) || 0,
      clickedLinks: String(data[i][clickedLinksCol] || '').trim(),
      lastActivity: data[i][lastOpenedCol] || '',
      bounced: String(data[i][bounceCol]).trim() === 'Yes'
    });
  }

  return recipients;
};

/**
 * Exports campaign recipient data as a CSV string for download.
 * @param {string} campaignId — The campaign UUID to export.
 * @returns {string} CSV-formatted string of recipient data.
 */
const exportCampaignCsv = (campaignId) => {
  const recipients = getCampaignRecipients(campaignId);
  const csvHeaders = ['Email', 'Send Status', 'Opens', 'Clicks', 'Clicked Links', 'Last Activity', 'Bounced'];

  const rows = [csvHeaders.join(',')];

  recipients.forEach(r => {
    const row = [
      `"${r.email}"`,
      `"${r.sendStatus}"`,
      r.opens,
      r.clicks,
      `"${r.clickedLinks.replace(/"/g, '""')}"`,
      r.lastActivity instanceof Date ? `"${r.lastActivity.toISOString()}"` : `"${r.lastActivity}"`,
      r.bounced ? 'Yes' : 'No'
    ];
    rows.push(row.join(','));
  });

  return rows.join('\n');
};
