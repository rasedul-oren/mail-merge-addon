/**
 * Settings.gs — User preferences stored in PropertiesService.
 * Manages throttle delay, tracking toggles, and default email fields.
 */

/** Default settings values. */
const DEFAULTS = {
  throttleDelay: 2,       // Seconds between each email send
  trackOpens: true,       // Inject open-tracking pixel
  trackClicks: true,      // Rewrite links for click tracking
  replyTo: '',            // Default reply-to address
  defaultCc: '',          // Default CC addresses
  defaultBcc: ''          // Default BCC addresses
};

/**
 * Reads user settings from PropertiesService, merges with defaults,
 * and returns the complete settings object.
 * @returns {Object} Settings with all keys guaranteed to have values.
 */
const getSettings = () => {
  const props = PropertiesService.getUserProperties();
  const stored = props.getProperty('mailMergeSettings');

  if (!stored) return { ...DEFAULTS };

  try {
    const parsed = JSON.parse(stored);
    // Merge with defaults so any missing keys get default values
    return { ...DEFAULTS, ...parsed };
  } catch (error) {
    console.error('Failed to parse settings, returning defaults:', error.message);
    return { ...DEFAULTS };
  }
};

/**
 * Saves user settings to PropertiesService.
 * @param {Object} settings — Settings object to persist.
 * @returns {Object} Confirmation.
 */
const saveSettings = (settings) => {
  const props = PropertiesService.getUserProperties();
  props.setProperty('mailMergeSettings', JSON.stringify(settings));

  return { success: true, message: 'Settings saved.' };
};
