/**
 * Utility functions that support various operations across the application.
 * Contains helper methods used by multiple modules.
 */

// Create namespace for Helpers
const HELPERS = {};

/**
 * Validates an email address format.
 * @param {String} email - The email address to validate.
 * @return {Boolean} Whether the email is valid.
 */
HELPERS.isValidEmail = function(email) {
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email);
};

/**
 * Generates a unique ID for tracking purposes.
 * @return {String} A unique ID.
 */
HELPERS.generateUniqueId = function() {
  return Utilities.getUuid();
};

/**
 * Formats a date as MM/DD/YYYY.
 * @param {Date} date - The date to format.
 * @return {String} The formatted date.
 */
HELPERS.formatDate = function(date = new Date()) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  
  return `${month}/${day}/${year}`;
};

/**
 * Gets all sheet names in the current spreadsheet.
 * @return {Array} List of sheet names.
 */
HELPERS.getSheetNames = function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  return sheets.map(sheet => sheet.getName());
};

/**
 * Finds or creates a sheet with the given name.
 * @param {String} sheetName - The name of the sheet to find or create.
 * @return {Sheet} The sheet object.
 */
HELPERS.getOrCreateSheet = function(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
};

/**
 * Processes text with placeholder substitutions.
 * @param {String} text - The text containing placeholders.
 * @param {Object} data - Key-value pairs for substitution.
 * @return {String} The processed text.
 */
HELPERS.processPlaceholders = function(text, data) {
  if (!text) return '';
  
  let processedText = text;
  
  // Replace all placeholders in format {{key}} with corresponding values
  Object.keys(data).forEach(key => {
    const placeholder = `{{${key}}}`;
    const value = data[key] || '';
    
    // Use regex with global flag to replace all occurrences
    processedText = processedText.replace(new RegExp(placeholder, 'g'), value);
  });
  
  // Add common dynamic placeholders
  processedText = processedText.replace(/{{today}}/g, HELPERS.formatDate());
  processedText = processedText.replace(/{{now}}/g, new Date().toLocaleTimeString());
  
  // Add configured values
  const config = CONFIG.getConfig();
  processedText = processedText.replace(/{{senderName}}/g, config.senderName || '');
  processedText = processedText.replace(/{{companyName}}/g, config.companyName || '');
  
  return processedText;
};

/**
 * Creates a 1px tracking pixel HTML for email open tracking.
 * @param {String} emailId - Unique ID to identify the email.
 * @return {String} HTML img tag for tracking pixel.
 */
HELPERS.createTrackingPixel = function(emailId) {
  const config = CONFIG.getConfig();
  const trackingUrl = config.trackingPixelUrl;
  
  if (!trackingUrl || trackingUrl.includes('YOUR_WEB_APP_URL')) {
    return ''; // No tracking if URL not configured
  }
  
  // Create a tracking URL with the email ID as parameter
  const fullTrackingUrl = `${trackingUrl}?id=${emailId}`;
  
  // Return an HTML img tag for a transparent 1px image
  return `<img src="${fullTrackingUrl}" width="1" height="1" style="display:none;">`;
};

/**
 * Safely parses JSON with error handling.
 * @param {String} jsonString - The JSON string to parse.
 * @param {Object} defaultValue - Default value if parsing fails.
 * @return {Object} The parsed object.
 */
HELPERS.safeJsonParse = function(jsonString, defaultValue = {}) {
  try {
    return JSON.parse(jsonString);
  } catch (e) {
    Logger.logError("JSON parse failed", e);
    return defaultValue;
  }
};

/**
 * Generates a verification code for confirming risky operations.
 * @return {String} A three-word phrase for verification.
 */
HELPERS.generateVerificationPhrase = function() {
  const words = [
    "apple", "banana", "cherry", "grape", "orange", 
    "peach", "melon", "lemon", "kiwi", "plum",
    "dog", "cat", "bird", "fish", "lion",
    "tiger", "bear", "wolf", "fox", "deer",
    "blue", "red", "green", "yellow", "purple",
    "pink", "brown", "black", "white", "orange"
  ];
  
  // Pick 3 random words
  const phrase = [];
  for (let i = 0; i < 3; i++) {
    const randomIndex = Math.floor(Math.random() * words.length);
    phrase.push(words[randomIndex]);
  }
  
  return phrase.join("-");
};

/**
 * Waits for a specified amount of time.
 * @param {Number} seconds - The time to wait in seconds.
 */
HELPERS.wait = function(seconds) {
  Utilities.sleep(seconds * 1000);
};
