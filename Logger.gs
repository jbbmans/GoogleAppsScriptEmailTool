/**
 * Provides logging functionality for the email tool.
 * Handles error logging, email sending logs, and tracking logs.
 */

// Create namespace for Logger
const Logger = {};

// Log sheet names
Logger.EMAIL_LOG_SHEET = "Email Log";
Logger.ERROR_LOG_SHEET = "Error Log";
Logger.OPEN_TRACKING_SHEET = "Open Tracking";

// Log column headers
Logger.EMAIL_LOG_HEADERS = [
  "Timestamp", "Type", "Recipient", "Subject", "Status", "Tracking ID"
];

Logger.ERROR_LOG_HEADERS = [
  "Timestamp", "Operation", "Error Details", "Stack"
];

Logger.OPEN_TRACKING_HEADERS = [
  "Timestamp", "Tracking ID", "Email", "Subject", "Original Send Time"
];

/**
 * Initializes log sheets if they don't exist.
 */
Logger.initLogSheets = function() {
  try {
    // Create Email Log sheet if needed
    const emailLogSheet = HELPERS.getOrCreateSheet(Logger.EMAIL_LOG_SHEET);
    const emailLogData = emailLogSheet.getDataRange().getValues();
    
    if (emailLogData.length === 0) {
      emailLogSheet.appendRow(Logger.EMAIL_LOG_HEADERS);
      emailLogSheet.getRange(1, 1, 1, Logger.EMAIL_LOG_HEADERS.length).setFontWeight("bold");
    }
    
    // Create Error Log sheet if needed
    const errorLogSheet = HELPERS.getOrCreateSheet(Logger.ERROR_LOG_SHEET);
    const errorLogData = errorLogSheet.getDataRange().getValues();
    
    if (errorLogData.length === 0) {
      errorLogSheet.appendRow(Logger.ERROR_LOG_HEADERS);
      errorLogSheet.getRange(1, 1, 1, Logger.ERROR_LOG_HEADERS.length).setFontWeight("bold");
    }
    
    // Create Open Tracking sheet if needed
    const trackingSheet = HELPERS.getOrCreateSheet(Logger.OPEN_TRACKING_SHEET);
    const trackingData = trackingSheet.getDataRange().getValues();
    
    if (trackingData.length === 0) {
      trackingSheet.appendRow(Logger.OPEN_TRACKING_HEADERS);
      trackingSheet.getRange(1, 1, 1, Logger.OPEN_TRACKING_HEADERS.length).setFontWeight("bold");
    }
    
    return true;
  } catch (error) {
    console.error("Failed to initialize log sheets:", error);
    return false;
  }
};

/**
 * Logs an error.
 * @param {String} operation - The operation that caused the error.
 * @param {Error|String} error - The error object or message.
 */
Logger.logError = function(operation, error) {
  try {
    Logger.initLogSheets();
    
    const errorLogSheet = HELPERS.getOrCreateSheet(Logger.ERROR_LOG_SHEET);
    
    // Create error log entry
    const timestamp = new Date();
    let errorMessage = error;
    let errorStack = "";
    
    // Handle both Error objects and string errors
    if (error instanceof Error) {
      errorMessage = error.message;
      errorStack = error.stack || "";
    }
    
    // Append error log
    errorLogSheet.appendRow([
      timestamp,
      operation,
      errorMessage,
      errorStack
    ]);
    
    // Log to console as well
    console.error(`[${timestamp}] ${operation}: ${errorMessage}`);
    
  } catch (e) {
    // If error logging fails, log to console as fallback
    console.error("Failed to log error:", e);
    console.error("Original error:", operation, error);
  }
};

/**
 * Logs an email send event.
 * @param {Object} emailData - Data about the email that was sent.
 */
Logger.logEmail = function(emailData) {
  try {
    Logger.initLogSheets();
    
    const emailLogSheet = HELPERS.getOrCreateSheet(Logger.EMAIL_LOG_SHEET);
    
    // Create email log entry
    const timestamp = new Date();
    const type = emailData.type || "unknown";
    const recipient = emailData.recipient || "";
    const subject = emailData.subject || "";
    const status = emailData.status || "unknown";
    const trackingId = emailData.trackingId || "";
    
    // Append email log
    emailLogSheet.appendRow([
      timestamp,
      type,
      recipient,
      subject,
      status,
      trackingId
    ]);
    
  } catch (error) {
    Logger.logError("Failed to log email", error);
  }
};

/**
 * Logs an email open event.
 * @param {Object} openData - Data about the email that was opened.
 */
Logger.logEmailOpen = function(openData) {
  try {
    Logger.initLogSheets();
    
    const trackingSheet = HELPERS.getOrCreateSheet(Logger.OPEN_TRACKING_SHEET);
    const emailLogSheet = HELPERS.getOrCreateSheet(Logger.EMAIL_LOG_SHEET);
    
    // Find original email data
    const trackingId = openData.trackingId;
    let email = "";
    let subject = "";
    let originalTimestamp = null;
    
    // Search for the tracking ID in the email log
    if (emailLogSheet) {
      const emailLogData = emailLogSheet.getDataRange().getValues();
      
      // Skip header row
      for (let i = 1; i < emailLogData.length; i++) {
        const row = emailLogData[i];
        if (row[5] === trackingId) { // Tracking ID is 6th column
          originalTimestamp = row[0]; // Timestamp is 1st column
          email = row[2]; // Recipient is 3rd column
          subject = row[3]; // Subject is 4th column
          break;
        }
      }
    }
    
    // Create open tracking log entry
    const timestamp = new Date();
    
    // Append tracking log
    trackingSheet.appendRow([
      timestamp,
      trackingId,
      email,
      subject,
      originalTimestamp
    ]);
    
  } catch (error) {
    Logger.logError("Failed to log email open", error);
  }
};

/**
 * Gets recent log entries.
 * @param {String} logType - Type of log to retrieve.
 * @param {Number} limit - Maximum number of entries to return.
 * @return {Array} Log entries.
 */
Logger.getRecentLogs = function(logType, limit = 100) {
  try {
    Logger.initLogSheets();
    
    let sheetName;
    switch (logType) {
      case "email":
        sheetName = Logger.EMAIL_LOG_SHEET;
        break;
      case "error":
        sheetName = Logger.ERROR_LOG_SHEET;
        break;
      case "tracking":
        sheetName = Logger.OPEN_TRACKING_SHEET;
        break;
      default:
        throw new Error(`Unknown log type: ${logType}`);
    }
    
    const sheet = HELPERS.getOrCreateSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { headers: data[0] || [], entries: [] };
    }
    
    // Get headers and entries
    const headers = data[0];
    
    // Get entries (newest first) up to the limit
    const entries = data.slice(1)
      .sort((a, b) => b[0] - a[0]) // Sort by timestamp desc
      .slice(0, limit);
    
    return { headers, entries };
    
  } catch (error) {
    Logger.logError(`Failed to get ${logType} logs`, error);
    return { headers: [], entries: [] };
  }
};

/**
 * Clears old log entries beyond a certain threshold.
 * @param {String} logType - Type of log to clear.
 * @param {Number} daysToKeep - Number of days of logs to keep.
 * @return {Object} Result of the operation.
 */
Logger.clearOldLogs = function(logType, daysToKeep = 30) {
  try {
    Logger.initLogSheets();
    
    let sheetName;
    switch (logType) {
      case "email":
        sheetName = Logger.EMAIL_LOG_SHEET;
        break;
      case "error":
        sheetName = Logger.ERROR_LOG_SHEET;
        break;
      case "tracking":
        sheetName = Logger.OPEN_TRACKING_SHEET;
        break;
      default:
        throw new Error(`Unknown log type: ${logType}`);
    }
    
    const sheet = HELPERS.getOrCreateSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        message: `No ${logType} logs to clear.`
      };
    }
    
    // Calculate cutoff date
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - daysToKeep);
    
    // Find rows to keep (header + recent entries)
    const rowsToKeep = [data[0]]; // Keep header row
    let deletedCount = 0;
    
    // Add recent entries
    for (let i = 1; i < data.length; i++) {
      const timestamp = data[i][0];
      if (timestamp instanceof Date && timestamp >= cutoff) {
        rowsToKeep.push(data[i]);
      } else {
        deletedCount++;
      }
    }
    
    // Clear sheet and set new data if any rows were deleted
    if (deletedCount > 0) {
      sheet.clear();
      if (rowsToKeep.length > 0) {
        sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
        // Restore header formatting
        sheet.getRange(1, 1, 1, rowsToKeep[0].length).setFontWeight("bold");
      }
    }
    
    return {
      success: true,
      message: `Cleared ${deletedCount} old ${logType} log entries.`
    };
    
  } catch (error) {
    Logger.logError(`Failed to clear ${logType} logs`, error);
    return {
      success: false,
      message: `Failed to clear logs: ${error.toString()}`
    };
  }
};
