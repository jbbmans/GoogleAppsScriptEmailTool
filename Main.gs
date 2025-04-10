/**
 * Main entry point for the Mass Email Tool.
 * Handles initialization, menu creation, and the main UI serving.
 */

// Global Constants
const APP_TITLE = "Mass Email Dashboard";
const SIDEBAR_TITLE = "Email Campaign Manager";
const SIDEBAR_WIDTH = 1200;
const SIDEBAR_HEIGHT = 900;

/**
 * Creates custom menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Email Tool')
    .addItem('Open Dashboard', 'showSidebar')
    .addSeparator()
    .addItem('Settings', 'openSettings')
    .addItem('View Logs', 'openLogs')
    .addToUi();
}

/**
 * Shows the main application sidebar.
 */
function showSidebar() {
  const ui = HtmlService.createHtmlOutputFromFile('index')
    .setTitle(SIDEBAR_TITLE)
    .setWidth(SIDEBAR_WIDTH)
    .setHeight(SIDEBAR_HEIGHT);
  
  SpreadsheetApp.getUi().showModalDialog(ui, APP_TITLE);
}

/**
 * Initial data loader that provides configuration to the UI.
 * Called when the UI first loads.
 * @return {Object} Configuration and initial data.
 */
function getInitialData() {
  try {
    const config = CONFIG.getConfig();
    const recipientSheets = RECIPIENTS.getAvailableRecipientSheets();
    const savedDrafts = COMPOSER.getSavedDrafts();
    const analytics = ANALYTICS.getRecentStats();
    const userSettings = SETTINGS.getUserSettings();
    
    Logger.log("Initial data loaded successfully");
    
    return {
      config: config,
      recipientSheets: recipientSheets,
      savedDrafts: savedDrafts,
      analytics: analytics,
      userSettings: userSettings,
      success: true
    };
  } catch (error) {
    Logger.logError("Failed to load initial data", error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Opens the settings modal in the UI.
 */
function openSettings() {
  showSidebar();
  // The UI will detect this flag and open settings modal
  return PropertiesService.getScriptProperties().setProperty('openSettings', 'true');
}

/**
 * Opens the logs view in the UI.
 */
function openLogs() {
  showSidebar();
  // The UI will detect this flag and open logs view
  return PropertiesService.getScriptProperties().setProperty('openLogs', 'true');
}
