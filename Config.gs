/**
 * Configuration management system for the mass email tool.
 * Handles default settings, admin configuration, and persistent storage.
 */

// Create namespace for Config
const CONFIG = {};

// Default configuration values
CONFIG.DEFAULTS = {
  adminEmail: Session.getActiveUser().getEmail(),
  senderName: "Mass Email Tool",
  trackingPixelUrl: "https://script.google.com/macros/s/YOUR_WEB_APP_URL/exec?id=",
  companyDomain: "",
  emailDelay: 1, // seconds between emails
  maxEmailsPerDay: 1500,
  theme: "light",
  placeholderTags: [
    { tag: "{{firstName}}", description: "Recipient's first name" },
    { tag: "{{lastName}}", description: "Recipient's last name" },
    { tag: "{{email}}", description: "Recipient's email address" },
    { tag: "{{today}}", description: "Current date" }
  ],
  templates: [
    { 
      name: "Basic Template", 
      subject: "Information from {{companyName}}",
      body: "<h2>Hello {{firstName}},</h2><p>Thank you for your interest in our services.</p><p>Best regards,<br>{{senderName}}</p>"
    },
    { 
      name: "Newsletter Template", 
      subject: "Newsletter: Monthly Update",
      body: "<h1>Monthly Newsletter</h1><h2>Hello {{firstName}},</h2><p>Here are our updates for this month...</p><p>Best regards,<br>{{senderName}}</p>"
    }
  ]
};

/**
 * Gets the entire configuration object.
 * @return {Object} The current configuration.
 */
CONFIG.getConfig = function() {
  const props = PropertiesService.getScriptProperties();
  const storedConfig = props.getProperty('config');
  
  if (!storedConfig) {
    // If no config exists, save and return defaults
    CONFIG.saveConfig(CONFIG.DEFAULTS);
    return CONFIG.DEFAULTS;
  }
  
  try {
    return JSON.parse(storedConfig);
  } catch (e) {
    Logger.logError("Failed to parse config", e);
    // Return defaults if parsing fails
    return CONFIG.DEFAULTS;
  }
};

/**
 * Saves the entire configuration object.
 * @param {Object} configObject - The configuration to save.
 * @return {Boolean} Success status.
 */
CONFIG.saveConfig = function(configObject) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('config', JSON.stringify(configObject));
    return true;
  } catch (e) {
    Logger.logError("Failed to save config", e);
    return false;
  }
};

/**
 * Updates specific configuration properties.
 * @param {Object} updates - Key-value pairs of properties to update.
 * @return {Object} The updated configuration.
 */
CONFIG.updateConfig = function(updates) {
  try {
    const currentConfig = CONFIG.getConfig();
    const updatedConfig = { ...currentConfig, ...updates };
    CONFIG.saveConfig(updatedConfig);
    return updatedConfig;
  } catch (e) {
    Logger.logError("Failed to update config", e);
    throw new Error("Failed to update configuration: " + e.toString());
  }
};

/**
 * Resets configuration to default values.
 * @return {Object} The default configuration.
 */
CONFIG.resetConfig = function() {
  try {
    CONFIG.saveConfig(CONFIG.DEFAULTS);
    return CONFIG.DEFAULTS;
  } catch (e) {
    Logger.logError("Failed to reset config", e);
    throw new Error("Failed to reset configuration: " + e.toString());
  }
};

/**
 * Gets the Web App URL for the current script deployment.
 * Used for tracking pixel and other client-side functionality.
 * @return {String} The Web App URL.
 */
CONFIG.getWebAppUrl = function() {
  try {
    const config = CONFIG.getConfig();
    if (config.trackingPixelUrl && config.trackingPixelUrl.indexOf('YOUR_WEB_APP_URL') === -1) {
      return config.trackingPixelUrl;
    }
    
    // Try to get current deployment
    const scriptId = ScriptApp.getScriptId();
    return `https://script.google.com/macros/s/${scriptId}/exec`;
  } catch (e) {
    Logger.logError("Failed to get Web App URL", e);
    return "";
  }
};
