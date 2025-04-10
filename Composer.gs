/**
 * Manages email composition, templates, and drafts.
 * Handles saving, loading, and processing of email content.
 */

// Create namespace for Composer
const COMPOSER = {};

// Constants
COMPOSER.DRAFT_STORAGE_KEY = 'emailDrafts';
COMPOSER.MAX_DRAFTS = 20;

/**
 * Gets all saved email drafts.
 * @return {Array} List of saved drafts.
 */
COMPOSER.getSavedDrafts = function() {
  try {
    const props = PropertiesService.getUserProperties();
    const draftsJson = props.getProperty(COMPOSER.DRAFT_STORAGE_KEY);
    
    if (!draftsJson) {
      return [];
    }
    
    return HELPERS.safeJsonParse(draftsJson, []);
  } catch (error) {
    Logger.logError("Error loading saved drafts", error);
    return [];
  }
};

/**
 * Saves a new email draft.
 * @param {Object} draft - Draft object with name, subject, and body.
 * @return {Object} The result of the operation.
 */
COMPOSER.saveDraft = function(draft) {
  try {
    if (!draft.name || !draft.subject || !draft.body) {
      return {
        success: false,
        message: "Draft name, subject, and body are required."
      };
    }
    
    // Set timestamp and ID
    draft.timestamp = new Date().toISOString();
    draft.id = draft.id || HELPERS.generateUniqueId();
    
    // Get existing drafts
    const drafts = COMPOSER.getSavedDrafts();
    
    // Check if draft with same name exists
    const existingIndex = drafts.findIndex(d => d.id === draft.id);
    
    if (existingIndex >= 0) {
      // Update existing draft
      drafts[existingIndex] = draft;
    } else {
      // Add new draft, keeping only newest MAX_DRAFTS
      drafts.unshift(draft);
      if (drafts.length > COMPOSER.MAX_DRAFTS) {
        drafts.pop();
      }
    }
    
    // Save drafts
    const props = PropertiesService.getUserProperties();
    props.setProperty(COMPOSER.DRAFT_STORAGE_KEY, JSON.stringify(drafts));
    
    return {
      success: true,
      message: "Draft saved successfully.",
      draft: draft
    };
    
  } catch (error) {
    Logger.logError("Error saving draft", error);
    return {
      success: false,
      message: "Failed to save draft: " + error.toString()
    };
  }
};

/**
 * Deletes a saved draft.
 * @param {String} draftId - ID of the draft to delete.
 * @return {Object} The result of the operation.
 */
COMPOSER.deleteDraft = function(draftId) {
  try {
    const drafts = COMPOSER.getSavedDrafts();
    const initialCount = drafts.length;
    
    // Filter out the draft to delete
    const updatedDrafts = drafts.filter(draft => draft.id !== draftId);
    
    if (updatedDrafts.length === initialCount) {
      return {
        success: false,
        message: "Draft not found."
      };
    }
    
    // Save updated drafts
    const props = PropertiesService.getUserProperties();
    props.setProperty(COMPOSER.DRAFT_STORAGE_KEY, JSON.stringify(updatedDrafts));
    
    return {
      success: true,
      message: "Draft deleted successfully."
    };
    
  } catch (error) {
    Logger.logError("Error deleting draft", error);
    return {
      success: false,
      message: "Failed to delete draft: " + error.toString()
    };
  }
};

/**
 * Gets available email templates.
 * @return {Array} List of email templates.
 */
COMPOSER.getTemplates = function() {
  try {
    const config = CONFIG.getConfig();
    return config.templates || [];
  } catch (error) {
    Logger.logError("Error loading templates", error);
    return [];
  }
};

/**
 * Processes an email template with recipient data.
 * @param {Object} emailData - Email subject, body, recipient data.
 * @return {Object} Processed email content.
 */
COMPOSER.processEmail = function(emailData) {
  try {
    const { subject, body, recipient } = emailData;
    
    if (!subject || !body || !recipient) {
      throw new Error("Missing required email data");
    }
    
    // Process placeholders in subject and body
    const processedSubject = HELPERS.processPlaceholders(subject, recipient);
    let processedBody = HELPERS.processPlaceholders(body, recipient);
    
    // Add tracking pixel if enabled
    if (emailData.addTracking) {
      const trackingPixel = HELPERS.createTrackingPixel(recipient.trackingId);
      processedBody += trackingPixel;
    }
    
    return {
      success: true,
      subject: processedSubject,
      body: processedBody,
      recipient: recipient
    };
    
  } catch (error) {
    Logger.logError("Error processing email", error);
    return {
      success: false,
      message: "Failed to process email: " + error.toString()
    };
  }
};

/**
 * Saves a new email template.
 * @param {Object} template - Template object with name, subject, and body.
 * @return {Object} The result of the operation.
 */
COMPOSER.saveTemplate = function(template) {
  try {
    if (!template.name || !template.subject || !template.body) {
      return {
        success: false,
        message: "Template name, subject, and body are required."
      };
    }
    
    // Get current templates
    const config = CONFIG.getConfig();
    const templates = config.templates || [];
    
    // Check if template with same name exists
    const existingIndex = templates.findIndex(t => t.name === template.name);
    
    if (existingIndex >= 0) {
      // Update existing template
      templates[existingIndex] = template;
    } else {
      // Add new template
      templates.push(template);
    }
    
    // Save updated templates to config
    CONFIG.updateConfig({ templates });
    
    return {
      success: true,
      message: "Template saved successfully."
    };
    
  } catch (error) {
    Logger.logError("Error saving template", error);
    return {
      success: false,
      message: "Failed to save template: " + error.toString()
    };
  }
};

/**
 * Deletes an email template.
 * @param {String} templateName - Name of the template to delete.
 * @return {Object} The result of the operation.
 */
COMPOSER.deleteTemplate = function(templateName) {
  try {
    const config = CONFIG.getConfig();
    const templates = config.templates || [];
    
    // Filter out the template to delete
    const updatedTemplates = templates.filter(t => t.name !== templateName);
    
    if (updatedTemplates.length === templates.length) {
      return {
        success: false,
        message: "Template not found."
      };
    }
    
    // Save updated templates
    CONFIG.updateConfig({ templates: updatedTemplates });
    
    return {
      success: true,
      message: "Template deleted successfully."
    };
    
  } catch (error) {
    Logger.logError("Error deleting template", error);
    return {
      success: false,
      message: "Failed to delete template: " + error.toString()
    };
  }
};

/**
 * Validates email settings before sending.
 * @param {Object} emailSettings - Settings for the email to send.
 * @return {Object} Validation results.
 */
COMPOSER.validateEmailSettings = function(emailSettings) {
  const errors = [];
  
  if (!emailSettings.subject || emailSettings.subject.trim() === "") {
    errors.push("Subject is required");
  }
  
  if (!emailSettings.body || emailSettings.body.trim() === "") {
    errors.push("Email body is required");
  }
  
  if (emailSettings.cc && emailSettings.cc.trim() !== "") {
    const ccEmails = emailSettings.cc.split(',').map(e => e.trim());
    for (const email of ccEmails) {
      if (!HELPERS.isValidEmail(email)) {
        errors.push(`Invalid CC email: ${email}`);
      }
    }
  }
  
  if (emailSettings.bcc && emailSettings.bcc.trim() !== "") {
    const bccEmails = emailSettings.bcc.split(',').map(e => e.trim());
    for (const email of bccEmails) {
      if (!HELPERS.isValidEmail(email)) {
        errors.push(`Invalid BCC email: ${email}`);
      }
    }
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
};

/**
 * Gets available placeholder tags for templates.
 * @return {Array} List of placeholder tags with descriptions.
 */
COMPOSER.getPlaceholderTags = function() {
  try {
    const config = CONFIG.getConfig();
    return config.placeholderTags || [];
  } catch (error) {
    Logger.logError("Error loading placeholder tags", error);
    return [];
  }
};
