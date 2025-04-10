/**
 * Manages email sending operations for the mass email tool.
 * Handles test emails, batch emails, and tracking.
 */

// Create namespace for EmailManager
const EMAILMANAGER = {};

// Constants
EMAILMANAGER.MAX_BATCH_SIZE = 50; // Maximum recipients per batch
EMAILMANAGER.DEFAULT_EMAIL_QUOTA = 1500; // Default daily quota 

/**
 * Sends a test email.
 * @param {Object} emailData - Email data including subject, body, and recipient.
 * @return {Object} Result of send operation.
 */
EMAILMANAGER.sendTestEmail = function(emailData) {
  try {
    const config = CONFIG.getConfig();
    const { subject, body, testEmail } = emailData;
    
    if (!subject || !body || !testEmail) {
      return {
        success: false,
        message: "Subject, body and test email address are required."
      };
    }
    
    if (!HELPERS.isValidEmail(testEmail)) {
      return {
        success: false,
        message: "Invalid test email address."
      };
    }
    
    // Create test recipient
    const recipient = {
      firstName: "Test",
      lastName: "User",
      email: testEmail,
      id: HELPERS.generateUniqueId(),
      trackingId: HELPERS.generateUniqueId()
    };
    
    // Process the email content
    const processedEmail = COMPOSER.processEmail({
      subject: subject,
      body: body,
      recipient: recipient,
      addTracking: emailData.addTracking || false
    });
    
    if (!processedEmail.success) {
      throw new Error(processedEmail.message);
    }
    
    // Send the email
    GmailApp.sendEmail(
      recipient.email,
      processedEmail.subject,
      "This email contains HTML content that cannot be displayed in plain text.",
      {
        name: config.senderName || "Mass Email Tool",
        htmlBody: processedEmail.body,
        cc: emailData.cc || "",
        bcc: emailData.bcc || ""
      }
    );
    
    // Log the test email
    Logger.logEmail({
      type: "test",
      recipient: recipient.email,
      subject: processedEmail.subject,
      status: "sent",
      trackingId: recipient.trackingId
    });
    
    return {
      success: true,
      message: "Test email sent successfully to " + recipient.email
    };
    
  } catch (error) {
    Logger.logError("Error sending test email", error);
    return {
      success: false,
      message: "Failed to send test email: " + error.toString()
    };
  }
};

/**
 * Validates a batch email request.
 * @param {Object} batchData - Batch email configuration and recipients.
 * @return {Object} Validation result.
 */
EMAILMANAGER.validateBatchEmail = function(batchData) {
  try {
    const { subject, body, recipients } = batchData;
    
    if (!subject || !body) {
      return {
        valid: false,
        message: "Subject and body are required."
      };
    }
    
    if (!recipients || !Array.isArray(recipients) || recipients.length === 0) {
      return {
        valid: false,
        message: "No valid recipients found."
      };
    }
    
    // Check remaining quota
    const quotaInfo = EMAILMANAGER.getEmailQuota();
    if (recipients.length > quotaInfo.remaining) {
      return {
        valid: false,
        message: `Email quota exceeded. You have ${quotaInfo.remaining} emails remaining today but are trying to send ${recipients.length}.`
      };
    }
    
    // Validate all recipient emails
    const invalidEmails = [];
    for (const recipient of recipients) {
      if (!recipient.email || !HELPERS.isValidEmail(recipient.email)) {
        invalidEmails.push(recipient.email || "(empty)");
        if (invalidEmails.length >= 5) break; // Limit to first 5 for display
      }
    }
    
    if (invalidEmails.length > 0) {
      return {
        valid: false,
        message: `Invalid email addresses found: ${invalidEmails.join(", ")}` +
                (invalidEmails.length >= 5 ? " and others..." : "")
      };
    }
    
    return {
      valid: true,
      message: `Validated ${recipients.length} recipients.`
    };
    
  } catch (error) {
    Logger.logError("Error validating batch email", error);
    return {
      valid: false,
      message: "Error validating email batch: " + error.toString()
    };
  }
};

/**
 * Sends batch emails to multiple recipients.
 * @param {Object} batchData - Batch email data.
 * @return {Object} Result of batch operation.
 */
EMAILMANAGER.sendBatchEmails = function(batchData) {
  try {
    const config = CONFIG.getConfig();
    const { subject, body, recipients, cc, bcc, addTracking, delaySeconds } = batchData;
    
    // Validate before sending
    const validation = EMAILMANAGER.validateBatchEmail(batchData);
    if (!validation.valid) {
      return {
        success: false,
        message: validation.message
      };
    }
    
    // Initialize counters
    let successCount = 0;
    let errorCount = 0;
    const errors = [];
    
    // Get email delay setting
    const delay = delaySeconds || config.emailDelay || 1;
    
    // Process and send emails
    for (const recipient of recipients) {
      try {
        // Process the email for this recipient
        const processedEmail = COMPOSER.processEmail({
          subject,
          body,
          recipient,
          addTracking: addTracking || false
        });
        
        if (!processedEmail.success) {
          throw new Error(processedEmail.message);
        }
        
        // Send the email
        GmailApp.sendEmail(
          recipient.email,
          processedEmail.subject,
          "This email contains HTML content that cannot be displayed in plain text.",
          {
            name: config.senderName || "Mass Email Tool",
            htmlBody: processedEmail.body,
            cc: cc || "",
            bcc: bcc || ""
          }
        );
        
        // Log successful email
        Logger.logEmail({
          type: "batch",
          recipient: recipient.email,
          subject: processedEmail.subject,
          status: "sent",
          trackingId: recipient.trackingId
        });
        
        // Update recipient status if from a sheet
        if (recipient.source === "sheet" && recipient.sourceDetail) {
          RECIPIENTS.updateRecipientStatus(
            recipient.sourceDetail, 
            recipient.email,
            "sent"
          );
        }
        
        successCount++;
        
        // Wait between emails to respect quotas
        if (delay > 0 && recipient !== recipients[recipients.length - 1]) {
          HELPERS.wait(delay);
        }
        
      } catch (error) {
        // Log email error
        Logger.logError(`Failed to send email to ${recipient.email}`, error);
        Logger.logEmail({
          type: "batch",
          recipient: recipient.email,
          subject: subject,
          status: "error",
          error: error.toString()
        });
        
        errors.push(`${recipient.email}: ${error.toString()}`);
        errorCount++;
      }
    }
    
    // Update analytics
    ANALYTICS.recordEmailBatch({
      batchSize: recipients.length,
      successCount,
      errorCount,
      timestamp: new Date()
    });
    
    return {
      success: true,
      message: `Batch email processed: ${successCount} sent, ${errorCount} failed.`,
      details: {
        total: recipients.length,
        successCount,
        errorCount,
        errors: errors.slice(0, 5) // Return first 5 errors only
      }
    };
    
  } catch (error) {
    Logger.logError("Error sending batch emails", error);
    return {
      success: false,
      message: "Failed to process email batch: " + error.toString()
    };
  }
};

/**
 * Records an email open event.
 * @param {String} trackingId - The tracking ID of the opened email.
 * @return {Boolean} Success status.
 */
EMAILMANAGER.recordEmailOpen = function(trackingId) {
  try {
    if (!trackingId) return false;
    
    // Log the open event
    Logger.logEmailOpen({
      trackingId: trackingId,
      timestamp: new Date()
    });
    
    // Update analytics
    ANALYTICS.recordEmailOpen(trackingId);
    
    return true;
  } catch (error) {
    Logger.logError(`Error recording email open: ${trackingId}`, error);
    return false;
  }
};

/**
 * Gets the current email quota status.
 * @return {Object} Quota information.
 */
EMAILMANAGER.getEmailQuota = function() {
  try {
    // Get remaining daily quota from Gmail
    const emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    const config = CONFIG.getConfig();
    const maxQuota = config.maxEmailsPerDay || EMAILMANAGER.DEFAULT_EMAIL_QUOTA;
    
    return {
      total: maxQuota,
      used: Math.max(0, maxQuota - emailQuotaRemaining),
      remaining: emailQuotaRemaining,
      percentage: Math.round((emailQuotaRemaining / maxQuota) * 100)
    };
  } catch (error) {
    Logger.logError("Error getting email quota", error);
    // Return fallback values
    return {
      total: EMAILMANAGER.DEFAULT_EMAIL_QUOTA,
      used: 0,
      remaining: EMAILMANAGER.DEFAULT_EMAIL_QUOTA,
      percentage: 100
    };
  }
};

/**
 * Handles tracking pixel requests to record email opens.
 * @param {Object} request - Web request with tracking parameters.
 * @return {Object} Response for the pixel tracking.
 */
EMAILMANAGER.handleTrackingPixel = function(request) {
  try {
    // Get tracking ID from request parameters
    const trackingId = request.parameter.id;
    
    if (!trackingId) {
      throw new Error("No tracking ID provided");
    }
    
    // Record the email open
    EMAILMANAGER.recordEmailOpen(trackingId);
    
    // Return a transparent 1x1 GIF
    return ContentService
      .createTextOutput("")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
      
  } catch (error) {
    Logger.logError("Error handling tracking pixel", error);
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
};
