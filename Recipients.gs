/**
 * Manages recipient data operations for the mass email tool.
 * Handles loading, validation, and management of email recipients.
 */

// Create namespace for Recipients
const RECIPIENTS = {};

// Recipient sheet data structure constants
RECIPIENTS.HEADERS = {
  FIRST_NAME: "First Name",
  LAST_NAME: "Last Name",
  EMAIL: "Email Address",
  TAGS: "Tags",
  STATUS: "Status",
  LAST_CONTACT: "Last Contact"
};

/**
 * Gets available recipient sheets that match date format pattern.
 * @return {Array} Array of sheet names.
 */
RECIPIENTS.getAvailableRecipientSheets = function() {
  const sheets = HELPERS.getSheetNames();
  // Filter sheets to find those that match MM/DD/YYYY pattern
  const datePattern = /^(0[1-9]|1[0-2])\/(0[1-9]|[12][0-9]|3[01])\/\d{4}$/;
  
  // Collect not only date-format sheets but also sheets named with "Recipients" prefix
  return sheets.filter(name => datePattern.test(name) || name.startsWith("Recipients"));
};

/**
 * Validates a recipient list from a sheet.
 * @param {String} sheetName - The name of the sheet containing recipients.
 * @return {Object} Validation results.
 */
RECIPIENTS.validateRecipientSheet = function(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { 
        valid: false, 
        message: `Sheet "${sheetName}" not found.` 
      };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { 
        valid: false, 
        message: "Sheet does not contain recipient data." 
      };
    }
    
    const headers = data[0];
    
    // Check required headers
    const requiredHeaders = [
      RECIPIENTS.HEADERS.FIRST_NAME,
      RECIPIENTS.HEADERS.LAST_NAME,
      RECIPIENTS.HEADERS.EMAIL
    ];
    
    for (const reqHeader of requiredHeaders) {
      if (!headers.includes(reqHeader)) {
        return { 
          valid: false, 
          message: `Required header "${reqHeader}" not found in sheet.` 
        };
      }
    }
    
    // Validate email addresses
    const emailIndex = headers.indexOf(RECIPIENTS.HEADERS.EMAIL);
    const invalidEmails = [];
    
    for (let i = 1; i < data.length; i++) {
      const email = data[i][emailIndex];
      if (email && !HELPERS.isValidEmail(email)) {
        invalidEmails.push(`Row ${i+1}: ${email}`);
        if (invalidEmails.length >= 5) break; // Limit to first 5 invalid emails
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
      message: `Sheet validated successfully. Found ${data.length - 1} recipients.`,
      count: data.length - 1
    };
    
  } catch (error) {
    Logger.logError(`Error validating recipient sheet: ${sheetName}`, error);
    return { 
      valid: false, 
      message: `Error validating sheet: ${error.toString()}` 
    };
  }
};

/**
 * Loads recipients from a sheet.
 * @param {String} sheetName - The name of the sheet containing recipients.
 * @return {Array} Array of recipient objects.
 */
RECIPIENTS.loadFromSheet = function(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // No data beyond headers
    }
    
    const headers = data[0];
    const firstNameIndex = headers.indexOf(RECIPIENTS.HEADERS.FIRST_NAME);
    const lastNameIndex = headers.indexOf(RECIPIENTS.HEADERS.LAST_NAME);
    const emailIndex = headers.indexOf(RECIPIENTS.HEADERS.EMAIL);
    const tagsIndex = headers.indexOf(RECIPIENTS.HEADERS.TAGS);
    
    // Process rows into recipient objects
    const recipients = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = row[emailIndex];
      
      // Skip rows with empty or invalid emails
      if (!email || !HELPERS.isValidEmail(email)) continue;
      
      recipients.push({
        firstName: row[firstNameIndex] || "",
        lastName: row[lastNameIndex] || "",
        email: email,
        tags: tagsIndex >= 0 ? (row[tagsIndex] || "") : "",
        source: "sheet",
        sourceDetail: sheetName,
        id: HELPERS.generateUniqueId()
      });
    }
    
    return recipients;
    
  } catch (error) {
    Logger.logError(`Error loading recipients from sheet: ${sheetName}`, error);
    throw new Error(`Failed to load recipients: ${error.toString()}`);
  }
};

/**
 * Validates a manually added recipient.
 * @param {Object} recipient - The recipient to validate.
 * @return {Object} Validation results.
 */
RECIPIENTS.validateManualRecipient = function(recipient) {
  const errors = [];
  
  if (!recipient.firstName || recipient.firstName.trim() === "") {
    errors.push("First name is required");
  }
  
  if (!recipient.email || recipient.email.trim() === "") {
    errors.push("Email address is required");
  } else if (!HELPERS.isValidEmail(recipient.email)) {
    errors.push("Email address is invalid");
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
};

/**
 * Adds tracking and status information to recipients.
 * @param {Array} recipients - Array of recipient objects.
 * @return {Array} Enhanced recipient objects.
 */
RECIPIENTS.enhanceRecipients = function(recipients) {
  return recipients.map(recipient => {
    // Ensure recipient has an ID
    if (!recipient.id) {
      recipient.id = HELPERS.generateUniqueId();
    }
    
    // Add processing status fields
    recipient.status = "pending";
    recipient.trackingId = HELPERS.generateUniqueId();
    
    return recipient;
  });
};

/**
 * Creates a recipient template batch sheet.
 * @return {String} The URL of the created sheet.
 */
RECIPIENTS.createTemplateSheet = function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const sheetName = HELPERS.formatDate(today);
  
  // Check if sheet already exists with today's date
  let sheet = ss.getSheetByName(sheetName);
  let counter = 1;
  
  // If sheet exists, add counter to name
  while (sheet) {
    const newName = `${sheetName}_${counter}`;
    sheet = ss.getSheetByName(newName);
    counter++;
  }
  
  // Create new sheet with unique name
  const finalName = counter > 1 ? `${sheetName}_${counter-1}` : sheetName;
  sheet = ss.insertSheet(finalName);
  
  // Set up headers
  const headers = [
    RECIPIENTS.HEADERS.FIRST_NAME,
    RECIPIENTS.HEADERS.LAST_NAME,
    RECIPIENTS.HEADERS.EMAIL,
    RECIPIENTS.HEADERS.TAGS,
    RECIPIENTS.HEADERS.STATUS,
    RECIPIENTS.HEADERS.LAST_CONTACT
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
  // Add sample data
  const sampleData = [
    ["John", "Doe", "john.doe@example.com", "customer", "", ""],
    ["Jane", "Smith", "jane.smith@example.com", "prospect", "", ""]
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length)
    .setValues(sampleData);
  
  // Format the sheet
  sheet.autoResizeColumns(1, headers.length);
  
  return ss.getUrl() + "#gid=" + sheet.getSheetId();
};

/**
 * Updates recipient status in sheet.
 * @param {String} sheetName - Name of the sheet.
 * @param {String} email - Email to update.
 * @param {String} status - New status value.
 * @return {Boolean} Success status.
 */
RECIPIENTS.updateRecipientStatus = function(sheetName, email, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailIndex = headers.indexOf(RECIPIENTS.HEADERS.EMAIL);
    const statusIndex = headers.indexOf(RECIPIENTS.HEADERS.STATUS);
    const lastContactIndex = headers.indexOf(RECIPIENTS.HEADERS.LAST_CONTACT);
    
    if (emailIndex === -1 || statusIndex === -1) return false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === email) {
        // Update status
        sheet.getRange(i + 1, statusIndex + 1).setValue(status);
        
        // Update last contact if that column exists
        if (lastContactIndex !== -1) {
          sheet.getRange(i + 1, lastContactIndex + 1).setValue(new Date());
        }
        
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    Logger.logError(`Error updating recipient status: ${email}`, error);
    return false;
  }
};
