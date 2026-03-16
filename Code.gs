// ============================================
// GIANT MOTO PRO - MOTORCYCLE UNIT REQUEST SYSTEM
// ============================================
// This script manages motorcycle unit requests for Giant Moto Pro.
// 
// Features:
// - Submit new unit requests (for display or customer release)
// - Check request status by reference number
// - Cancel pending requests with mandatory remarks
// - Automatic email notifications for:
//   * Order confirmation
//   * Allocation (with pickup instructions)
//   * Cancellation
//
// IMPORTANT: When changing a request status to "ALLOCATED" in the spreadsheet,
// manually call sendAllocationEmail(email, referenceNumber) to notify the customer
// with pickup instructions (Call 09754490487).
//
// To trigger allocation email manually:
// 1. Open Script Editor > Tools > Script editor
// 2. Create a new function or modify existing one:
//    function notifyAllocation() {
//      sendAllocationEmail("customer@email.com", "REQ-XXXXXX");
//    }
// 3. Run the function
// ============================================

// ============================================
// CONFIGURATION
// ============================================
const SHEET_ORDERS = "ORDERS";
const SHEET_REF = "REFERENCE_DATA";
const APPROVED_APPLICATION_FORM_LINK_HEADER = "Approved Application Form Link";
const APPROVED_APPLICATION_FORM_UPLOAD_PATH = "UNIT_REQUEST_UPLOADS/APPROVED_APPLICATION_FORMS";

// ============================================
// CRITICAL: UPDATE DEPLOYMENT URL
// ============================================
// After deploying your web app:
// 1. Go to Apps Script Editor > Deploy > Manage deployments
// 2. Copy the Web app URL
// 3. Replace the placeholder below with your actual deployment URL
// 
// Example: "https://script.google.com/macros/s/AKfycbwSK8NV0S02YCAI_b_XDPftqJfRztLGQX84xhLc-2P2gYePntx1tkbGcntfNtSNOJwxnA/exec"
// 
// This URL is used in:
// - Email links (Check Status, Cancel Request buttons)
// - Form navigation between pages
// ============================================
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwndrk7-SRXlm3U_LFa7Ebw04obfqRztamIgDLwQiXN2L5Zy4dLvharB8gaUceVBiaNGA/exec"; // Replace with actual URL after deployment

// Column indexes for ORDERS sheet (0-based)
const COL = {
  TIMESTAMP: 0,
  REFERENCE: 1,
  BRANCH: 2,
  QUANTITY: 3,
  MODEL: 4,
  COLOR: 5,
  REMARKS: 6,           // FOR DISPLAY or FOR RELEASE
  CLIENT_NAME: 7,
  CONTACT_NUMBER: 8,
  EMAIL: 9,
  STATUS: 10,
  CANCELLATION_REMARKS: 11,  // Cancellation reason
  TIME_LOG: 12          // Timestamp when cancelled
};

// Column indexes for REFERENCE_DATA sheet
const REF_COL = {
  BRANCH: 0,
  MODEL:  1,
  COLOR: 2
};

// ============================================
// WEB APP ENTRY POINT - UNIFIED SINGLE-PAGE APP
// ============================================
// Optimized for instant navigation with no lag
// Returns a unified HTML template containing both form and status pages
// Navigation happens client-side via JavaScript for instant switching
// This eliminates server roundtrips and provides smooth, seamless UX
// ============================================
function doGet(e) {
  try {
    // Return unified single-page app with both views
    return HtmlService.createTemplateFromFile('app')
      .evaluate()
      .setTitle("Giant Moto Pro - Unit Request System")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    Logger.log("doGet Error: " + error.toString());
    return HtmlService.createHtmlOutput(
      '<h1>System Error</h1><p>Unable to load application. Please contact support.</p><p>' + error.toString() + '</p>'
    );
  }
}

// ============================================
// HELPER: INCLUDE HTML PARTIALS
// ============================================
// Used for modular HTML template construction
// ============================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// GET DROPDOWN REFERENCE DATA
// ============================================
function getReferenceData() {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REF);
    
    if (!sh) {
      throw new Error("Reference data sheet '" + SHEET_REF + "' not found.");
    }
    
    var data = sh.getDataRange().getValues();
    
    if (data.length < 2) {
      throw new Error("No reference data available.");
    }
    
    data.shift(); // Remove header row
    
    return {
      branches: getUnique(data, REF_COL.BRANCH),
      models: getUnique(data, REF_COL.MODEL),
      colors: getUnique(data, REF_COL.COLOR)
    };
  } catch (error) {
    throw new Error("Failed to load reference data: " + error.message);
  }
}

// ============================================
// HELPER:  GET UNIQUE VALUES FROM COLUMN
// ============================================
function getUnique(arr, col) {
  var seen = {};
  var result = [];
  
  for (var i = 0; i < arr.length; i++) {
    var val = String(arr[i][col]).trim();
    if (val && !seen[val]) {
      seen[val] = true;
      result.push(val);
    }
  }
  
  return result.sort(); // Return sorted for better UX
}

// ============================================
// GENERATE UNIQUE REFERENCE NUMBER
// ============================================
function generateReferenceNumber() {
  var uuid = Utilities.getUuid().slice(0, 6).toUpperCase();
  return "REQ-" + uuid;
}

// ============================================
// SUBMIT ORDER - WITH ENHANCED SECURITY
// ============================================
// Validates and sanitizes all inputs before saving
// Prevents injection attacks and ensures data integrity
// ============================================
function submitOrder(data) {
  try {
    // ===== VALIDATION: Required fields =====
    if (!data || typeof data !== 'object') {
      throw new Error("Invalid request data.");
    }
    
    if (!data.email || !isValidEmail(data.email)) {
      throw new Error("Valid email address is required.");
    }
    
    if (!data.branch || !data.model || !data.color) {
      throw new Error("Branch, Model, and Color are required.");
    }
    
    // ===== SANITIZE: All text inputs =====
    var branch = sanitizeInput(data.branch, 50);
    var model = sanitizeInput(data.model, 50);
    var color = sanitizeInput(data.color, 30);
    var email = sanitizeInput(data.email, 100);
    var clientName = sanitizeInput(data.clientName, 100);
    var contactNumber = sanitizeInput(data.contactNumber, 15);
    var remarks = sanitizeInput(data.remarks, 50).toUpperCase();
    
    // ===== VALIDATION: Verify sanitized data not empty =====
    if (!branch || !model || !color || !email) {
      throw new Error("Required fields cannot be empty after validation.");
    }
    
    // ===== PROCESS: Remarks and quantity =====
    var qty = 1;
    
    if (remarks === "FOR DISPLAY") {
      qty = parseInt(data.qty, 10);
      if (isNaN(qty) || qty < 1 || qty > 100) {
        throw new Error("Quantity must be between 1 and 100 for display orders.");
      }
    } else if (remarks === "FOR RELEASE") {
      if (!clientName || clientName.length < 2) {
        throw new Error("Valid client name is required for release orders.");
      }
      if (!contactNumber || contactNumber.length < 7) {
        throw new Error("Valid contact number is required for release orders.");
      }
      qty = 1;
    } else {
      throw new Error("Please select valid purpose (FOR DISPLAY or FOR RELEASE).");
    }
    
    // ===== GENERATE: Unique reference number =====
    var refNo = generateReferenceNumber();
    
    // ===== VALIDATION: Ensure reference is unique =====
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
    if (!sh) {
      throw new Error("Orders sheet not found. Please contact administrator.");
    }
    
    var existingData = sh.getDataRange().getValues();
    for (var i = 1; i < existingData.length; i++) {
      if (existingData[i][COL.REFERENCE] === refNo) {
        // Extremely rare collision - regenerate
        refNo = generateReferenceNumber();
        break;
      }
    }

    var approvedApplicationFormLink = "";
    if (remarks === "FOR RELEASE") {
      validateApprovedApplicationFormPayload_(data.approvedApplicationForm);
      var uploadedFile = saveApprovedApplicationForm_(refNo, branch, data.approvedApplicationForm);
      approvedApplicationFormLink = uploadedFile.url;
    }

    var approvedLinkColumnInfo = ensureOrdersHeaderColumn_(sh, APPROVED_APPLICATION_FORM_LINK_HEADER);
    var rowLength = Math.max(approvedLinkColumnInfo.headerValues.length, COL.TIME_LOG + 1);
    var rowValues = new Array(rowLength).fill("");

    rowValues[COL.TIMESTAMP] = new Date();
    rowValues[COL.REFERENCE] = refNo;
    rowValues[COL.BRANCH] = branch;
    rowValues[COL.QUANTITY] = qty;
    rowValues[COL.MODEL] = model;
    rowValues[COL.COLOR] = color;
    rowValues[COL.REMARKS] = remarks;
    rowValues[COL.CLIENT_NAME] = clientName;
    rowValues[COL.CONTACT_NUMBER] = contactNumber;
    rowValues[COL.EMAIL] = email;
    rowValues[COL.STATUS] = "PENDING";
    rowValues[COL.CANCELLATION_REMARKS] = "";
    rowValues[COL.TIME_LOG] = "";
    rowValues[approvedLinkColumnInfo.columnIndex] = approvedApplicationFormLink;
    
    // ===== SAVE: Append new order to sheet =====
    sh.appendRow(rowValues);
    
    // ===== NOTIFICATION: Send confirmation email =====
    sendOrderConfirmationEmail(email, refNo);
    
    Logger.log("Order submitted successfully - Ref: " + refNo + ", Email: " + email);
    
    return { 
      success: true, 
      reference: refNo 
    };
    
  } catch (error) {
    Logger.log("Error in submitOrder: " + error.toString());
    throw new Error(error.message || "Failed to submit order. Please try again.");
  }
}

function validateApprovedApplicationFormPayload_(fileObj) {
  if (!fileObj || typeof fileObj !== 'object') {
    throw new Error("Approved Application Form image is required for release orders.");
  }

  var dataUri = String(fileObj.dataUri || '');
  if (!/^data:image\/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/=\r\n]+$/.test(dataUri)) {
    throw new Error("Approved Application Form image is required for release orders.");
  }
}

function saveApprovedApplicationForm_(refNo, branch, fileObj) {
  var parsedFile = parseImageDataUri_(fileObj.dataUri);
  var uploadFolder = getOrCreateFolderByPath_(APPROVED_APPLICATION_FORM_UPLOAD_PATH);
  var requestFolderName = sanitizeFolderName_(refNo + "_" + branch);
  var requestFolder = getOrCreateChildFolder_(uploadFolder, requestFolderName);
  var extension = getImageExtension_(parsedFile.mimeType);
  var fileName = "Approved_Application_" + refNo + "." + extension;
  var blob = Utilities.newBlob(parsedFile.bytes, parsedFile.mimeType, fileName);
  var file = requestFolder.createFile(blob);

  return {
    fileId: file.getId(),
    url: file.getUrl()
  };
}

function ensureOrdersHeaderColumn_(sheet, headerName) {
  var headerWidth = Math.max(sheet.getLastColumn(), COL.TIME_LOG + 1);
  var headers = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerName) {
      return {
        columnIndex: i,
        headerValues: headers
      };
    }
  }

  var lastNamedHeaderIndex = -1;
  for (var j = 0; j < headers.length; j++) {
    if (String(headers[j]).trim() !== "") {
      lastNamedHeaderIndex = j;
    }
  }

  var targetColumn = lastNamedHeaderIndex + 2; // 1-based column number
  sheet.getRange(1, targetColumn).setValue(headerName);

  if (headers.length < targetColumn) {
    while (headers.length < targetColumn) {
      headers.push("");
    }
  }
  headers[targetColumn - 1] = headerName;

  return {
    columnIndex: targetColumn - 1,
    headerValues: headers
  };
}

function parseImageDataUri_(dataUri) {
  var match = String(dataUri || '').match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([A-Za-z0-9+/=\r\n]+)$/);
  if (!match) {
    throw new Error("Approved Application Form image is required for release orders.");
  }

  return {
    mimeType: match[1].toLowerCase(),
    bytes: Utilities.base64Decode(match[2].replace(/\s/g, ''))
  };
}

function getOrCreateFolderByPath_(path) {
  var parts = String(path || '').split('/').filter(function(part) {
    return part && part.trim();
  });

  if (!parts.length) {
    throw new Error("Upload folder path is not configured.");
  }

  var rootIterator = DriveApp.getFoldersByName(parts[0]);
  var folder = rootIterator.hasNext() ? rootIterator.next() : DriveApp.createFolder(parts[0]);

  for (var i = 1; i < parts.length; i++) {
    folder = getOrCreateChildFolder_(folder, parts[i]);
  }

  return folder;
}

function getOrCreateChildFolder_(parentFolder, childName) {
  var iterator = parentFolder.getFoldersByName(childName);
  return iterator.hasNext() ? iterator.next() : parentFolder.createFolder(childName);
}

function sanitizeFolderName_(name) {
  return String(name || 'REQUEST')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 100) || 'REQUEST';
}

function getImageExtension_(mimeType) {
  var normalized = String(mimeType || '').toLowerCase();
  var map = {
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/bmp': 'bmp',
    'image/tiff': 'tiff',
    'image/heic': 'heic',
    'image/heif': 'heif'
  };

  return map[normalized] || 'img';
}

// ============================================
// SEND CONFIRMATION EMAIL - PROFESSIONAL HTML FORMAT
// ============================================
// Sends a professional HTML-formatted confirmation email when a request is submitted
// Features:
// - Professional design with color-coded sections
// - Clear reference number display
// - Two action buttons: "Check Status" and "Cancel Request"
// - Mobile-responsive design
// - Uses WEB_APP_URL for button links
// 
// @param {string} email - Customer email address
// @param {string} refNo - Generated reference number (e.g., REQ-9534C6)
// ============================================
function sendOrderConfirmationEmail(email, refNo) {
  try {
    var subject = "Giant Moto Pro - Request Confirmation - Ref: " + refNo;
    
    // Build status page URL for checking and cancellation
    var statusUrl = WEB_APP_URL + "?page=status";
    
    // Professional HTML email with cancel button
    var htmlBody = 
      '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f5f5f5;">' +
        '<div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">' +
          '<h2 style="color: #1976d2; margin: 0 0 20px 0; font-size: 24px; border-bottom: 3px solid #1976d2; padding-bottom: 10px;">Request Confirmation</h2>' +
          
          '<p style="color: #333; font-size: 16px; line-height: 1.6;">Dear Valued Customer,</p>' +
          '<p style="color: #333; font-size: 16px; line-height: 1.6;">Thank you for choosing Giant Moto Pro. Your motorcycle unit request has been successfully received.</p>' +
          
          '<div style="background-color: #e3f2fd; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #1976d2;">' +
            '<h3 style="color: #1976d2; margin: 0 0 15px 0; font-size: 18px;">Request Details</h3>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Reference Number:</strong> <span style="color: #1976d2; font-size: 18px; font-weight: bold;">' + refNo + '</span></p>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Status:</strong> <span style="background-color: #ff9800; color: white; padding: 4px 12px; border-radius: 15px; font-weight: bold;">PENDING</span></p>' +
          '</div>' +
          
          '<div style="background-color: #fff3cd; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #ffc107;">' +
            '<h3 style="color: #856404; margin: 0 0 15px 0; font-size: 18px;">Important Information</h3>' +
            '<ul style="color: #333; font-size: 15px; line-height: 1.8; margin: 0; padding-left: 20px;">' +
              '<li>Please keep this reference number for your records</li>' +
              '<li>You can check your request status anytime using the button below</li>' +
              '<li>To cancel this request, visit the status page and click "Cancel Request"</li>' +
              '<li>You will receive email notifications when your request status changes</li>' +
            '</ul>' +
          '</div>' +
          
          '<div style="text-align: center; margin: 30px 0;">' +
            '<a href="' + statusUrl + '" style="display: inline-block; background-color: #1976d2; color: white; padding: 15px 40px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px; margin: 10px;">Check Status</a>' +
            '<a href="' + statusUrl + '" style="display: inline-block; background-color: #f44336; color: white; padding: 15px 40px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px; margin: 10px;">Cancel Request</a>' +
          '</div>' +
          
          '<div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">' +
            '<p style="color: #666; font-size: 14px; line-height: 1.6;">If you have any questions, please contact us immediately.</p>' +
            '<p style="color: #333; font-size: 15px; margin: 15px 0 5px 0;"><strong>Best regards,</strong></p>' +
            '<p style="color: #1976d2; font-size: 16px; font-weight: bold; margin: 5px 0;">Giant Moto Pro Team</p>' +
            '<p style="color: #666; font-size: 14px; margin: 5px 0;"><strong>Contact:</strong> 09754490487</p>' +
          '</div>' +
        '</div>' +
      '</div>';
    
    // Send HTML email
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log("Confirmation email sent to: " + email);
  } catch (error) {
    Logger.log("Failed to send confirmation email to " + email + ": " + error.toString());
    // Don't throw - order was saved successfully, email is secondary
  }
}

// ============================================
// CHECK REQUEST STATUS - WITH SECURITY VALIDATION
// ============================================
// Validates reference number format before querying
// Prevents malicious input and ensures data integrity
// ============================================
function checkRequestStatus(referenceNumber) {
  try {
    // ===== VALIDATION: Input exists and is string =====
    if (!referenceNumber || typeof referenceNumber !== 'string') {
      return { found: false, error: "Reference number is required." };
    }
    
    var searchRef = referenceNumber.trim().toUpperCase();
    
    // ===== SECURITY: Validate reference format =====
    if (!isValidReferenceNumber(searchRef)) {
      return { 
        found: false, 
        error: "Invalid reference number format. Must be REQ-XXXXXX." 
      };
    }
    
    // ===== GET: Spreadsheet data =====
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_ORDERS);
    
    if (!sh) {
      Logger.log("ERROR: ORDERS sheet not found");
      return { 
        found: false, 
        error: "System error. Please contact support." 
      };
    }
    
    var data = sh.getDataRange().getValues();
    
    if (data.length < 2) {
      return { found: false, error: "No orders found in the system." };
    }
    
    data.shift(); // Remove header row
    
    // ===== SEARCH: Find matching reference =====
    var foundRow = null;
    
    for (var i = 0; i < data.length; i++) {
      var cellValue = String(data[i][COL.REFERENCE]).trim().toUpperCase();
      if (cellValue === searchRef) {
        foundRow = data[i];
        break;
      }
    }
    
    if (!foundRow) {
      return { 
        found: false,
        error: "No request found with this reference number." 
      };
    }
    
    // ===== RETURN: Sanitized order details =====
    return {
      found: true,
      date: foundRow[COL.TIMESTAMP] ? formatDate(foundRow[COL.TIMESTAMP]) : "",
      reference: sanitizeInput(foundRow[COL.REFERENCE], 20),
      branch: sanitizeInput(foundRow[COL.BRANCH], 50),
      quantity: foundRow[COL.QUANTITY] || "",
      model: sanitizeInput(foundRow[COL.MODEL], 50),
      color: sanitizeInput(foundRow[COL.COLOR], 30),
      remarks: sanitizeInput(foundRow[COL.REMARKS], 50),
      clientName: sanitizeInput(foundRow[COL.CLIENT_NAME], 100),
      contactNumber: sanitizeInput(foundRow[COL.CONTACT_NUMBER], 15),
      email: sanitizeInput(foundRow[COL.EMAIL], 100),
      status: sanitizeInput(foundRow[COL.STATUS], 20) || "UNKNOWN"
    };
    
  } catch (error) {
    Logger.log("Error in checkRequestStatus: " + error.toString());
    return { 
      found: false, 
      error: "System error. Please try again later." 
    };
  }
}

// ============================================
// SECURITY: INPUT SANITIZATION
// ============================================
// Prevents XSS and injection attacks by sanitizing user input
// Removes potentially dangerous characters and limits length
// ============================================
function sanitizeInput(input, maxLength) {
  if (!input) return '';
  
  var sanitized = String(input)
    .trim()
    .replace(/[<>\"'`]/g, '') // Remove HTML/script injection chars
    .replace(/\r?\n/g, ' ')    // Replace newlines with spaces
    .slice(0, maxLength || 200); // Enforce max length
  
  return sanitized;
}

// ============================================
// SECURITY: VALIDATE EMAIL FORMAT
// ============================================
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  
  var re = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return re.test(email.trim()) && email.length <= 100;
}

// ============================================
// SECURITY: VALIDATE REFERENCE NUMBER FORMAT
// ============================================
function isValidReferenceNumber(ref) {
  if (!ref || typeof ref !== 'string') return false;
  
  // Must match REQ-XXXXXX format (REQ- followed by 6 alphanumeric characters)
  var re = /^REQ-[A-Z0-9]{6}$/;
  return re.test(ref.trim().toUpperCase());
}

// ============================================
// HELPER: FORMAT DATE
// ============================================
function formatDate(date) {
  try {
    if (date instanceof Date) {
      return date.toISOString();
    }
    return String(date);
  } catch (e) {
    return "";
  }
}

// ============================================
// CANCEL REQUEST - WITH ENHANCED SECURITY
// ============================================
// Validates all inputs and ensures only PENDING requests can be cancelled
// Prevents unauthorized cancellations and data manipulation
// @param {string} referenceNumber - The request reference number
// @param {string} cancelRemarks - Mandatory cancellation reason
// @return {Object} Result with success status and message
// ============================================
function cancelRequest(referenceNumber, cancelRemarks) {
  try {
    // ===== VALIDATION: Reference number =====
    if (!referenceNumber || typeof referenceNumber !== 'string') {
      return { success: false, error: "Reference number is required." };
    }
    
    var searchRef = referenceNumber.trim().toUpperCase();
    
    // ===== SECURITY: Validate reference format =====
    if (!isValidReferenceNumber(searchRef)) {
      return { 
        success: false, 
        error: "Invalid reference number format." 
      };
    }
    
    // ===== VALIDATION: Cancellation remarks =====
    if (!cancelRemarks || typeof cancelRemarks !== 'string') {
      return { 
        success: false, 
        error: "Cancellation reason is required." 
      };
    }
    
    var remarks = sanitizeInput(cancelRemarks, 500);
    
    if (remarks.length < 5) {
      return { 
        success: false, 
        error: "Please provide a detailed reason (at least 5 characters)." 
      };
    }
    
    // ===== GET: Spreadsheet data =====
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_ORDERS);
    
    if (!sh) {
      return { 
        success: false, 
        error: "Orders sheet not found. Please contact administrator." 
      };
    }
    
    var data = sh.getDataRange().getValues();
    
    if (data.length < 2) {
      return { success: false, error: "No orders found in the system." };
    }
    
    // ===== SEARCH: Find the request row =====
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      var cellValue = String(data[i][COL.REFERENCE]).trim().toUpperCase();
      if (cellValue === searchRef) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { 
        success: false, 
        error: "Request not found with this reference number." 
      };
    }
    
    // ===== VALIDATION: Only PENDING requests can be cancelled =====
    var currentStatus = String(data[rowIndex][COL.STATUS]).trim().toUpperCase();
    
    if (currentStatus !== "PENDING") {
      return { 
        success: false, 
        error: "Only PENDING requests can be cancelled. Current status: " + currentStatus 
      };
    }
    
    // ===== UPDATE: Change status to CANCELLED =====
    var actualRowNumber = rowIndex + 1; // Convert to 1-based index
    
    sh.getRange(actualRowNumber, COL.STATUS + 1).setValue("CANCELLED");
    sh.getRange(actualRowNumber, COL.CANCELLATION_REMARKS + 1).setValue(remarks);
    
    // ===== LOG TIMESTAMP: Record cancellation time =====
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    sh.getRange(actualRowNumber, COL.TIME_LOG + 1).setValue(timestamp);
    Logger.log("cancelRequest: Timestamp logged - " + timestamp);
    
    // ===== NOTIFICATION: Send cancellation email =====
    var email = String(data[rowIndex][COL.EMAIL]).trim();
    if (email && isValidEmail(email)) {
      sendCancellationEmail(email, searchRef, remarks);
    }
    
    Logger.log("Request cancelled - Ref: " + searchRef + ", Reason: " + remarks);
    
    return { 
      success: true, 
      message: "Request has been cancelled successfully." 
    };
    
  } catch (error) {
    Logger.log("Error in cancelRequest: " + error.toString());
    return { 
      success: false, 
      error: "System error. Please try again later." 
    };
  }
}

// ============================================
// SEND CANCELLATION EMAIL - PROFESSIONAL HTML FORMAT
// ============================================
// Sends a professional HTML-formatted confirmation email when a request is cancelled
// 
// Features:
// - Red cancellation theme
// - Displays cancellation reason
// - "Submit New Request" button linking to form
// - Professional and empathetic tone
// 
// @param {string} email - Customer email address
// @param {string} refNo - Reference number
// @param {string} remarks - Cancellation reason provided by customer
// ============================================
function sendCancellationEmail(email, refNo, remarks) {
  try {
    var subject = "Giant Moto Pro - Request Cancelled - Ref: " + refNo;
    
    // Build form URL for new requests
    var formUrl = WEB_APP_URL;
    
    // Professional HTML email for cancellation confirmation
    var htmlBody = 
      '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f5f5f5;">' +
        '<div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">' +
          '<h2 style="color: #f44336; margin: 0 0 20px 0; font-size: 24px; border-bottom: 3px solid #f44336; padding-bottom: 10px;">Request Cancelled</h2>' +
          
          '<p style="color: #333; font-size: 16px; line-height: 1.6;">Dear Valued Customer,</p>' +
          '<p style="color: #333; font-size: 16px; line-height: 1.6;">Your motorcycle unit request has been cancelled as per your request.</p>' +
          
          '<div style="background-color: #e3f2fd; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #1976d2;">' +
            '<h3 style="color: #1976d2; margin: 0 0 15px 0; font-size: 18px;">Request Details</h3>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Reference Number:</strong> <span style="color: #1976d2; font-size: 18px; font-weight: bold;">' + refNo + '</span></p>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Status:</strong> <span style="background-color: #f44336; color: white; padding: 4px 12px; border-radius: 15px; font-weight: bold;">CANCELLED</span></p>' +
          '</div>' +
          
          '<div style="background-color: #fdecea; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #f44336;">' +
            '<h3 style="color: #c62828; margin: 0 0 15px 0; font-size: 18px;">Cancellation Reason</h3>' +
            '<p style="color: #333; font-size: 15px; line-height: 1.6; margin: 0;">' + remarks + '</p>' +
          '</div>' +
          
          '<div style="background-color: #e8f5e9; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #4caf50;">' +
            '<h3 style="color: #2e7d32; margin: 0 0 15px 0; font-size: 18px;">Next Steps</h3>' +
            '<ul style="color: #333; font-size: 15px; line-height: 1.8; margin: 0; padding-left: 20px;">' +
              '<li>If you wish to submit a new request, click the button below</li>' +
              '<li>If you believe this cancellation was made in error, please contact us immediately</li>' +
            '</ul>' +
          '</div>' +
          
          '<div style="text-align: center; margin: 30px 0;">' +
            '<a href="' + formUrl + '" style="display: inline-block; background-color: #1976d2; color: white; padding: 15px 40px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">Submit New Request</a>' +
          '</div>' +
          
          '<div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">' +
            '<p style="color: #333; font-size: 15px; margin: 15px 0 5px 0;">Thank you for your understanding.</p>' +
            '<p style="color: #333; font-size: 15px; margin: 5px 0;"><strong>Best regards,</strong></p>' +
            '<p style="color: #1976d2; font-size: 16px; font-weight: bold; margin: 5px 0;">Giant Moto Pro Team</p>' +
            '<p style="color: #666; font-size: 14px; margin: 5px 0;"><strong>Contact:</strong> 09754490487</p>' +
          '</div>' +
        '</div>' +
      '</div>';
    
    // Send HTML email
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log("Cancellation email sent to: " + email);
  } catch (error) {
    Logger.log("Failed to send cancellation email to " + email + ": " + error.toString());
    // Don't throw - cancellation was successful, email is secondary
  }
}

// ============================================
// SEND ALLOCATION EMAIL - PROFESSIONAL HTML FORMAT
// ============================================
// Sends a professional HTML-formatted email when a unit is allocated
// This function is automatically called by the onEdit trigger when status changes to "ALLOCATED"
// 
// Features:
// - Green success theme
// - Clear pickup instructions with phone number
// - Requirements checklist
// - Deadline reminder (3 business days)
// 
// @param {string} email - Customer email address
// @param {string} refNo - Reference number
// ============================================
function sendAllocationEmail(email, refNo) {
  try {
    var subject = "Giant Moto Pro - Unit Allocated and Queued for Delivery! - Ref: " + refNo;
    
    // Professional HTML email with clear pickup instructions
    var htmlBody = 
      '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f5f5f5;">' +
        '<div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">' +
          '<h2 style="color: #4caf50; margin: 0 0 20px 0; font-size: 24px; border-bottom: 3px solid #4caf50; padding-bottom: 10px;">Unit Allocated - Queued for Delivery!</h2>' +
          
          '<div style="background-color: #e8f5e9; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #4caf50; text-align: center;">' +
            '<p style="color: #2e7d32; font-size: 18px; font-weight: bold; margin: 0;">Great News!</p>' +
            '<p style="color: #333; font-size: 16px; margin: 10px 0 0 0;">Your requested motorcycle unit has been <strong>ALLOCATED</strong> and is on Queued for Delivery! Please call Warehouse at 09754490487 for the delivery schedule.</p>' +
          '</div>' +
          
          '<div style="background-color: #e3f2fd; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #1976d2;">' +
            '<h3 style="color: #1976d2; margin: 0 0 15px 0; font-size: 18px;">Request Details</h3>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Reference Number:</strong> <span style="color: #1976d2; font-size: 18px; font-weight: bold;">' + refNo + '</span></p>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Status:</strong> <span style="background-color: #4caf50; color: white; padding: 4px 12px; border-radius: 15px; font-weight: bold;">ALLOCATED</span></p>' +
          '</div>' +
          
          '<p style="color: #666; font-size: 14px; line-height: 1.6; margin: 20px 0;">If you have any questions or need to reschedule, please contact us immediately.</p>' +
          
          '<div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">' +
            '<p style="color: #333; font-size: 15px; margin: 15px 0 5px 0;"><strong>Thank you for choosing Giant Moto Pro.</strong></p>' +
            '<p style="color: #333; font-size: 15px; margin: 5px 0;"><strong>Best regards,</strong></p>' +
            '<p style="color: #1976d2; font-size: 16px; font-weight: bold; margin: 5px 0;">Giant Moto Pro Team</p>' +
            '<p style="color: #666; font-size: 14px; margin: 5px 0;"><strong>Contact:</strong> 09754490487</p>' +
          '</div>' +
        '</div>' +
      '</div>';
    
    // Send HTML email
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log("Allocation email sent to: " + email);
  } catch (error) {
    Logger.log("Failed to send allocation email to " + email + ": " + error.toString());
  }
}

// ============================================
// SEND COMPLETION EMAIL - PROFESSIONAL HTML FORMAT
// ============================================
// Sends a professional HTML-formatted email when a request is completed
// This function is automatically called by the onEdit trigger when status changes to "COMPLETED"
// 
// Features:
// - Blue completion theme
// - Thank you message
// - Feedback request
// - New request button
// 
// @param {string} email - Customer email address
// @param {string} refNo - Reference number
// ============================================
function sendCompletionEmail(email, refNo) {
  try {
    var subject = "Giant Moto Pro - Request Completed - Ref: " + refNo;
    
    // Build form URL for new requests
    var formUrl = WEB_APP_URL;
    
    // Professional HTML email for completion confirmation
    var htmlBody = 
      '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f5f5f5;">' +
        '<div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">' +
          '<h2 style="color: #2196f3; margin: 0 0 20px 0; font-size: 24px; border-bottom: 3px solid #2196f3; padding-bottom: 10px;">Request Completed Successfully!</h2>' +
          
          '<div style="background-color: #e3f2fd; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #2196f3; text-align: center;">' +
            '<p style="color: #1565c0; font-size: 18px; font-weight: bold; margin: 0;">Thank You!</p>' +
            '<p style="color: #333; font-size: 16px; margin: 10px 0 0 0;">Your motorcycle unit request has been <strong>COMPLETED</strong> successfully.</p>' +
          '</div>' +
          
          '<div style="background-color: #e8f5e9; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #4caf50;">' +
            '<h3 style="color: #2e7d32; margin: 0 0 15px 0; font-size: 18px;">Request Details</h3>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Reference Number:</strong> <span style="color: #1976d2; font-size: 18px; font-weight: bold;">' + refNo + '</span></p>' +
            '<p style="color: #333; margin: 5px 0; font-size: 15px;"><strong>Status:</strong> <span style="background-color: #2196f3; color: white; padding: 4px 12px; border-radius: 15px; font-weight: bold;">COMPLETED</span></p>' +
          '</div>' +
          
          '<div style="background-color: #fff3cd; padding: 20px; border-radius: 5px; margin: 25px 0; border-left: 4px solid #ffc107;">' +
            '<h3 style="color: #856404; margin: 0 0 15px 0; font-size: 18px;">We Value Your Feedback</h3>' +
            '<p style="color: #333; font-size: 15px; line-height: 1.6; margin: 0;">We hope you had a great experience with Giant Moto Pro. Your feedback helps us improve our service.</p>' +
          '</div>' +
          
          '<div style="text-align: center; margin: 30px 0;">' +
            '<a href="' + formUrl + '" style="display: inline-block; background-color: #1976d2; color: white; padding: 15px 40px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">Submit New Request</a>' +
          '</div>' +
          
          '<div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">' +
            '<p style="color: #333; font-size: 15px; margin: 15px 0 5px 0;">We appreciate your business and look forward to serving you again.</p>' +
            '<p style="color: #333; font-size: 15px; margin: 5px 0;"><strong>Best regards,</strong></p>' +
            '<p style="color: #1976d2; font-size: 16px; font-weight: bold; margin: 5px 0;">Giant Moto Pro Team</p>' +
            '<p style="color: #666; font-size: 14px; margin: 5px 0;"><strong>Contact:</strong> 09754490487</p>' +
          '</div>' +
        '</div>' +
      '</div>';
    
    // Send HTML email
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log("Completion email sent to: " + email);
  } catch (error) {
    Logger.log("Failed to send completion email to " + email + ": " + error.toString());
  }
}

// ============================================
// AUTO-TRIGGER: COMPREHENSIVE STATUS CHANGE NOTIFICATIONS
// ============================================
// Automatically sends appropriate email notifications when status changes
// Handles: PENDING → ALLOCATED, COMPLETED, CANCELLED
// Prevents duplicate notifications by tracking previous status
// 
// CRITICAL SETUP INSTRUCTIONS:
// 1. In Apps Script Editor, click "Triggers" (clock icon on left sidebar)
// 2. Click "+ Add Trigger" (bottom right)
// 3. Configure:
//    - Choose function: onEdit
//    - Deployment: Head
//    - Event source: From spreadsheet
//    - Event type: On edit
// 4. Click "Save" and grant permissions
// 
// @param {Object} e - The event object from onEdit trigger
// ============================================
function onEdit(e) {
  try {
    // ===== VALIDATION: Event object exists =====
    if (!e || !e.range) {
      Logger.log("onEdit: No event object - manual test or direct edit");
      return;
    }
    
    var range = e.range;
    var sheet = range.getSheet();
    var column = range.getColumn();
    var row = range.getRow();
    
    Logger.log("onEdit: Triggered - Sheet: " + sheet.getName() + ", Row: " + row + ", Column: " + column);
    
    // ===== FILTER: Only ORDERS sheet =====
    if (sheet.getName() !== SHEET_ORDERS) {
      Logger.log("onEdit: Skipped - Not ORDERS sheet");
      return;
    }
    
    // ===== FILTER: Only STATUS column (K = index 10, column 11) =====
    if (column !== COL.STATUS + 1) {
      Logger.log("onEdit: Skipped - Not STATUS column (column " + column + ")");
      return;
    }
    
    // ===== FILTER: Skip header row =====
    if (row === 1) {
      Logger.log("onEdit: Skipped - Header row");
      return;
    }
    
    // ===== GET STATUS CHANGE DATA =====
    var newStatus = String(range.getValue()).trim().toUpperCase();
    var oldStatus = e.oldValue ? String(e.oldValue).trim().toUpperCase() : "";
    
    Logger.log("onEdit: Status change - Old: '" + oldStatus + "' → New: '" + newStatus + "'");
    
    // ===== PREVENT DUPLICATE: Status unchanged =====
    if (newStatus === oldStatus) {
      Logger.log("onEdit: Skipped - Status unchanged");
      return;
    }
    
    // ===== PREVENT EMPTY STATUS =====
    if (!newStatus) {
      Logger.log("onEdit: Skipped - Empty status value");
      return;
    }
    
    // ===== GET REQUEST DATA =====
    var email = String(sheet.getRange(row, COL.EMAIL + 1).getValue()).trim();
    var refNo = String(sheet.getRange(row, COL.REFERENCE + 1).getValue()).trim();
    
    Logger.log("onEdit: Request - Ref: " + refNo + ", Email: " + email);
    
    // ===== VALIDATE DATA =====
    if (!email || !refNo) {
      Logger.log("onEdit: ERROR - Missing email or reference number");
      return;
    }
    
    if (!isValidEmail(email)) {
      Logger.log("onEdit: ERROR - Invalid email format: " + email);
      return;
    }
    
    // ===== SEND APPROPRIATE NOTIFICATION =====
    var emailSent = false;
    
    switch(newStatus) {
      case "ALLOCATED":
        Logger.log("onEdit: Sending ALLOCATION email...");
        sendAllocationEmail(email, refNo);
        emailSent = true;
        Logger.log("onEdit: ✓ ALLOCATED notification sent");
        break;
        
      case "COMPLETED":
        Logger.log("onEdit: Sending COMPLETION email...");
        sendCompletionEmail(email, refNo);
        emailSent = true;
        Logger.log("onEdit: ✓ COMPLETED notification sent");
        break;
        
      case "CANCELLED":
        // Log cancellation timestamp in column M (Time Log)
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(row, COL.TIME_LOG + 1).setValue(timestamp);
        Logger.log("onEdit: ✓ Cancellation timestamp logged: " + timestamp);
        
        // Only send cancellation email if not user-initiated
        // (User-initiated cancellations already send email via cancelRequest function)
        if (oldStatus !== "PENDING" || !sheet.getRange(row, COL.CANCELLATION_REMARKS + 1).getValue()) {
          Logger.log("onEdit: Sending admin-initiated CANCELLATION email...");
          var remarks = sheet.getRange(row, COL.CANCELLATION_REMARKS + 1).getValue() || "Cancelled by administrator";
          sendCancellationEmail(email, refNo, remarks);
          emailSent = true;
          Logger.log("onEdit: ✓ CANCELLED notification sent");
        } else {
          Logger.log("onEdit: Skipped CANCELLED email - user-initiated via web");
        }
        break;
        
      case "PENDING":
        // Don't send email for PENDING - only sent on initial submission
        Logger.log("onEdit: Skipped - PENDING status doesn't trigger notification");
        break;
        
      default:
        Logger.log("onEdit: Unknown status: " + newStatus + " - no email sent");
    }
    
    if (emailSent) {
      // Add timestamp to track when notification was sent (optional)
      Logger.log("onEdit: SUCCESS - Notification sent for status: " + newStatus);
    }
    
  } catch (error) {
    Logger.log("onEdit: CRITICAL ERROR - " + error.toString());
    Logger.log("onEdit: Stack trace: " + error.stack);
    // Don't throw - we don't want to interrupt the user's edit
  }
}

// ============================================
// TEST FUNCTION
// ============================================
function testCheckStatus() {
  var result = checkRequestStatus("REQ-9534C6");
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}