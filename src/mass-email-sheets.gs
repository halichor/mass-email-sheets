const LOG_SHEET_NAME = 'ConsoleLog';
const STATUS_COLUMN_NAME = 'Status'; // Automatically updated

function sendFlexibleMailMerge() {
  const ui = SpreadsheetApp.getUi(); // Get the UI object to show a message box
  const response = ui.alert('Are you sure you want to send the emails?', ui.ButtonSet.YES_NO);

  // Check if the user clicked "Yes" to proceed
  if (response == ui.Button.NO) {
    ui.alert('Mail merge process canceled.');
    return; // Exit the function if the user cancels
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  const data = sheet.getDataRange().getValues();

  // Change the header (index) row here. The value should be (-1) from the current row, since indexes are 0-based.
  // Ex. Your column headers are in row 6, so the value should be 5.
  const HEADER_ROW_INDEX = 1;  
  const DATA_START_ROW_INDEX = HEADER_ROW_INDEX + 2;  // Adjust this to skip rows. Use (+ 1) to read data right after the header.
 
  const headers = data[HEADER_ROW_INDEX];
  const statusColIndex = headers.indexOf(STATUS_COLUMN_NAME);
  const now = new Date();
  
  // Reads the Google Doc ID from cell C1. Change this if the cell location is different.
  const docId = sheet.getRange("C1").getValue();
  
  if (!docId) {
    logToConsoleSheet("No template Doc ID found", "Please enter a valid Google Doc ID in cell C1.");
    return;
  }

  for (let i = DATA_START_ROW_INDEX; i < data.length; i++) {
  const row = data[i];
    const rowData = {};
    headers.forEach((header, index) => rowData[header.trim()] = row[index]);

    const to = rowData["To"];
    if (!to) {
      logToConsoleSheet("Skipped row (missing 'To')", JSON.stringify(row));
      continue;
    }

    try {
      const cc = rowData["Cc"] || "";
      const bcc = rowData["Bcc"] || "";
      const schedule = rowData["Schedule"];
      const subject = replacePlaceholders(rowData["Subject"] || "No Subject", rowData);
      const htmlBody = generateBodyFromGoogleDoc(docId, rowData);
      const attachmentLinks = getAttachmentLinks(rowData["AttachmentIDs"]);

      // Get Drive file attachments
      const attachments = getDriveAttachments(attachmentLinks);

      // If no attachments are found, log it but don't stop the email process
      if (attachments.length === 0) {
        logToConsoleSheet("No attachments found", `To: ${to} | Attachments: ${attachmentLinks}`);
      }

      // Check attachments permissions before sending
      checkAttachmentsPermissions(attachmentLinks, rowData["To"] + "," + rowData["Cc"] + "," + rowData["Bcc"]);

    // Get the (first) "Send As" email signature in the active account 
    function getGmailSignature() {
      try {
        const sendAsList = Gmail.Users.Settings.SendAs.list('me');
        const signature = sendAsList.sendAs[0].signature || ''; // Get the first sendAs account signature
        return signature;
      } catch (error) {
        Logger.log('Error fetching signature: ' + error.message);
        return '';  // Return an empty string if there's an error
      }
    }

    // Function to remove specific emojis using Unicode
    // List all broken emojis here just in case, or they'll show up as broken characters in the signature
    function removeSpecificEmojis(str) {
      return str.replace(/\u{1F4EB}|\u{1F4E0}/gu, ''); // Remove the specific emojis by Unicode
    }

    const signature = "<br>" + removeSpecificEmojis(getGmailSignature()); // Remove specific emojis from signature

    // Define email options
    const emailOptions = {
      cc: cc,
      bcc: bcc,
      htmlBody: htmlBody + signature,  // Append the cleaned-up signature
      attachments: attachments.length > 0 ? attachments : [],  // Only include attachments if present
    };

      let sentTime = "";

      if (schedule === "") {
        // No schedule, send immediately
        GmailApp.sendEmail(to, subject, "", emailOptions);
        sentTime = new Date().toLocaleString();
        logToConsoleSheet("Email sent successfully", `To: ${to} at ${sentTime}`);
        sheet.getRange(i + 1, statusColIndex + 1).setValue(`Sent: ${sentTime}`);
      } else {
        // There is a schedule, create draft and send at scheduled time
        const scheduledDate = new Date(schedule);
        if (isNaN(scheduledDate)) {
          // Alert if the schedule date is invalid
          ui.alert('Invalid Schedule Date Format', `Row ${i + 1}: Please enter a valid date format in the "Schedule" column. The email will not be sent.`, ui.ButtonSet.OK);
          logToConsoleSheet("Error: Invalid Schedule Date", `Row ${i + 1}: Invalid date format`);
          sheet.getRange(i + 1, statusColIndex + 1).setValue("Error: Invalid date format");
          continue;
        }
      
        if (scheduledDate > now) {
          // Create a draft email
          const draft = GmailApp.createDraft(to, subject, "", emailOptions);
      
          // Wait until the scheduled time and then send the draft
          const timeInMillis = scheduledDate.getTime() - now.getTime();
          if (timeInMillis > 0) {
            // Sleep until the scheduled time
            Utilities.sleep(timeInMillis);  // Sleep until the scheduled time
            draft.send();  // Send the draft email immediately after sleep
            sentTime = scheduledDate.toLocaleString();
            logToConsoleSheet("Scheduled email sent", `To: ${to}, At: ${sentTime}`);
            sheet.getRange(i + 1, statusColIndex + 1).setValue(`Scheduled: ${sentTime}`);
          } else {
            ui.alert('Past Date Error', `Row ${i + 1}: The schedule date is in the past. Please provide a future date. The email will not be sent.`, ui.ButtonSet.OK);
            logToConsoleSheet("Error: Past Date", `Row ${i + 1}: Schedule date is in the past`);
            sheet.getRange(i + 1, statusColIndex + 1).setValue("Error: Past date");
          }
        } else {
          // Schedule date is in the past, show an error
          ui.alert('Past Date Error', `Row ${i + 1}: The schedule date is in the past. Please provide a future date. The email will not be sent.`, ui.ButtonSet.OK);
          logToConsoleSheet("Error: Past Date", `Row ${i + 1}: Schedule date is in the past`);
          sheet.getRange(i + 1, statusColIndex + 1).setValue("Error: Past date");
        }
      }      

    } catch (err) {
      logToConsoleSheet("Error sending email", `To: ${to} | Error: ${err.message}`);
      if (statusColIndex !== -1) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue("Error: " + err.message);
      }
    }
  }
}

// Replace {{placeholders}} in strings
function replacePlaceholders(template, data, attachmentLinks) {
  return template.replace(/{{(.*?)}}/g, (_, key) => {
    let value = data[key.trim()] ?? '';

    // If the key/value is "AttachmentIDs", do not convert the links into clickable HTML. Send as an attachment instead.
    if (key.trim() === "AttachmentIDs") {
      return value; // Return the raw links as is
    }

    // Check if the value is a URL (using regex for HTTP/HTTPS links)
    if (value.match(/^https?:\/\/[^\s]+$/)) {
      return value; // Treat it as a raw link and return it as is, no <a> tag. Drive links will show under "Shared in Drive"
    }

    // Otherwise, return the value as-is
    return value;
  });
}

// Generate rich HTML body from a Google Doc template with inline styles
function generateBodyFromGoogleDoc(docId, data) {
  const url = `https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=text/html`;
  
  // Get OAuth token for authorization
  const token = ScriptApp.getOAuthToken();
  
  // Fetch the content of the Google Doc as HTML with inline styles
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    }
  });
  
  // Get the HTML content of the Google Doc
  let html = response.getContentText();
  
  // Optionally sanitize for Gmail (removing styles, fonts, tables, etc.)
  html = sanitizeForGmail(html);
  
  // Replace placeholders in the HTML with data
  html = replacePlaceholders(html, data);
  
  return html;
}

// Sanitize template HTML to retain only essential inline styles, remove unwanted ones, and convert tables to paragraphs
function sanitizeForGmail(html) {
  // Remove the <style> block entirely
  html = html.replace(/<style[\s\S]*?<\/style>/gi, '');

  // Remove the <body> tag with inline styles (e.g., background-color, max-width, etc.)
  html = html.replace(/<body[^>]*style="[^"]*"[^>]*>/gi, '<body>'); // Remove inline styles from body tag

  // Remove all inline styles except for allowed properties
  html = html.replace(/<([a-zA-Z]+)([^>]*)style="([^"]*)"/gi, (match, tag, attrs, style) => {
    const allowedStyles = ['color', 'font-size', 'text-decoration', 'font-weight', 'text-align'];
    const styles = style.split(';').filter(s => {
      const [property] = s.split(':').map(val => val.trim());
      return allowedStyles.includes(property);
    }).join(';');
    
    // Only return the tag with allowed inline styles
    return `<${tag}${attrs}${styles ? ` style="${styles}"` : ''}>`;
  });

  // Ensure that <table>, <tr>, <td> and <th> are converted into <p> tags
  html = html.replace(/<table[^>]*>/gi, '');
  html = html.replace(/<\/table>/gi, '');

  // Convert <tr> and <td> to <p> tags for Gmail compatibility
  html = html.replace(/<tr[^>]*>/gi, '<div>');
  html = html.replace(/<\/tr>/gi, '</div>');
  html = html.replace(/<td[^>]*>/gi, '<div>');
  html = html.replace(/<\/td>/gi, '</div>');

  // Replace <th> with <p> as well
  html = html.replace(/<th[^>]*>/gi, '<p>');
  html = html.replace(/<\/th>/gi, '</p>');

  // Fix: remove any extra closing tags like ">>" caused by unbalanced tags
  html = html.replace(/<\s*([a-zA-Z]+)[^>]*>[\s]*<\s*\//g, '></');  // Look for mismatched tags and fix them

  // Replace all occurrences of ">>" with ">"
  html = html.replace(/>>/g, '>');  // Fix any stray ">>" characters

  // Clean up any stray "><" that may have been inserted
  html = html.replace(/><\s*/g, '> <'); // Ensure no stray closing tags

  // Fix self-closing tags
  html = html.replace(/<([a-zA-Z]+)[^>]*\/>/g, '<$1>'); // Fix self-closing tags like <img />

  // Clean up extra spaces in <a> tag href attributes (remove leading/trailing spaces)
  html = html.replace(/<a\s+href="(.*?)\s+"/g, '<a href="$1"'); // Remove trailing spaces in href values
  html = html.replace(/<a\s+href="\s+(.*?)"/g, '<a href="$1"'); // Remove leading spaces in href values

  // Collapse multiple newlines or line breaks into a single <br>
  html = html.replace(/(\r\n|\n|\r){2,}/g, '<br>');

  // Remove whitespace between tags
  html = html.replace(/>\s+</g, '><');

  // Trim the final result
  return html.trim();
}

// Get attachment links from a comma-separated string
function getAttachmentLinks(linkString) {
  if (!linkString) return [];
  return linkString.split(',').map(link => link.trim()).filter(Boolean);
}

// Get Drive file attachments based on the file links (Google Drive file URLs)
function getDriveAttachments(attachmentLinks) {
  const attachments = [];
  
  attachmentLinks.forEach(link => {
    const fileId = extractFileIdFromLink(link);
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        attachments.push(file);  // Attach file
        // attachments.push(file.getAs(MimeType.PDF)); - Attach as PDF (or change MimeType if needed)
      } catch (e) {
        Logger.log("Error retrieving file: " + e.message);
      }
    }
  });
  
  return attachments;
}

// Extract file ID from a Google Drive URL
function extractFileIdFromLink(link) {
  const regex = /(?:drive|docs)\.google\.com\/.*?\/d\/(.*?)(?:\/|$)/;
  const match = link.match(regex);
  return match ? match[1] : null;
}

// Check if all recipients have permission to view the attachment links
function checkAttachmentsPermissions(attachmentLinks, recipients) {
  const emails = recipients.split(',').map(email => email.trim()).filter(Boolean);
  let filesMissingPermission = [];
  
  attachmentLinks.forEach(link => {
    const fileId = extractFileIdFromLink(link);
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        const fileEditors = file.getEditors().map(user => user.getEmail());
        const fileViewers = file.getViewers().map(user => user.getEmail());
        
        const missingPermissions = emails.filter(email => !fileViewers.includes(email) && !fileEditors.includes(email));
        
        // Check if file is shared as "Anyone with the link can view"
        if (file.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK) {
          return;  // If it's shared as "Anyone with the link", it's considered accessible
        }
        
        if (missingPermissions.length > 0) {
          filesMissingPermission.push({
            file: link,
            missingPermissions: missingPermissions.join(', ')
          });
        }
      } catch (e) {
        Logger.log("Error retrieving file permissions: " + e.message);
      }
    }
  });

  // If any files are missing permissions, inform the user
  if (filesMissingPermission.length > 0) {
    const ui = SpreadsheetApp.getUi();
    let message = 'The following attachments do not have the necessary view permissions for some recipients:\n\n';
    
    filesMissingPermission.forEach(file => {
      message += `File: ${file.file}\nMissing Permissions: ${file.missingPermissions}\n\n`;
    });
    
    ui.alert('Permissions Warning', message, ui.ButtonSet.OK);
    logToConsoleSheet('Permissions Warning', message);
  }
}

// Log to a "ConsoleLog" sheet
function logToConsoleSheet(message, details = "") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(["Timestamp", "Message", "Details"]);
  }
  logSheet.appendRow([new Date(), message, details]);
}

// Run this to preview the HTML output before sending emails (appears in console)
function testGenerateHtmlFromGoogleDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  const docId = sheet.getRange("B1").getValue(); // Assuming the Google Doc ID is in B1

  const rowData = {}; // Replace with actual data for placeholders
  const htmlBody = generateBodyFromGoogleDoc(docId, rowData);

  Logger.log(htmlBody); // Logs the generated HTML for inspection
}
