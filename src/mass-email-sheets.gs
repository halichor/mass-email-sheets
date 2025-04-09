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
  
  // Adjusted for header row starting at row 2 (index 1)
  const headers = data[1];  // Read headers from row 2
  const statusColIndex = headers.indexOf(STATUS_COLUMN_NAME);
  const now = new Date();
  
  // Get the Google Doc ID from cell B1 in the "Data" tab
  const docId = sheet.getRange("B1").getValue();
  
  if (!docId) {
    logToConsoleSheet("No template Doc ID found", "Please enter a valid Google Doc ID in cell B1.");
    return;
  }

  for (let i = 2; i < data.length; i++) {  // Start from row 3 (index 2)
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

      if (attachments.length === 0) {
        logToConsoleSheet("Error: No valid attachments found", `To: ${to} | Attachments: ${attachmentLinks}`);
        sheet.getRange(i + 1, statusColIndex + 1).setValue("Error: No valid attachments");
        continue;
      }

      // Check attachments permissions before sending
      checkAttachmentsPermissions(attachmentLinks, rowData["To"] + "," + rowData["Cc"] + "," + rowData["Bcc"]);

      // Define email options
      const emailOptions = {
        cc: cc,
        bcc: bcc,
        htmlBody: htmlBody,
        attachments: attachments,
      };

      let sentTime = "";
      if (schedule === "") {
        // Send email immediately
        GmailApp.sendEmail(to, subject, "", emailOptions);
        sentTime = new Date().toLocaleString();
        logToConsoleSheet("Email sent successfully", `To: ${to} at ${sentTime}`);
        sheet.getRange(i + 1, statusColIndex + 1).setValue(`Sent: ${sentTime}`);
      } else {
        // Validate schedule format
        const scheduledDate = new Date(schedule);
        if (isNaN(scheduledDate)) {
          // Alert if the schedule date is invalid
          ui.alert('Invalid Schedule Date Format', `Row ${i + 1}: Please enter a valid date format in the "Schedule" column. The email will not be sent.`, ui.ButtonSet.OK);
          logToConsoleSheet("Error: Invalid Schedule Date", `Row ${i + 1}: Invalid date format`);
          sheet.getRange(i + 1, statusColIndex + 1).setValue("Error: Invalid date format");
          continue;
        }

        if (scheduledDate > now) {
          // Schedule email using a time-driven trigger
          const draft = GmailApp.createDraft(to, subject, "", emailOptions);

          // Trigger the email to send at the specified time
          const timeInMillis = scheduledDate.getTime() - now.getTime();
          if (timeInMillis > 0) {
            // Wait for the scheduled time and then send the draft
            Utilities.sleep(timeInMillis);  // Sleep until the scheduled time
            draft.send();  // Send the draft email
            sentTime = scheduledDate.toLocaleString();
            logToConsoleSheet("Scheduled email sent", `To: ${to}, At: ${sentTime}`);
            sheet.getRange(i + 1, statusColIndex + 1).setValue(`Scheduled: ${sentTime}`);
          } else {
            ui.alert('Past Date Error', `Row ${i + 1}: The schedule date is in the past. Please provide a future date. The email will not be sent.`, ui.ButtonSet.OK);
            logToConsoleSheet("Error: Past Date", `Row ${i + 1}: Schedule date is in the past`);
            sheet.getRange(i + 1, statusColIndex + 1).setValue("Error: Past date");
          }
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
function replacePlaceholders(template, data) {
  return template.replace(/{{(.*?)}}/g, (_, key) => data[key.trim()] ?? '');
}

// Generate rich HTML body from a Google Doc template
function generateBodyFromGoogleDoc(docId, data) {
  const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${docId}&exportFormat=html`;
  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token
    }
  });

  let html = response.getContentText();

  // Optional: Clean up Word-style formatting, fonts, etc.
  html = sanitizeForGmail(html);

  // Replace placeholders
  html = replacePlaceholders(html, data);

  return html;
}

// Sanitize template HTML to plain text
function sanitizeForGmail(html) {
  // Remove styles and fonts
  html = html.replace(/<style[\s\S]*?<\/style>/gi, '');
  html = html.replace(/<[^>]*style="[^"]*"[^>]*>/gi, tag => tag.replace(/style="[^"]*"/gi, ''));
  html = html.replace(/<font[^>]*>|<\/font>/gi, '');
  
  // Optional: remove extra table wrappers if you want
  html = html.replace(/<table[^>]*>|<\/table>/gi, '');
  html = html.replace(/<tr[^>]*>|<\/tr>/gi, '');
  html = html.replace(/<td[^>]*>|<\/td>/gi, '');

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
        attachments.push(file.getAs(MimeType.PDF));  // Attach as PDF (or change MimeType if needed)
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
