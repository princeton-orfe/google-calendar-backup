function backupCalendarsFromSheet() {
  const { sheetId, sheetTabName, backupFolderName } = getConfigOrThrow();
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName(sheetTabName);

  if (!sheet) {
    const msg = `Sheet tab '${sheetTabName}' not found in spreadsheet '${sheetId}'. Verify the Script Property SHEET_TAB_NAME.`;
    Logger.log(msg);
    throw new Error(msg);
  }

  const parentFolder = getOrCreateFolder(backupFolderName); // Main backup folder
  let log = "";
  let successCount = 0;
  let failureCount = 0;

  // Get all rows and exclude the header
  const data = sheet.getDataRange().getValues().slice(1); // Skip header row

  data.forEach(row => {
    const calendarName = row[0];
    const icsUrl = (row[1] || "").toString().trim();
    const sanitizedCalendarName = calendarName.replace(/[^a-zA-Z0-9_-]/g, "_");

    if (!calendarName || !icsUrl) {
      log += `⚠️ Missing data for one row. Skipping row with Calendar Name: '${calendarName}'.\n`;
      failureCount++;
      return;
    }

    // Validate Google Calendar ICS URL shape
    if (!isValidGoogleCalendarIcsUrl(icsUrl)) {
      log += `⚠️ Invalid Google Calendar ICS URL for '${calendarName}'. Skipping this row.\n`;
      failureCount++;
      return;
    }

    // Create a subfolder for each calendar
    const calendarFolder = getOrCreateFolder(sanitizedCalendarName, parentFolder);
    
    try {
      const response = UrlFetchApp.fetch(icsUrl);
      const fileName = `${sanitizedCalendarName}_${formatDate(new Date())}.ics`;
      calendarFolder.createFile(fileName, response.getContentText(), MimeType.PLAIN_TEXT);
      log += `✅ Successfully saved calendar: ${calendarName} as ${fileName}\n`;
      successCount++;
      
      // Manage file retention
      log += manageBackupRetention(calendarFolder, sanitizedCalendarName);
    } catch (error) {
      log += `❌ Failed to retrieve calendar: ${calendarName}. Error: ${error.message}\n`;
      failureCount++;
    }
  });

  // Summary report
  log += `\nSummary:\nTotal Successes: ${successCount}\nTotal Failures: ${failureCount}\n`;
  sendSummaryEmail(log);
}

function getConfigOrThrow() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty("SHEET_ID");
  const sheetTabName = props.getProperty("SHEET_TAB_NAME");
  const backupFolderName = props.getProperty("BACKUP_FOLDER_NAME") || "Google Calendar Backups";

  if (!sheetId || !sheetTabName) {
    const missing = [];
    if (!sheetId) missing.push("SHEET_ID");
    if (!sheetTabName) missing.push("SHEET_TAB_NAME");
    const msg = `Missing Script Properties: ${missing.join(", ")}.
Set them in Apps Script: Project Settings > Script properties (or File > Project properties > Script properties).`;
    Logger.log(msg);
    throw new Error(msg);
  }

  return { sheetId, sheetTabName, backupFolderName };
}

/**
 * Utility: Log resolved configuration for sanity checks.
 * Call manually from the Apps Script editor to verify configuration.
 */
function logResolvedConfiguration() {
  const { sheetId, sheetTabName, backupFolderName } = getConfigOrThrow();
  Logger.log(`SHEET_ID: ${sheetId}`);
  Logger.log(`SHEET_TAB_NAME: ${sheetTabName}`);
  Logger.log(`BACKUP_FOLDER_NAME: ${backupFolderName}`);
}

/**
 * Basic validator to check whether a URL looks like a Google Calendar ICS link.
 * Accepts hosts calendar.google.com or www.google.com and paths under /calendar/ical/.../basic.ics
 */
function isValidGoogleCalendarIcsUrl(url) {
  if (!url || typeof url !== "string") return false;
  const trimmed = url.trim();
  const re = /^https:\/\/(calendar\.google\.com|www\.google\.com)\/calendar\/ical\/.+\/basic\.ics(?:\?.*)?$/i;
  return re.test(trimmed);
}

function getOrCreateFolder(folderName, parentFolder = DriveApp) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function manageBackupRetention(folder, calendarName) {
  const files = folder.getFiles();
  const fileData = [];

  while (files.hasNext()) {
    const file = files.next();
    fileData.push({ file: file, createdDate: file.getDateCreated() });
  }

  // Sort by createdDate, newest first
  fileData.sort((a, b) => b.createdDate - a.createdDate);

  // Keep last 3 daily, 1 weekly, and 1 monthly backups
  const dailyLimit = 3;
  const weeklyLimit = 1;
  const monthlyLimit = 1;
  let dailyCount = 0;
  let weeklyCount = 0;
  let monthlyCount = 0;
  let cleanupLog = "";

  for (let i = 0; i < fileData.length; i++) {
    const ageInDays = (new Date() - fileData[i].createdDate) / (1000 * 60 * 60 * 24);
    let keepFile = false;

    if (ageInDays <= 1 && dailyCount < dailyLimit) {
      dailyCount++;
      keepFile = true;
    } else if (ageInDays <= 7 && weeklyCount < weeklyLimit) {
      weeklyCount++;
      keepFile = true;
    } else if (ageInDays > 7 && monthlyCount < monthlyLimit) {
      monthlyCount++;
      keepFile = true;
    }

    if (!keepFile) {
      fileData[i].file.setTrashed(true);
      cleanupLog += `🗑️ Deleted old backup: ${fileData[i].file.getName()} from ${calendarName} backups\n`;
    }
  }
  return cleanupLog;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd");
}

function sendSummaryEmail(log) {
  const userEmail = Session.getActiveUser().getEmail();
  const subject = "Google Calendar Backup Summary";
  const body = `Here is the summary of your Google Calendar backup:\n\n${log}`;
  
  MailApp.sendEmail(userEmail, subject, body);
  Logger.log("Summary email sent to " + userEmail);
}
