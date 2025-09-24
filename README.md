# Google Calendar Backup

Automated Google Calendar to Google Drive backup using Google Apps Script. Each calendar is exported via its private iCal (ICS) URL and saved as a dated `.ics` file into Drive, one subfolder per calendar. The script keeps a small rolling history (daily/weekly/monthly) and emails a summary after each run.

## What this script does

- Reads calendar names and private iCal URLs from a Google Sheet tab (default name you choose; configured via Script Properties).
- Downloads each ICS feed and writes a file named `<CalendarName>_YYYYMMDD.ics` to Drive under `<BACKUP_FOLDER_NAME>/<CalendarName>/` (default: `Google Calendar Backups`).
- Retains recent backups and automatically trashes older ones:
  - Up to 3 daily backups (≤ 1 day old)
  - 1 weekly backup (≤ 7 days old)
  - 1 monthly backup (> 7 days old)
- Sends a summary email to the executing user after completion.

## Prerequisites

- Google Workspace account with access to the calendars you want to back up (or the private ICS links for those calendars).
- Permission to create files/folders in Google Drive where backups will be stored.
- Google Apps Script access to create a simple time-driven (cron) trigger.

Notes on iCal links:
- Use the calendar’s Secret address in iCal format. In Google Calendar (web) > Settings > [Select calendar] > Integrate calendar > "Secret address in iCal format". Copy that URL. Treat it as a secret.
- If the secret address is not shown, an administrator may have disabled it. Coordinate with your Workspace admin.

## Set up the Google Sheet

1. Create a new Google Sheet (or use an existing one) that will hold the calendars to back up.
2. Add a sheet (tab) named `Calendar URLs` (or any name you choose; you'll set it in Script Properties as `SHEET_TAB_NAME`).
3. In row 1, add the headers exactly:
   - Column A: `Calendar Name`
   - Column B: `Calendar URL`
4. For each calendar, add a row with:
   - A human-readable name (used for the Drive subfolder and file prefix).
   - The private ICS URL.

Example rows:

Calendar Name | Calendar URL
--------------|-------------
Undergrad Advising | https://calendar.google.com/calendar/ical/.../basic.ics
Seminar Series | https://calendar.google.com/calendar/ical/.../basic.ics

5. Copy the Sheet ID from the URL. For a Sheet at:
   - `https://docs.google.com/spreadsheets/d/1AbCDefGhIJkLMNopQRsTUvwxyz1234567890/edit#gid=0`
   - The Sheet ID is `1AbCDefGhIJkLMNopQRsTUvwxyz1234567890`.

## Configure the Apps Script

1. Open Google Apps Script (https://script.google.com) and create a new Standalone project (or use an existing one).
2. Create a file and paste the contents of `gcal-backup.gs` from this repository.
3. Set Script Properties so the script can find your Google Sheet and tab and (optionally) choose the backup folder:
  - Open Project Settings (gear icon) > Script properties (or File > Project properties > Script properties).
  - Add the following key/value pairs:
    - `SHEET_ID` = your Google Sheet ID (see above for how to copy it from the URL)
    - `SHEET_TAB_NAME` = the name of the sheet tab that holds the data (e.g., `Calendar URLs`)
    - `BACKUP_FOLDER_NAME` (optional) = the Drive folder to store backups (default: `Google Calendar Backups`)
  - Save the properties.
4. Verify the script time zone (used for date-stamping files): Apps Script editor > Project Settings > Time zone. The script formats dates using the project time zone.
5. Save the project.

## Authorization

On first run, Apps Script will request authorization for:
- Reading Google Sheets (`SpreadsheetApp`)
- Fetching external URLs (`UrlFetchApp`)
- Creating files in Drive (`DriveApp`)
- Sending email (`MailApp`)

Review and accept to proceed. If your organization restricts scopes, you may need admin approval.

## Run a manual test

1. In the Apps Script editor, select the function `backupCalendarsFromSheet` and click Run.
2. Watch the execution logs and check Google Drive for:
  - A new top-level folder named `Google Calendar Backups` (or your custom `BACKUP_FOLDER_NAME`)
   - Subfolders for each calendar, containing a file like `Undergrad_Advising_20250115.ics`.
3. After completion, verify you received a summary email.

## Schedule automatic backups

Set up a time-driven trigger to run daily.

1. Apps Script editor > Triggers (clock icon) > Add Trigger
2. Choose function: `backupCalendarsFromSheet`
3. Deployment: Head
4. Event source: Time-driven
5. Type of time based trigger: Day timer
6. Time of day: Choose a quiet time (e.g., 2–4 AM)
7. Save

The retention logic assumes a daily cadence to maintain a small daily/weekly/monthly history.

## Retention policy details

The script lists all `.ics` files in each calendar’s folder and sorts them by creation date (newest first). It keeps:
- Up to 3 files that are ≤ 1 day old (daily)
- Up to 1 file that is ≤ 7 days old (weekly)
- Up to 1 file that is > 7 days old (monthly)

All other older files are sent to the Drive trash. Adjust counts by editing `dailyLimit`, `weeklyLimit`, and `monthlyLimit` in `manageBackupRetention`.

Important:
- The logic uses the file creation timestamp in Drive. If you move or copy files manually, creation dates may differ.
- Trashed files are recoverable until permanently deleted per your Drive retention policy.

## Security considerations

- Treat ICS URLs as secrets. Anyone with the URL can read calendar contents. Store them only in controlled documents and do not commit them to source control.
- Restrict sharing on the Google Sheet to the minimum necessary.
- If you rotate a calendar’s secret address, update the Sheet.
- Consider using a Shared Drive with appropriate access controls if these backups are departmental assets.

## Troubleshooting

- No files appear in Drive:
  - Confirm Script Properties are set: `SHEET_ID` and `SHEET_TAB_NAME`.
  - Confirm the Sheet ID and the tab name match exactly what you set in Script Properties.
  - Ensure there are rows under the header and no blank rows before the data.
  - Manually open a calendar’s ICS URL in a browser to verify it downloads.

- Missing configuration error at start:
  - The script now exits with an informative error if `SHEET_ID` or `SHEET_TAB_NAME` are not set in Script Properties.
  - Set them under Project Settings > Script properties and re-run.

- Authorization errors:
  - Re-run and grant all requested scopes, or ask your admin to approve.

- 403/404 fetching ICS:
  - The ICS link may be disabled or rotated. Regenerate the "Secret address in iCal format" and update the Sheet.

- Exceeded maximum execution time:
  - Reduce the number of calendars per run, or create multiple triggers/functions that process subsets.

- Duplicate/odd file names:
  - Calendar names are sanitized to alphanumeric, underscore, and hyphen. Update names in the Sheet if needed.

## Customization tips

- Change the email recipient: `sendSummaryEmail` currently emails the executing user. Modify it to send to a group or specific address.
- Configure locations via Script Properties: `SHEET_ID`, `SHEET_TAB_NAME`, and optional `BACKUP_FOLDER_NAME` (default: `Google Calendar Backups`).
- Different retention policy: Edit `manageBackupRetention` thresholds and limits.

## Repository contents

- `gcal-backup.gs` — Google Apps Script source.
- `README.md` — This setup and usage guide.

## Ownership and execution context

This script can run under a dedicated automation account or a user account. Ensure the chosen account has Drive access to the destination and the ability to fetch the calendars’ ICS URLs.
