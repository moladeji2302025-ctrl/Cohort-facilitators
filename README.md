# Cohort-facilitators

## Google Sheets + Google Drive storage

Responses are saved to a Google Sheet. Headshots and logos are uploaded as real files to a **Google Drive** folder and their shareable links are stored in the sheet.

> **Important:** Sign in to Google as **mofejeriiseoluwa@gmail.com** before following these steps. The Apps Script runs as the signed-in account, so all files (Sheet and Drive folder) will live in that account.

1. Create a Google Sheet for responses inside the `mofejeriiseoluwa@gmail.com` Google Drive.
2. Open **Extensions → Apps Script** and paste the script below.
3. Update `SHEET_NAME` if you want a custom sheet tab name.
4. Update `ALLOWED_ORIGIN` to your deployed site domain.
5. Deploy as **Web app** (Execute as: **Me**, Who has access: **Anyone**).
6. Copy the Web App URL into `GOOGLE_SHEET_WEB_APP_URL` in `index.html`.

A Drive folder called **"Cohort Facilitators Submissions"** will be created automatically on the first submission, with `Headshots/` and `Logos/` subfolders inside it.

### Apps Script (Code.gs)
```javascript
const SHEET_NAME = "Responses";
const ALLOWED_ORIGIN = "https://yourdomain.com";
const DRIVE_FOLDER_NAME = "Cohort Facilitators Submissions";

function getOrCreateFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parent.createFolder(name);
}

function getRootFolder() {
  const existing = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function saveFileToDrive(fileData, subfolderName) {
  if (!fileData || !fileData.dataUrl) return "";
  try {
    const matches = fileData.dataUrl.match(/^data:([^;]+);base64,(.+)$/);
    if (!matches) return "";
    const mimeType = matches[1];
    const blob = Utilities.newBlob(
      Utilities.base64Decode(matches[2]),
      mimeType,
      fileData.name || "upload"
    );
    const subfolder = getOrCreateFolder(getRootFolder(), subfolderName);
    const file = subfolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    console.error("Drive upload failed:", err);
    return "";
  }
}

function doPost(e) {
  let payload = {};
  try {
    payload = JSON.parse(e.postData.contents || "{}");
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid JSON payload" }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
  }

  const headshotUrl = saveFileToDrive(payload.headshot, "Headshots");
  const logoUrl = saveFileToDrive(payload.logoUpload, "Logos");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const headers = [
    "Submitted At",
    "Full Name",
    "Phone Number",
    "Email Address",
    "Social Handles",
    "Headshot URL",
    "Primary Role",
    "Mentor Interest",
    "Mentor Availability",
    "Short Bio",
    "Organization Name",
    "Organization Address",
    "Logo URL",
    "Branding Items"
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  sheet.appendRow([
    payload.submittedAt || new Date().toISOString(),
    payload.fullName || "",
    payload.phoneNumber || "",
    payload.emailAddress || "",
    JSON.stringify(payload.socialHandles || []),
    headshotUrl,
    payload.primaryRole || "",
    payload.mentorInterest || "",
    payload.mentorAvailability || "",
    payload.shortBio || "",
    payload.orgName || "",
    payload.orgAddress || "",
    logoUrl,
    payload.brandingItems || ""
  ]);
  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
}

function doOptions() {
  return ContentService.createTextOutput("")
    .setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN)
    .setHeader("Access-Control-Allow-Methods", "POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}
```

The Web App URL is public by design. Consider adding basic validation, quotas, or token checks in Apps Script to reduce spam submissions.

### Optional hardening (rate limit example)
```javascript
const RATE_LIMIT_WINDOW_SEC = 60;
const RATE_LIMIT_MAX = 30;

function isRateLimited() {
  const cache = CacheService.getScriptCache();
  const key = "rate-limit";
  const count = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(count), RATE_LIMIT_WINDOW_SEC);
  return count > RATE_LIMIT_MAX;
}
```

Then inside `doPost`, add:
```javascript
if (isRateLimited()) {
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Rate limit exceeded" }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
}
```
