# Cohort-facilitators

## Google Sheets storage

1. Create a Google Sheet for responses.
2. Open **Extensions → Apps Script** and paste the script below.
3. Update `SHEET_NAME` if you want a custom sheet tab name.
4. Update `ALLOWED_ORIGIN` to your deployed site domain.
5. Deploy as **Web app** (Execute as: **Me**, Who has access: **Anyone**).
6. Copy the Web App URL into `GOOGLE_SHEET_WEB_APP_URL` in `index.html`.

### Apps Script (Code.gs)
```javascript
const SHEET_NAME = "Responses";
const ALLOWED_ORIGIN = "https://yourdomain.com";

function doPost(e) {
  let payload = {};
  try {
    payload = JSON.parse(e.postData.contents || "{}");
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid JSON payload" }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const headers = [
    "Submitted At",
    "Full Name",
    "Phone Number",
    "Email Address",
    "Social Handles",
    "Headshot",
    "Primary Role",
    "Mentor Interest",
    "Mentor Topic",
    "Mentor Availability",
    "Short Bio",
    "Session Preference",
    "Organization Name",
    "Organization Address",
    "Logo Upload",
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
    JSON.stringify(payload.headshot || null),
    payload.primaryRole || "",
    payload.mentorInterest || "",
    payload.mentorTopic || "",
    payload.mentorAvailability || "",
    payload.shortBio || "",
    JSON.stringify(payload.sessionPreference || []),
    payload.orgName || "",
    payload.orgAddress || "",
    JSON.stringify(payload.logoUpload || null),
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

The headshot and logo fields are posted as JSON payloads (including data URLs). If you want to store the images in Drive instead of the sheet, extend the Apps Script to write files and store the resulting file links.

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
