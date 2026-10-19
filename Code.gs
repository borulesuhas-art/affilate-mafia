/**
 * TASK REWARDS APP — Apps Script backend
 * -------------------------------------------------
 * 1. Open your Google Sheet (the one with Tasks / Submissions / Dashboard tabs).
 * 2. Extensions -> Apps Script.
 * 3. Delete any starter code, paste ALL of this file in.
 * 4. Click Deploy -> New deployment -> type: Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Copy the Web app URL it gives you — paste it into API_URL in index.html.
 * -------------------------------------------------
 */

var TASKS_SHEET = "Tasks";
var SUBMISSIONS_SHEET = "Submissions";
var DRIVE_FOLDER_NAME = "TaskRewards Screenshots";

function doGet(e) {
  var action = e.parameter.action;
  if (action === "tasks") {
    return jsonResponse(getActiveTasks());
  }
  return jsonResponse({ ok: true, message: "Task Rewards API is live." });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var result = addSubmission(body);
    return jsonResponse({ ok: true, submissionId: result });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function getActiveTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASKS_SHEET);
  var data = sheet.getDataRange().getValues();
  var tasks = [];
  // Columns: A TaskID, B OfferName, C Platform, D Steps, E ReferralLink,
  //          F PayoutToUser, G YourCommission, H NetProfit, I Status
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] && row[8] === "Active") {
      tasks.push({
        taskId: row[0],
        offerName: row[1],
        platform: row[2],
        steps: row[3],
        referralLink: row[4],
        payoutToUser: row[5]
      });
    }
  }
  return tasks;
}

function addSubmission(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);

  // find next empty row based on column A (Submission ID)
  var colA = sheet.getRange("A2:A1000").getValues();
  var rowIndex = -1;
  for (var i = 0; i < colA.length; i++) {
    if (!colA[i][0]) { rowIndex = i + 2; break; } // +2 because range starts at row 2
  }
  if (rowIndex === -1) throw new Error("Submissions sheet is full. Extend the formula rows first.");

  var submissionId = "S" + Utilities.formatString("%03d", rowIndex - 1);
  var screenshotUrl = "";

  if (body.screenshotBase64) {
    var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
    var bytes = Utilities.base64Decode(body.screenshotBase64);
    var blob = Utilities.newBlob(bytes, body.screenshotMimeType || "image/jpeg",
                                  (body.screenshotFileName || submissionId) + ".jpg");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    screenshotUrl = file.getUrl();
  }

  // Write only the real (non-formula) columns: A,B,C,D,E,G,H,I
  sheet.getRange(rowIndex, 1).setValue(submissionId);                 // A Submission ID
  sheet.getRange(rowIndex, 2).setValue(new Date());                   // B Date
  sheet.getRange(rowIndex, 3).setValue(body.userName || "");          // C User Name
  sheet.getRange(rowIndex, 4).setValue(body.mobile || "");            // D Mobile Number
  sheet.getRange(rowIndex, 5).setValue(body.taskId || "");            // E Task ID (triggers F,J,K,L formulas)
  sheet.getRange(rowIndex, 7).setValue(screenshotUrl);                // G Screenshot Link
  sheet.getRange(rowIndex, 8).setValue(body.upiId || "");             // H UPI ID
  sheet.getRange(rowIndex, 9).setValue("Pending");                    // I Status

  return submissionId;
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
