// ===================== CONFIG =====================
// Paste this into the Apps Script project bound to (or pointed at) your spreadsheet.
// Deploy as: Execute as "User accessing the web app", Access "Anyone within mbstu.ac.bd"
// You must deploy this from your own @mbstu.ac.bd account for the domain option to appear.

const SHEET_ID = 'YOUR_SPREADSHEET_ID'; // copy from the sheet's URL
const ALLOWED_DOMAIN = 'mbstu.ac.bd';

// Expected tabs in the spreadsheet: Attendance, Submissions, QuizResults
// Attendance columns:   Email | Week | Date | Timestamp
// Submissions columns:  Email | Week | DriveLink | Timestamp
// QuizResults columns:  Email | QuizId | Score | Timestamp

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// Confirms the caller is signed in with a verified @mbstu.ac.bd account.
// Belt-and-suspenders: the domain-restricted deployment already blocks
// outsiders before this code even runs, but checking again is cheap.
function getVerifiedEmail() {
  const email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@' + ALLOWED_DOMAIN)) {
    throw new Error('Access denied: sign in with an @' + ALLOWED_DOMAIN + ' account.');
  }
  return email;
}

// GET handles both reads and JSONP writes (anything that needs a response back)
function doGet(e) {
  return route(e);
}

// POST handles simple fire-and-forget writes (matches your existing form pattern)
function doPost(e) {
  return route(e);
}

function route(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch (action) {
      case 'markAttendance':   result = markAttendance(e.parameter); break;
      case 'submitAssignment': result = submitAssignment(e.parameter); break;
      case 'submitQuiz':       result = submitQuiz(e.parameter); break;
      case 'getMyAttendance':  result = getMyAttendance(); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return respond(result, e.parameter.callback);
}

// If a callback param is present, wrap as JSONP (script-tag readable).
// Otherwise return plain JSON (used by your existing no-cors fetch calls,
// which never read the body anyway).
function respond(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function markAttendance(p) {
  const email = getVerifiedEmail();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // avoid race conditions when many students click at once
  try {
    const sheet = getSheet('Attendance');
    const today = Utilities.formatDate(new Date(), 'GMT+6', 'yyyy-MM-dd');
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === email && rows[i][1] === p.week && rows[i][2] === today) {
        return { status: 'already_marked' };
      }
    }
    sheet.appendRow([email, p.week, today, new Date()]);
    return { status: 'ok' };
  } finally {
    lock.releaseLock();
  }
}

function submitAssignment(p) {
  const email = getVerifiedEmail();
  if (!p.driveLink) throw new Error('Missing Drive link.');
  getSheet('Submissions').appendRow([email, p.week, p.driveLink, new Date()]);
  return { status: 'ok' };
}

// Simple example: answers come in as "a,c,b,d" and the key lives in
// row 1 of QuizResults' neighbouring "QuizKeys" sheet, same format.
// Replace with your real quiz structure — this just shows the pattern.
function submitQuiz(p) {
  const email = getVerifiedEmail();
  const keyRow = getSheet('QuizKeys').getRange(1, 1, 1, 50).getValues()[0]
    .filter(v => v !== '');
  const studentAnswers = String(p.answers).split(',');
  let score = 0;
  studentAnswers.forEach((ans, i) => { if (ans === keyRow[i]) score++; });

  getSheet('QuizResults').appendRow([email, p.quizId, score, new Date()]);
  return { status: 'ok', score: score, total: keyRow.length };
}

function getMyAttendance() {
  const email = getVerifiedEmail();
  const rows = getSheet('Attendance').getDataRange().getValues();
  const records = rows.filter((row, i) => i > 0 && row[0] === email)
    .map(row => ({ week: row[1], date: row[2] }));
  return { status: 'ok', records: records };
}
