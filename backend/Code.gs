/** @OnlyCurrentDoc */

const SHEET = 'feedings';
const MEALS = ['breakfast', 'lunch', 'dinner'];

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getSheet() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET);
}

function findRow(sheet, date) {
  const dates = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  for (let i = dates.length - 1; i >= 1; i--) {
    if (dates[i][0] === date) return i + 1;
  }
  return -1;
}

function readDay(sheet, date) {
  const row = findRow(sheet, date);
  const result = { date };
  MEALS.forEach((m, i) => {
    result[m] = row > 0 ? sheet.getRange(row, i + 2).getValue() === true : false;
  });
  return result;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json(readDay(getSheet(), today()));
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const col = MEALS.indexOf(body.meal);
  if (col < 0) return json({ error: 'meal must be breakfast, lunch, or dinner' });

  const lock = LockService.getScriptLock();   // stops two simultaneous taps clobbering each other
  lock.waitLock(5000);
  try {
    const sheet = getSheet();
    const date = today();
    let row = findRow(sheet, date);
    if (row < 0) {
      sheet.appendRow(["'" + date, false, false, false]);  // leading ' keeps the date as text
      row = sheet.getLastRow();
    }
    sheet.getRange(row, col + 2).setValue(body.fed === true);
    return json(readDay(sheet, date));
  } finally {
    lock.releaseLock();
  }
}
