// Paste this into Extensions → Apps Script of your Google Sheet, then Deploy → New deployment → Web app
// (Execute as: Me, Who has access: Anyone). Copy the web-app URL into FORM_ENDPOINT in index.html.
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);
  let headers = sheet.getLastRow() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  Object.keys(data).forEach(k => { if (!headers.includes(k)) { headers.push(k); sheet.getRange(1, headers.length).setValue(k); } });
  sheet.appendRow(headers.map(h => data[h] ?? ''));
  return ContentService.createTextOutput('ok');
}
