function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);
  const fields = ['timestamp','name','email','responsibleFor','balancePct','reason','jarTaps','yearFeeling','freedomWord','tellWho','leaveBehind'];
  const labels = [
    'Timestamp',
    'Name',
    'Email',
    '1. What is not theirs to carry',
    '2. Balance (percent toward self vs others)',
    '3. Real reason for carrying it',
    '4. Heavy moments tapped (jar count)',
    '5. Feeling a year from now',
    '6. Freedom word',
    '7. Tell who',
    '8. Leave light for others'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(labels);
  }

  sheet.appendRow(fields.map(f => {
    const v = data[f];
    if (v === undefined) return '';
    return Array.isArray(v) ? v.join(', ') : v;
  }));
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}
