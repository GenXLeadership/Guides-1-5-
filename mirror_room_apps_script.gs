function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);
  const fields = ['timestamp','name','email','doorA','doorB','pull','blocker','reason','almostMoment','confidence','wantedFeeling','firstMove','moveTiming','tellWho','leaveLight'];
  const labels = [
    'Timestamp',
    'Name',
    'Email',
    '1. Door A',
    '2. Door B',
    '3. Which pulls harder',
    '4. What is blocking the choice',
    '5. Real reason for the block',
    '6. The almost-moment',
    '7. Confidence level',
    '8. Wanted feeling',
    '9. First move',
    '10. Move timing',
    '11. Tell who',
    '12. Leave light for others'
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
