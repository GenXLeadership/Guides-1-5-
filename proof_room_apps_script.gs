function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);
  const fields = ['timestamp','name','email','honestVoice','howLong','bodyLocation','loudVoice','fearReveal','fraudMoment','yearFromNow','wantedFeeling','permission','firstMove','moveTiming','tellWho','leaveLight'];
  const labels = [
    'Timestamp',
    'Name',
    'Email',
    '1. Honest voice',
    '2. How long feeling this way',
    '3. Where it sits in the body',
    '4. The loud voice in their head',
    '5. What they fear people will see',
    '6. A specific fraud moment',
    '7. Feeling a year from now',
    '8. Wanted feeling',
    '9. Permission declared',
    '10. First move',
    '11. Move timing',
    '12. Tell who',
    '13. Leave light for others'
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
