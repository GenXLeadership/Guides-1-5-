function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);
  const fields = ['timestamp','name','email','decision','duration','bodyLocations','twoAM','fear','lastStopped','yearFeeling','wantedFeeling','firstMove','moveTiming','tellWho','leaveLight'];
  const labels = [
    'Timestamp',
    'Name',
    'Email',
    '1. Decision being weighed',
    '2. How long carrying it',
    '3. Where it sits in the body',
    '4. The 2am version of it',
    '5. Fear underneath it',
    '6. What has stopped them so far',
    '7. Feeling a year from now',
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
