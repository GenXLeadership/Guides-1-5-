function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);
  const fields = ['timestamp','name','email','stillDoing','tankLevel','bodySigns','storyTelling','fearSlowDown','restMoment','yearFromNow','wantedFeeling','permission','firstMove','moveTiming','tellWho','leaveLight'];
  const labels = [
    'Timestamp',
    'Name',
    'Email',
    '1. Still doing',
    '2. Tank level',
    '3. Where it sits in the body',
    '4. Story they tell themselves',
    '5. What they fear if they slow down',
    '6. Last time they rested',
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
