// =====================================================
//  MEDIC DRONE SURVEY — Google Apps Script
//  Paste this entire file into Google Apps Script
//  then deploy as a Web App (instructions below)
// =====================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Define column headers (matches all survey questions)
    var headers = [
      'timestamp',
      'a1', 'a2', 'a3', 'a4',
      'b1', 'b2', 'b3', 'b4', 'b5',
      'c1_0', 'c1_1', 'c1_2', 'c1_3', 'c1_4', 'c1_5', 'c1_6',
      'c2', 'c3', 'c4', 'c5',
      'd1', 'd2', 'd3',
      'd4_0', 'd4_1', 'd4_2', 'd4_3', 'd4_4',
      'd5_impact_rating', 'd6', 'd7'
    ];

    // If sheet is empty, write the header row first
    if (sheet.getLastRow() === 0) {
      var friendlyHeaders = [
        'Timestamp',
        'A1: Role', 'A2: Facility type', 'A3: Years at facility', 'A4: Distance from referral hospital',
        'B1: Emergencies per month', 'B2: Emergency types', 'B3: Time to get resources', 'B4: Patient harm from delays', 'B5: Challenges during emergencies',
        'C1: Essential medicines availability', 'C1: Vaccines availability', 'C1: Blood products availability',
        'C1: First aid supplies availability', 'C1: Ultrasound availability', 'C1: PPE availability', 'C1: Lab kits availability',
        'C2: Frequent stockouts', 'C3: Ultrasound access', 'C4: Pregnant women miss scans', 'C5: Highest demand supplies',
        'D1: Awareness of drone delivery', 'D2: Most valuable drone deliveries', 'D3: Concerns about drones',
        'D4: Reduce deaths', 'D4: Facility can receive drones', 'D4: Community acceptance', 'D4: Better than current supply chain', 'D4: Trust drones with medicines',
        'D5: Impact rating (1-10)', 'D6: Likelihood of use', 'D7: Additional comments'
      ];
      sheet.appendRow(friendlyHeaders);

      // Style the header row
      var headerRange = sheet.getRange(1, 1, 1, friendlyHeaders.length);
      headerRange.setBackground('#1D9E75');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Build the data row
    var row = headers.map(function(key) {
      return data[key] !== undefined ? data[key] : '';
    });

    sheet.appendRow(row);

    // Auto-resize columns for readability
    sheet.autoResizeColumns(1, headers.length);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handles browser test visits to your script URL
function doGet(e) {
  return ContentService
    .createTextOutput('Medic Survey endpoint is live and ready.')
    .setMimeType(ContentService.MimeType.TEXT);
}
