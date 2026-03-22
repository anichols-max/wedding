/**
 * Google Apps Script — RSVP webhook receiver (smart match)
 *
 * MATCHING LOGIC:
 * 1. First + last name both found in cell → MATCH (confident)
 * 2. First name only, and it's unique across all rows → MATCH
 * 3. First name only, multiple rows have it → NO MATCH (too risky)
 * 4. No match at all → append at bottom tagged "not-on-list"
 *
 * Only writes to RSVP response fields — never overwrites pre-filled data.
 */

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = e.parameter;

  var headerMap = {};
  for (var i = 0; i < headers.length; i++) {
    headerMap[headers[i].toString().trim().toLowerCase()] = i;
  }

  var submittedName = (data.household || '').trim();
  var nameParts = submittedName.toLowerCase().split(/\s+/);
  var firstName = nameParts[0] || '';
  var lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : '';

  if (!firstName) {
    return appendNewRow(sheet, headers, headerMap, data, 'not-on-list');
  }

  var householdCol = headerMap['household'];
  if (householdCol === undefined) {
    return appendNewRow(sheet, headers, headerMap, data, 'no-household-column');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return appendNewRow(sheet, headers, headerMap, data, 'not-on-list');
  }

  var householdData = sheet.getRange(2, householdCol + 1, lastRow - 1, 1).getValues();

  // Pass 1: look for rows where BOTH first and last name appear
  var fullMatches = [];
  if (lastName) {
    for (var r = 0; r < householdData.length; r++) {
      var cell = householdData[r][0].toString().toLowerCase();
      if (!cell) continue;
      if (cell.indexOf(firstName) !== -1 && cell.indexOf(lastName) !== -1) {
        fullMatches.push(r + 2);
      }
    }
  }

  // If exactly one full match, use it
  if (fullMatches.length === 1) {
    return updateRow(sheet, headerMap, data, fullMatches[0]);
  }

  // Pass 2: look for rows where first name appears
  var firstNameMatches = [];
  for (var r2 = 0; r2 < householdData.length; r2++) {
    var cell2 = householdData[r2][0].toString().toLowerCase();
    if (!cell2) continue;

    // Split cell into individual names (split on & , and whitespace)
    var cellNames = cell2.split(/[&,]+/).map(function(s) { return s.trim(); });
    for (var n = 0; n < cellNames.length; n++) {
      var namePart = cellNames[n].split(/\s+/)[0]; // first word of each name chunk
      if (namePart === firstName) {
        firstNameMatches.push(r2 + 2);
        break;
      }
    }
  }

  // If exactly one first-name match, use it
  if (firstNameMatches.length === 1) {
    return updateRow(sheet, headerMap, data, firstNameMatches[0]);
  }

  // Multiple first-name matches or no matches — can't be sure, append as new
  return appendNewRow(sheet, headers, headerMap, data, 'not-on-list');
}

function updateRow(sheet, headerMap, data, matchRow) {
  var fieldsToUpdate = [
    'rsvp_status', 'attending',
    'adults_attending', 'kids_attending',
    'food_choices', 'dietary_notes', 'drinkers'
  ];

  for (var f = 0; f < fieldsToUpdate.length; f++) {
    var field = fieldsToUpdate[f];
    var col = headerMap[field];
    if (col !== undefined && data[field] !== undefined && data[field] !== '') {
      sheet.getRange(matchRow, col + 1).setValue(data[field]);
    }
  }

  var tagsCol = headerMap['tags'];
  if (tagsCol !== undefined) {
    var existingTags = sheet.getRange(matchRow, tagsCol + 1).getValue().toString();
    if (existingTags.indexOf('web-rsvp') === -1) {
      var newTags = existingTags ? existingTags + ', web-rsvp' : 'web-rsvp';
      sheet.getRange(matchRow, tagsCol + 1).setValue(newTags);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok', matched: true, row: matchRow }))
    .setMimeType(ContentService.MimeType.JSON);
}

function appendNewRow(sheet, headers, headerMap, data, tag) {
  var row = headers.map(function(header) {
    var key = header.toString().trim().toLowerCase();
    return data[key] !== undefined ? data[key] : '';
  });

  var tagsCol = headerMap['tags'];
  if (tagsCol !== undefined) {
    row[tagsCol] = tag;
  }

  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok', matched: false, tag: tag }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'RSVP webhook is live (smart match)' }))
    .setMimeType(ContentService.MimeType.JSON);
}
