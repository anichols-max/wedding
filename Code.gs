/**
 * Google Apps Script — RSVP webhook receiver (fuzzy match)
 *
 * MATCHING LOGIC (4 passes, both doPost and doGet):
 * 1. Exact: first + last name both found as substrings in cell → MATCH
 * 2. Fuzzy: first + last name both fuzzy-match words in cell → MATCH
 * 3. First name exact, unique across rows → MATCH
 * 4. First name fuzzy, unique across rows → MATCH
 * 5. No match → append at bottom tagged "not-on-list" (5-row gap)
 *
 * Only writes to RSVP response fields — never overwrites pre-filled data.
 */

// ═══ FUZZY MATCHING HELPERS ═══

function levenshtein(a, b) {
  var m = a.length, n = b.length;
  var dp = [];
  for (var i = 0; i <= m; i++) {
    dp[i] = [i];
    for (var j = 1; j <= n; j++) {
      dp[i][j] = i === 0 ? j : 0;
    }
  }
  for (var i = 1; i <= m; i++) {
    for (var j = 1; j <= n; j++) {
      if (a[i - 1] === b[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
  }
  return dp[m][n];
}

function namesClose(a, b) {
  // Allow ~1 typo per 3 characters (generous for wedding RSVP)
  // 3-char name: 1 edit, 4-5: 2, 6-8: 3, 9+: 4
  var maxDist = Math.ceil(Math.max(a.length, b.length) / 3);
  return levenshtein(a, b) <= maxDist;
}

function extractWords(cellText) {
  // Split "Chad Gurgdel & Amanda Smith" into ["chad", "gurgdel", "amanda", "smith"]
  return cellText.toLowerCase().replace(/[&,]+/g, ' ').split(/\s+/).filter(function(w) { return w.length > 0; });
}

function cellContainsFuzzy(cellWords, name) {
  for (var i = 0; i < cellWords.length; i++) {
    if (namesClose(cellWords[i], name)) return true;
  }
  return false;
}

// ═══ SHARED MATCHING ═══

function findMatchRow(householdData, firstName, lastName, tagsCol) {
  // Returns sheet row number (2-indexed) or -1

  // Pass 1: Exact — both first and last found as substrings
  var fullMatches = [];
  if (lastName) {
    for (var r = 0; r < householdData.length; r++) {
      if (tagsCol !== undefined && tagsCol !== -1 && householdData[r].tags && householdData[r].tags.indexOf('not-on-list') !== -1) continue;
      var cell = householdData[r].cell;
      if (cell.indexOf(firstName) !== -1 && cell.indexOf(lastName) !== -1) {
        fullMatches.push(householdData[r].row);
      }
    }
    if (fullMatches.length === 1) return fullMatches[0];
  }

  // Pass 2: Fuzzy — both first and last fuzzy-match words in cell
  var fuzzyMatches = [];
  if (lastName) {
    for (var r = 0; r < householdData.length; r++) {
      if (tagsCol !== undefined && tagsCol !== -1 && householdData[r].tags && householdData[r].tags.indexOf('not-on-list') !== -1) continue;
      var words = householdData[r].words;
      if (cellContainsFuzzy(words, firstName) && cellContainsFuzzy(words, lastName)) {
        fuzzyMatches.push(householdData[r].row);
      }
    }
    if (fuzzyMatches.length === 1) return fuzzyMatches[0];
  }

  // Pass 3: First name exact, unique
  var firstExact = [];
  for (var r = 0; r < householdData.length; r++) {
    if (tagsCol !== undefined && tagsCol !== -1 && householdData[r].tags && householdData[r].tags.indexOf('not-on-list') !== -1) continue;
    var words = householdData[r].words;
    // Check first word of each name chunk (split by & or ,)
    var chunks = householdData[r].cell.split(/[&,]+/);
    for (var c = 0; c < chunks.length; c++) {
      var chunkFirst = chunks[c].trim().split(/\s+/)[0];
      if (chunkFirst === firstName) {
        firstExact.push(householdData[r].row);
        break;
      }
    }
  }
  if (firstExact.length === 1) return firstExact[0];

  // Pass 4: First name fuzzy, unique
  var firstFuzzy = [];
  for (var r = 0; r < householdData.length; r++) {
    if (tagsCol !== undefined && tagsCol !== -1 && householdData[r].tags && householdData[r].tags.indexOf('not-on-list') !== -1) continue;
    var chunks = householdData[r].cell.split(/[&,]+/);
    for (var c = 0; c < chunks.length; c++) {
      var chunkFirst = chunks[c].trim().split(/\s+/)[0];
      if (namesClose(chunkFirst, firstName)) {
        firstFuzzy.push(householdData[r].row);
        break;
      }
    }
  }
  if (firstFuzzy.length === 1) return firstFuzzy[0];

  return -1;
}

function buildHouseholdList(sheet, householdCol, tagsCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numCols = sheet.getLastColumn();
  var dataRange = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var list = [];
  for (var r = 0; r < dataRange.length; r++) {
    var cellRaw = dataRange[r][householdCol].toString().toLowerCase();
    if (!cellRaw) continue;
    list.push({
      row: r + 2,
      cell: cellRaw,
      words: extractWords(cellRaw),
      tags: tagsCol !== undefined ? dataRange[r][tagsCol].toString() : '',
      data: dataRange[r]
    });
  }
  return list;
}

// ═══ POST — RSVP SUBMISSION ═══

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

  var tagsCol = headerMap['tags'];
  var list = buildHouseholdList(sheet, householdCol, tagsCol);
  var matchRow = findMatchRow(list, firstName, lastName, tagsCol);

  if (matchRow !== -1) {
    return updateRow(sheet, headerMap, data, matchRow);
  }

  return appendNewRow(sheet, headers, headerMap, data, 'not-on-list');
}

// ═══ UPDATE MATCHED ROW ═══

function updateRow(sheet, headerMap, data, matchRow) {
  var fieldsToUpdate = [
    'rsvp_status', 'attending',
    'adults_attending', 'kids_attending',
    'food_choices', 'dietary_notes', 'drinkers',
    'notes', 'ceremony_rsvp', 'ceremony_dinner'
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

// ═══ APPEND UNMATCHED ROW (5-row gap) ═══

function appendNewRow(sheet, headers, headerMap, data, tag) {
  var row = headers.map(function(header) {
    var key = header.toString().trim().toLowerCase();
    return data[key] !== undefined ? data[key] : '';
  });

  var tagsCol = headerMap['tags'];
  if (tagsCol !== undefined) {
    row[tagsCol] = tag;
  }

  // Find last "real" row (not tagged "not-on-list"), then append 5 rows below
  var lastReal = 1;
  if (tagsCol !== undefined) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var allTags = sheet.getRange(2, tagsCol + 1, lastRow - 1, 1).getValues();
      for (var t = 0; t < allTags.length; t++) {
        if (allTags[t][0].toString().indexOf('not-on-list') === -1) {
          lastReal = t + 2;
        }
      }
    }
  }

  var targetRow = Math.max(lastReal + 6, sheet.getLastRow() + 1);
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok', matched: false, tag: tag }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══ GET — LOOKUP / STATUS ═══

function doGet(e) {
  var action = (e.parameter.action || '').trim();

  if (action === 'lookup') {
    return lookupGuest(e);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'RSVP webhook is live (fuzzy match)' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookupGuest(e) {
  var first = (e.parameter.first || '').trim().toLowerCase();
  var last = (e.parameter.last || '').trim().toLowerCase();

  if (!first) {
    return ContentService
      .createTextOutput(JSON.stringify({ found: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerMap = {};
  for (var i = 0; i < headers.length; i++) {
    headerMap[headers[i].toString().trim().toLowerCase()] = i;
  }

  var householdCol = headerMap['household'];
  var ceremonyCol = headerMap['ceremony_count'];
  if (householdCol === undefined) {
    return ContentService
      .createTextOutput(JSON.stringify({ found: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var tagsCol = headerMap['tags'];
  var list = buildHouseholdList(sheet, householdCol, tagsCol);
  var matchRow = findMatchRow(list, first, last, tagsCol);

  if (matchRow === -1) {
    return ContentService
      .createTextOutput(JSON.stringify({ found: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Find the matched entry to read ceremony column
  var ceremonyInvited = false;
  for (var i = 0; i < list.length; i++) {
    if (list[i].row === matchRow) {
      if (ceremonyCol !== undefined) {
        var rawVal = list[i].data[ceremonyCol];
        ceremonyInvited = rawVal === true || rawVal === 'TRUE' || rawVal === 1 || rawVal === '1' || rawVal === 'yes' || rawVal === 'Yes';
      }
      break;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ found: true, ceremony_invited: ceremonyInvited }))
    .setMimeType(ContentService.MimeType.JSON);
}
