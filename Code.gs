const ID_COLUMN = 9; // Column I

/**
 * Finds the sheet row (1-based, including header) whose ID column (col 9)
 * matches the given id. Comparison is done as trimmed strings so that
 * stray whitespace never causes a false "not found".
 * Returns -1 if no matching row exists.
 */
function findRowById(sheet, id) {
  const targetId = String(id).trim();
  if (!targetId) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1; // no data rows, only header (or empty sheet)

  const ids = sheet.getRange(2, ID_COLUMN, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const cellId = String(ids[i][0]).trim();
    if (cellId === targetId) {
      return i + 2; // convert 0-based offset to actual sheet row
    }
  }
  return -1;
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let postData = {};
  try {
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Invalid JSON'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = postData.action || '';

  // ---------- DELETE ----------
  if (action === 'delete') {
    const id = postData.id || '';
    const row = findRowById(sheet, id);
    if (row !== -1) {
      sheet.deleteRow(row);
      return ContentService.createTextOutput(JSON.stringify({status: 'deleted', row: row}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({status: 'not_found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ---------- CREATE / UPDATE ----------
  // Accepts either { data: [ {...} ] } or a bare object { organization: ..., id: ... }
  const item = postData.data && postData.data[0] ? postData.data[0] : postData;

  const organization = item.organization || '';
  const contactName = item.contactName || '';
  const brand = item.brand || '';
  const model = item.model || '';
  const year = item.year || '';
  const frequency = item.frequency || '';
  const location = item.location || '';
  const date = item.date || '';
  const id = String(item.id || '').trim();

  const row = [organization, contactName, brand, model, year, frequency, location, date, id];

  if (id) {
    const existingRow = findRowById(sheet, id);
    if (existingRow !== -1) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
      return ContentService.createTextOutput(JSON.stringify({status: 'updated', row: existingRow}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  sheet.appendRow(row);

  return ContentService.createTextOutput(JSON.stringify({status: 'created'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}
