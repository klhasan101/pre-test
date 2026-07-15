/**
 * Google Apps Script endpoint for index.html.
 * اربط هذا السكربت بجدول Google Sheets المطلوب ثم انشره كتطبيق ويب.
 */
const SHEET_NAME = 'Responses';
const CODE_COUNTER_KEY = 'PARTICIPANT_CODE_COUNTER';

function doGet(e) {
  const callback = String((e.parameter && e.parameter.callback) || 'receiveParticipantCode')
    .replace(/[^a-zA-Z0-9_.$]/g, '');
  const response = { ok: false };

  if (e.parameter && e.parameter.action === 'nextCode') {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const properties = PropertiesService.getScriptProperties();
      const nextNumber = Number(properties.getProperty(CODE_COUNTER_KEY) || 0) + 1;
      properties.setProperty(CODE_COUNTER_KEY, String(nextNumber));
      response.ok = true;
      response.code = 'T' + nextNumber;
    } finally {
      lock.releaseLock();
    }
  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(response) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateResponseSheet_();
    const headers = getHeaders_(sheet, data);
    if (hasParticipantCode_(sheet, headers, data.participant_code)) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, duplicate: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const row = headers.map(function (header) {
      return Object.prototype.hasOwnProperty.call(data, header) ? data[header] : '';
    });
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, participant_code: data.participant_code }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function hasParticipantCode_(sheet, headers, participantCode) {
  if (!participantCode || sheet.getLastRow() < 2) return false;
  const codeColumn = headers.indexOf('participant_code') + 1;
  if (!codeColumn) return false;
  const codes = sheet.getRange(2, codeColumn, sheet.getLastRow() - 1, 1).getDisplayValues();
  return codes.some(function (row) {
    return row[0] === String(participantCode);
  });
}

function getOrCreateResponseSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function getHeaders_(sheet, data) {
  const lastColumn = sheet.getLastColumn();
  let headers = lastColumn
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(String)
    : [];

  const newHeaders = Object.keys(data).filter(function (key) {
    return headers.indexOf(key) === -1;
  });
  headers = headers.concat(newHeaders);

  if (newHeaders.length || sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return headers;
}
