/**
 * Chick-fil-A Ft. Totten — Monthly Marketing Plan Builder
 * Google Apps Script backend: saves each submitted plan to a Google Sheet
 * and emails the Operator for review & approval.
 *
 * SETUP
 *  1. Create a new Google Sheet (this will hold submitted plans).
 *  2. Extensions ▸ Apps Script. Delete any sample code and paste this file.
 *  3. Change OPERATOR_EMAIL below to your address.
 *  4. Limit permissions to THIS sheet only: Project Settings (gear icon)
 *     ▸ tick "Show appsscript.json manifest file", open appsscript.json,
 *     and replace its contents with the provided deploy/appsscript.json.
 *     This requests the narrow "spreadsheets.currentonly" scope (only the
 *     bound sheet) plus send-email — not access to all your Sheets/Drive.
 *  5. Deploy ▸ New deployment ▸ type "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  6. Copy the Web app URL (ends in /exec) and paste it into the
 *     app's CONFIG.sheetsEndpoint (see the SETUP GUIDE).
 *
 * NOTE: this script uses getActiveSpreadsheet(), which is compatible with
 * the spreadsheets.currentonly scope — it can only touch this one sheet.
 */

var OPERATOR_EMAIL = 'alan.thompson@cfafttotten.com';
var SHEET_NAME = 'Plans';

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var sheet = getSheet_();
  sheet.appendRow([
    new Date(),
    body.month || '',
    body.preparedBy || '',
    'Pending Review',
    body.plan || '',
    JSON.stringify(body.data || {})
  ]);

  MailApp.sendEmail({
    to: OPERATOR_EMAIL,
    subject: body.subject || 'New marketing plan submitted',
    body: (body.plan || '') +
      '\n\n— Open the Marketing Plan Builder to review, edit, and approve.',
    htmlBody: buildHtml_(body)
  });

  return json_({ ok: true });
}

/** Renders the submitted plan as a clean, branded HTML email. */
function buildHtml_(body) {
  var month = esc_(body.month || '');
  var by = esc_(body.preparedBy || 'Marketing Lead');
  var d = body.data || {};
  var objective = esc_(d.objective || '');

  var rows = String(body.plan || '').split('\n');
  var htmlLines = [];
  for (var i = 0; i < rows.length; i++) {
    var line = rows[i];
    var trimmed = line.replace(/\s+$/, '');
    if (trimmed === '') { htmlLines.push('<div style="height:10px"></div>'); continue; }

    // Indented list items ("  - ..." or "     ...")
    if (/^\s{2,}-\s/.test(line)) {
      htmlLines.push('<div style="margin:2px 0 2px 18px;font-size:14px;color:#333;line-height:1.5">&bull; ' + esc_(line.replace(/^\s*-\s/, '')) + '</div>');
      continue;
    }
    if (/^\s{4,}\S/.test(line)) {
      htmlLines.push('<div style="margin:2px 0 2px 30px;font-size:13px;color:#5B6770;line-height:1.5">' + esc_(line.trim()) + '</div>');
      continue;
    }
    // ALL-CAPS section heading (e.g. the title, "Team contests:")
    var m = trimmed.match(/^([A-Za-z][^:]*):\s*(.*)$/);
    if (/^[A-Z0-9 &—–-]+$/.test(trimmed) && trimmed.length > 3) {
      htmlLines.push('<div style="font-size:15px;font-weight:700;color:#004F71;margin:16px 0 6px">' + esc_(trimmed) + '</div>');
    } else if (m) {
      htmlLines.push('<div style="font-size:14px;color:#333;margin:3px 0;line-height:1.55"><strong style="color:#004F71">' + esc_(m[1]) + ':</strong> ' + esc_(m[2]) + '</div>');
    } else {
      htmlLines.push('<div style="font-size:14px;color:#333;margin:3px 0;line-height:1.55">' + esc_(trimmed) + '</div>');
    }
  }

  return '' +
    '<div style="margin:0;padding:24px 12px;background:#FCF7E7;font-family:Helvetica,Arial,sans-serif">' +
      '<div style="max-width:640px;margin:0 auto;background:#fff;border:1px solid #EEEDEB;border-radius:16px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06)">' +
        '<div style="background:#DD0033;padding:26px 30px;color:#fff">' +
          '<div style="font-size:11px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:rgba(255,255,255,.85)">Chick-fil-A Ft. Totten · Marketing Plan' + (month ? ' · ' + month : '') + '</div>' +
          (objective ? '<div style="font-size:22px;font-weight:700;line-height:1.25;margin-top:10px">' + objective + '</div>' : '') +
          '<div style="font-size:13px;color:rgba(255,255,255,.9);margin-top:10px">Prepared by ' + by + '</div>' +
        '</div>' +
        '<div style="padding:26px 30px">' + htmlLines.join('') + '</div>' +
        '<div style="padding:20px 30px;background:#FCF7E7;border-top:1px solid #EEEDEB;font-size:13px;color:#5B6770;line-height:1.55">' +
          'Open the Marketing Plan Builder and click <strong>Load latest submitted plan</strong> on the final step to review, edit, and record your decision.' +
        '</div>' +
      '</div>' +
    '</div>';
}

function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Returns the most recently submitted plan (its saved data blob),
 * so the tool can reload the latest plan for editing/approval.
 * Optional: append ?preparedBy=Name to fetch the latest for one person.
 */
function doGet(e) {
  var sheet = getSheet_();
  var rows = sheet.getDataRange().getValues();
  var who = (e && e.parameter && e.parameter.preparedBy || '').toLowerCase();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!who || String(rows[i][2]).toLowerCase() === who) {
      return json_({
        date: rows[i][0], month: rows[i][1], preparedBy: rows[i][2],
        status: rows[i][3], plan: rows[i][4], data: safeParse_(rows[i][5])
      });
    }
  }
  return json_({});
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Submitted', 'Month', 'Prepared by', 'Status', 'Plan (text)', 'Data (JSON)']);
  }
  return sheet;
}

function safeParse_(s) { try { return JSON.parse(s); } catch (err) { return {}; } }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
