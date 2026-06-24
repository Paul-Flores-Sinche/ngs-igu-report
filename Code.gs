/**
 * NGS NT – IGU Daily Test Report · Google Apps Script backend
 *
 * Deploy as:  Web App → Execute as: Me → Who has access: Anyone
 * After each change, create a NEW deployment (don't edit existing) and update
 * the SHEET_URL constant in index.html with the new URL.
 *
 * Report approach: builds an HTML string → Utilities.newBlob() → DriveApp.createFile().
 * Saves as .html (opens in browser, printable as PDF). No DocumentApp or external API calls.
 */

const CONFIG = {
  FOLDER_ID : '12YgmAFL5sYvTqwgD3MYsuOlbxU6K_F1B',
  SHEET_NAME: 'IGU Reports v2',   // new sheet, created fresh with the 83 correct headers on first run; old "IGU Reports" is left untouched
};

// ─── WEB APP ENTRY POINTS ─────────────────────────────────────────────────────

function doPost(e) {
  try {
    console.log('doPost triggered, payload size:', e.postData.contents.length);
    const data = JSON.parse(e.postData.contents);
    console.log('Parsed OK — date:', data.date, '| line:', data.productionLine);

    const pdfFile = generateAndSavePDF(data);
    console.log('PDF saved:', pdfFile.getUrl());

    logToSheet(data, pdfFile.getUrl());
    console.log('Sheet logged');

    return jsonOut({ status: 'ok', driveUrl: pdfFile.getUrl(), fileName: pdfFile.getName() });
  } catch (err) {
    console.error('doPost FAILED:', err.toString(), err.stack);
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getReports') {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      if (!sheet || sheet.getLastRow() < 2) {
        return jsonOut({ status: 'ok', reports: [] });
      }
      const tz = Session.getScriptTimeZone();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const col = function(name) { return headers.indexOf(name); };
      const fmt = function(v) {
        if (!v || v === '') return '';
        if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm');
        return String(v);
      };
      const fmtDate = function(v) {
        if (!v || v === '') return '';
        if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
        return String(v);
      };
      const reports = rows.map(function(row, i) {
        return {
          rowIndex:       i + 2,
          timestamp:      fmt(row[col('Timestamp')]),
          date:           fmtDate(row[col('Date')]),
          shiftStart:     String(row[col('Shift Start')] || ''),
          productionLine: String(row[col('Production Line')] || ''),
          completedBy:    String(row[col('Completed By')] || ''),
          overallResult:  String(row[col('Overall Result')] || ''),
          driveUrl:       String(row[col('Drive URL')] || ''),
          approvedBy:     String(row[col('Approved By')] || ''),
          approvedAt:     fmt(row[col('Approved At')]),
        };
      });
      return jsonOut({ status: 'ok', reports: reports });
    } catch(err) {
      return jsonOut({ status: 'error', message: err.toString() });
    }
  }

  if (action === 'approve') {
    try {
      const rowIndex = parseInt(e.parameter.rowIndex);
      const supervisorName = e.parameter.supervisorName || '';
      if (!rowIndex || !supervisorName) {
        return jsonOut({ status: 'error', message: 'Missing rowIndex or supervisorName' });
      }
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const approvedByCol = headers.indexOf('Approved By') + 1;
      const approvedAtCol = headers.indexOf('Approved At') + 1;
      if (!approvedByCol || !approvedAtCol) {
        return jsonOut({ status: 'error', message: 'Approval columns not found — redeploy required' });
      }
      sheet.getRange(rowIndex, approvedByCol).setValue(supervisorName);
      sheet.getRange(rowIndex, approvedAtCol).setValue(new Date());
      return jsonOut({ status: 'ok' });
    } catch(err) {
      return jsonOut({ status: 'error', message: err.toString() });
    }
  }

  return jsonOut({ status: 'ok', message: 'NGS IGU Report API running' });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GOOGLE SHEETS LOGGING ────────────────────────────────────────────────────

function logToSheet(data, driveUrl) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = [
      'Timestamp','Date','Applicable Cert','Shift Start','Shift Finish','Time Finished','Total Units','Production Line',
      'Completed By','Overall Result','Pass Count','Fail Count',
      'S3.1','S3.2','S3.3','S3.4','S3.5',
      'S4.1','S4.2','S4.2 Trace','S4.3','S4.4','S4.5','S4.6','S4.7',
      'S5','S5.1','S5.2','S5.3','S5.4','S5.5',
      'S6','S6.1 Primary','S6.1 Secondary','S6.2','S6.3','S6.4','S6.5','S6.6','S6.7','S6.8',
      'S7','S7.1','S7.2','S7.3',
      'S8.1 PF','S8.1 Act','S8.2 Act',
      'S9 Act','S9.1','S9.2','S9.3','S9.4','S9.5','S9.6','S9.7','S9.8','S9.9',
      'S10','S10.1','S10.2',
      'Spacer Manufacturer','Spacer Batch','Spacer Product','Spacer Size',
      'Spacer Wall','Spacer Qty','Spacer Grade','Spacer Finish','Spacer Receipt',
      'Desic. Manufacturer','Desic. Batch','Desic. Product',
      'Mol Sieve','Desic. Receipt','Desic. Expiry',
      'Sealant Manufacturer','Sealant Batch','Sealant Product',
      'Sealant Receipt','Sealant Expiry',
      'Desic Fill Time','Desic Assem Time',
      'Drive URL','Approved By','Approved At',
    ];
    sheet.appendRow(headers);
    const hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground('#0b2e40').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(), data.date||'', data.applicableCert||'',
    data.shiftStart||'', data.shiftFinish||'', data.timeFinished||'',
    data.totalUnits||'', data.productionLine||'', data.completedBy||'',
    data.overallResult||'', data.passCount||0, data.failCount||0,
    data.s31||'', data.s32||'', data.s33||'', data.s34||'', data.s35||'',
    data.s41||'', data.s42||'', data.s4_trace_confirmed||'', data.s43||'', data.s44||'', data.s45||'', data.s46||'', data.s47||'',
    data.s51||'', data.s52||'', data.s53||'', data.s54||'', data.s55||'', data.s56||'',
    data.s61||'', data.s62||'', data.s62b||'', data.s63||'', data.s64||'', data.s65||'',
    data.s66||'', data.s67||'', data.s68||'', data.s69||'',
    data.s71||'', data.s72||'', data.s73||'', data.s74||'',
    data.s81||'', data.s81_act||'', data.s82||'',
    data.s91||'', data.s92||'', data.s93||'', data.s94||'', data.s95||'',
    data.s96||'', data.s97||'', data.s98||'', data.s99||'', data.s910||'',
    data.s101||'', data.s102||'', data.s103||'',
    data.s41_manufacturer||'', data.s41_batch||'', data.s41_product||'', data.s41_size||'',
    data.s41_wall_thickness||'', data.s41_quantity||'', data.s41_grade||'', data.s41_finish||'', data.s41_date_receipt||'',
    data.s51_manufacturer||'', data.s51_batch||'', data.s51_product||'',
    data.s51_mol_sieve||'', data.s51_date_receipt||'', data.s51_expiry||'',
    data.s61_manufacturer||'', data.s61_batch||'', data.s61_product||'',
    data.s61_date_receipt||'', data.s61_expiry||'',
    data.desiccantFillTime||'', data.desiccantAssemTime||'',
    driveUrl||'', '', '',
  ]);
}

// ─── REPORT GENERATION ────────────────────────────────────────────────────────
//
// Requires the Drive API advanced service:
//   Apps Script editor → left sidebar "Services" (+) → Drive API → Add
//
function generateAndSavePDF(data) {
  const tz         = Session.getScriptTimeZone();
  const dateStr    = data.date || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const safeDate   = dateStr.replace(/\//g, '-');
  const safeAuthor = (data.completedBy || 'report').substring(0, 20).replace(/[^a-zA-Z0-9]/g, '_');
  const fileName   = 'IGU_Report_' + safeDate + '_' + safeAuthor;

  const html = buildHtmlReport(data, dateStr, tz);
  console.log('HTML built, length:', html.length);

  // Upload the HTML blob and ask Drive to convert it to a native Google Doc.
  // The resulting file opens and renders correctly — no raw code shown.
  const blob = Utilities.newBlob(html, MimeType.HTML, fileName + '.html');
  // Drive advanced service v3: create() replaces v2's insert(); the HTML blob is
  // auto-converted to a Google Doc because the target mimeType is a Google type.
  const meta = Drive.Files.create(
    {
      name    : fileName,
      mimeType: 'application/vnd.google-apps.document',
      parents : [CONFIG.FOLDER_ID],
    },
    blob
  );
  console.log('Doc created:', meta.id);

  return DriveApp.getFileById(meta.id);
}

// ─── HTML REPORT BUILDER ──────────────────────────────────────────────────────

function buildHtmlReport(data, dateStr, tz) {
  const generated = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  const passCount = parseInt(data.passCount) || 0;
  const failCount = parseInt(data.failCount) || 0;
  const total     = passCount + failCount;
  const pct       = total > 0 ? Math.round(passCount / total * 100) : 0;
  const hasFails  = failCount > 0;

  // ── Utility helpers ────────────────────────────────────────────────────────
  function esc(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }
  function val(s) { return esc(s) || '—'; }

  // ── Badge renderers ────────────────────────────────────────────────────────
  const BADGE = 'display:inline-block;font-weight:700;font-size:10px;padding:2px 10px;border-radius:3px;color:#fff;';
  function pfBadge(r) {
    if (!r || r === 'Not checked')
      return '<span style="color:#aaa;font-size:10px;">—</span>';
    if (r === 'Pass')
      return '<span style="' + BADGE + 'background:#27AE60;">&#10003; PASS</span>';
    if (r === 'Fail')
      return '<span style="' + BADGE + 'background:#C0392B;">&#10007; FAIL</span>';
    return '<span style="color:#888;font-size:10px;">' + esc(r) + '</span>';
  }
  function actBadge(r) {
    if (!r || r === 'Not completed')
      return '<span style="color:#aaa;font-size:10px;">—</span>';
    if (r === 'Yes')
      return '<span style="' + BADGE + 'background:#27AE60;">&#10003; YES</span>';
    if (r === 'No')
      return '<span style="' + BADGE + 'background:#C0392B;">&#10007; NO</span>';
    return '<span style="color:#888;font-size:10px;">' + esc(r) + '</span>';
  }

  // ── Shared cell styles ─────────────────────────────────────────────────────
  const NUM_TD  = 'width:40px;text-align:center;font-weight:700;font-size:11px;' +
                  'color:#1C2E3D;background:#f0f4f6;border-right:1px solid #dce8ed;padding:7px 4px;';
  const TXT_TD  = 'font-size:11px;padding:7px 10px;color:#1a2a36;';
  const BADGE_TD= 'width:90px;text-align:center;border-left:1px solid #dce8ed;padding:6px;';

  // ── Row builders ───────────────────────────────────────────────────────────
  function secBanner(num, title, std) {
    return '<tr><td colspan="3" style="background:#1C2E3D;padding:7px 12px;">' +
      '<span style="display:inline-block;background:#2E86AB;color:#fff;font-weight:700;font-size:11px;' +
      'width:24px;height:24px;line-height:24px;text-align:center;border-radius:3px;' +
      'margin-right:8px;vertical-align:middle;">' + esc(num) + '</span>' +
      '<span style="color:#fff;font-weight:700;font-size:12px;vertical-align:middle;">' + esc(title) + '</span>' +
      (std ? '<span style="float:right;font-size:9px;color:rgba(255,255,255,0.45);line-height:24px;">' + esc(std) + '</span>' : '') +
      '</td></tr>';
  }

  function rowPF(num, title, pfResult) {
    return '<tr>' +
      '<td style="' + NUM_TD + '">' + esc(num) + '</td>' +
      '<td style="' + TXT_TD + '">' + esc(title) + '</td>' +
      '<td style="' + BADGE_TD + '">' + pfBadge(pfResult) + '</td>' +
      '</tr>';
  }

  function rowAct(num, title, actResult) {
    return '<tr>' +
      '<td style="' + NUM_TD + '">' + esc(num) + '</td>' +
      '<td style="' + TXT_TD + '" colspan="2">' + esc(title) +
      '&emsp;' + actBadge(actResult) + '</td>' +
      '</tr>';
  }

  // Detail sub-table: pairs is array of up to 3 [label, value] pairs per row
  function detailRow(pairs) {
    let cells = '';
    for (let i = 0; i < 3; i++) {
      if (pairs[i]) {
        cells +=
          '<td width="17%" style="font-size:9px;font-weight:700;color:#2E86AB;text-transform:uppercase;' +
          'letter-spacing:0.5px;background:#f0f6f9;padding:4px 8px;' +
          'border:1px solid #dce8ed;">' + esc(pairs[i][0]) + '</td>' +
          '<td width="17%" style="font-size:10px;padding:4px 8px;color:#1a2a36;border:1px solid #dce8ed;">' +
          val(pairs[i][1]) + '</td>';
      } else {
        cells += '<td colspan="2" style="border:1px solid #dce8ed;"></td>';
      }
    }
    return '<tr>' + cells + '</tr>';
  }

  function detailBlock(rows) {
    return '<tr><td colspan="3" style="padding:4px 10px 8px 50px;background:#f8fafb;">' +
      '<table width="100%" cellpadding="4" cellspacing="0" border="1" style="border-collapse:collapse;font-size:10px;">' + rows.join('') + '</table>' +
      '</td></tr>';
  }

  function photoRow(dataUrl, caption) {
    if (!dataUrl || dataUrl.length < 200) return '';
    return '<tr><td colspan="3" style="padding:6px 10px 8px 50px;background:#f8fafb;">' +
      '<img src="' + dataUrl + '" style="max-width:200px;max-height:150px;' +
      'border:1px solid #b8c4cc;border-radius:3px;">' +
      '<div style="font-size:9px;color:#888;margin-top:3px;">&uarr; ' + esc(caption) + '</div>' +
      '</td></tr>';
  }

  function metaCell(label, value) {
    return '<td style="padding:7px 14px;border-right:1px solid #dce8ed;">' +
      '<div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1px;' +
      'color:#5a7a8a;margin-bottom:2px;">' + esc(label) + '</div>' +
      '<div style="font-size:12px;color:#1a2a36;font-weight:600;">' + val(value) + '</div>' +
      '</td>';
  }

  function rowNotes(notes, sig) {
    if (!notes && !sig) return '';
    return '<tr style="background:#f9f9f9;">' +
      '<td colspan="2" style="padding:4px 8px;font-size:10px;color:#555;font-style:italic;width:70%"><b>Notes:</b> ' + (notes || '—') + '</td>' +
      '<td colspan="2" style="padding:4px 8px;font-size:10px;color:#555;width:30%"><b>Signature:</b> ' + (sig || '—') + '</td>' +
      '</tr>';
  }

  function cbList(ids, labels, data) {
    return ids.map((id, i) => (['Yes', 'true', true, 'on'].indexOf(data[id]) !== -1 ? '☑' : '☐') + ' ' + labels[i]).join('<br>');
  }

  // ── Derived values ─────────────────────────────────────────────────────────
  const overallBg  = hasFails ? '#fdecea' : '#e8f8ef';
  const overallClr = hasFails ? '#C0392B' : '#27AE60';
  const overallBdr = hasFails ? '#C0392B' : '#27AE60';

  const TBL  = 'style="width:100%;border-collapse:collapse;"';
  const STBL = 'style="width:100%;border-collapse:collapse;border:1px solid #b8c4cc;margin-bottom:14px;"';

  // ── Sealant detail block (item 6 — origin/shelf life) ─────────────────────
  const sealantDetail = (data.s61_manufacturer || data.s61_batch || data.s61_product)
    ? detailBlock([
        detailRow([['Manufacturer',data.s61_manufacturer],['Batch No.',data.s61_batch],['Product / Code',data.s61_product]]),
        detailRow([['Date of Receipt',data.s61_date_receipt],['Expiry Date',data.s61_expiry],null]),
      ])
    : '';

  // ── Desiccant exposure label ───────────────────────────────────────────────
  const exposureLabel = 'Desiccant Exposure: Maximum 2 hours exposure to atmosphere (recommended 40 min). Time of filling and assembly recorded.' +
    ' [Fill: ' + val(data.desiccantFillTime) + ' → Assembly: ' + val(data.desiccantAssemTime) + ']';

  // ══════════════════════════════════════════════════════════════════════════
  // HTML assembly
  // ══════════════════════════════════════════════════════════════════════════
  return '<!DOCTYPE html>\n' +
'<html lang="en">\n' +
'<head>\n' +
'<meta charset="UTF-8">\n' +
'<title>IGU Daily Test Report – ' + esc(dateStr) + '</title>\n' +
'<style>\n' +
'* { box-sizing:border-box; margin:0; padding:0; }\n' +
'body { font-family:Arial,sans-serif; font-size:11px; color:#222; background:#fff; padding:16px; max-width:920px; margin:0 auto; }\n' +
'@media print { body{padding:0;} .nobreak{page-break-inside:avoid;} }\n' +
'</style>\n' +
'</head>\n' +
'<body>\n\n' +

// ══ DOCUMENT HEADER ═══════════════════════════════════════════════════════
'<table ' + TBL + ' style="background:#1C2E3D;border-radius:6px 6px 0 0;margin-bottom:0;">\n' +
'<tr>\n' +
'  <td style="padding:14px 18px;width:88px;border-right:1px solid rgba(255,255,255,0.12);">\n' +
'    <div style="font-size:26px;font-weight:900;color:#fff;letter-spacing:-1px;line-height:1;">NGS</div>\n' +
'    <div style="font-size:7px;color:rgba(255,255,255,0.4);letter-spacing:2px;margin-top:2px;text-transform:uppercase;">Northern</div>\n' +
'    <div style="font-size:7px;color:rgba(255,255,255,0.4);letter-spacing:2px;text-transform:uppercase;">Glass</div>\n' +
'  </td>\n' +
'  <td style="padding:14px 18px;text-align:center;">\n' +
'    <div style="font-size:17px;font-weight:700;color:#fff;letter-spacing:1px;">Daily Test Report Checklist</div>\n' +
'    <div style="font-size:11px;color:rgba(255,255,255,0.65);margin-top:4px;">IGU – Aluminium Spacer Bar System</div>\n' +
'    <div style="font-size:9px;color:rgba(255,255,255,0.4);margin-top:3px;letter-spacing:3px;">AS 4666:2012</div>\n' +
'  </td>\n' +
'  <td style="padding:14px 18px;text-align:right;width:192px;border-left:1px solid rgba(255,255,255,0.12);">\n' +
'    <div style="font-size:9px;color:rgba(255,255,255,0.38);font-family:monospace;line-height:1.9;">TRF-IGU-Test-Report<br>Alum-Spacer-Bar</div>\n' +
'    <div style="font-size:14px;font-weight:700;color:#27AE60;margin-top:4px;letter-spacing:1px;">CSi ID 7709</div>\n' +
'    <div style="font-size:9px;color:rgba(255,255,255,0.5);margin-top:4px;">PAS-Mark' +
'      <span style="background:#27AE60;color:#fff;font-weight:700;padding:1px 6px;border-radius:2px;margin-left:4px;">A</span></div>\n' +
'  </td>\n' +
'</tr>\n' +
'</table>\n\n' +

// ══ COMPANY STRIP ════════════════════════════════════════════════════════
'<table ' + TBL + ' style="background:#243a4d;margin-bottom:0;">\n' +
'<tr>\n' +
'  <td style="padding:5px 18px;font-size:10px;color:rgba(255,255,255,0.65);">Company:&nbsp;<strong style="color:#fff;">Northern Glass Solutions Pty Ltd</strong></td>\n' +
'  <td style="padding:5px 18px;font-size:10px;color:rgba(255,255,255,0.65);border-left:1px solid rgba(255,255,255,0.1);">Location:&nbsp;<strong style="color:#fff;">84 O\'Sullivan Cres, East Arm, NT 0822</strong></td>\n' +
'  <td style="padding:5px 18px;font-size:10px;color:rgba(255,255,255,0.65);border-left:1px solid rgba(255,255,255,0.1);text-align:right;">Standard:&nbsp;<strong style="color:#fff;">AS 4666:2012</strong></td>\n' +
'</tr>\n' +
'</table>\n\n' +

// ══ META FIELDS ══════════════════════════════════════════════════════════
'<table ' + TBL + ' style="border:1px solid #b8c4cc;border-top:none;margin-bottom:0;">\n' +
'<tr style="background:#f0f4f6;">' +
  metaCell('Date', dateStr) +
  metaCell('Applicable Certificate', data.applicableCert) +
  metaCell('Shift Start', data.shiftStart) +
  metaCell('Shift Finish', data.shiftFinish) +
  metaCell('Total Units', data.totalUnits) +
'</tr>\n' +
'<tr style="background:#f8fafb;">' +
  metaCell('Production Line', data.productionLine) +
'  <td colspan="4" style="padding:7px 14px;">' +
'    <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#5a7a8a;margin-bottom:2px;">Completed By</div>' +
'    <div style="font-size:12px;color:#1a2a36;font-weight:600;">' + val(data.completedBy) + '</div>' +
'  </td>\n' +
'</tr>\n' +
'</table>\n\n' +

// ══ SUMMARY STRIP ════════════════════════════════════════════════════════
'<table ' + TBL + ' style="border:1px solid #b8c4cc;border-top:none;margin-bottom:16px;">\n' +
'<tr>\n' +
'  <td style="text-align:center;padding:10px 14px;border-right:1px solid #dce8ed;">' +
'    <div style="font-size:26px;font-weight:700;color:#2E86AB;">' + total + '</div>' +
'    <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-top:2px;">Items Checked</div>' +
'  </td>\n' +
'  <td style="text-align:center;padding:10px 14px;border-right:1px solid #dce8ed;">' +
'    <div style="font-size:26px;font-weight:700;color:#27AE60;">' + passCount + '</div>' +
'    <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-top:2px;">Pass</div>' +
'  </td>\n' +
'  <td style="text-align:center;padding:10px 14px;border-right:1px solid #dce8ed;">' +
'    <div style="font-size:26px;font-weight:700;color:#C0392B;">' + failCount + '</div>' +
'    <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-top:2px;">Fail</div>' +
'  </td>\n' +
'  <td style="text-align:center;padding:10px 14px;border-right:1px solid #dce8ed;">' +
'    <div style="font-size:26px;font-weight:700;color:#1C2E3D;">' + pct + '%</div>' +
'    <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-top:2px;">Complete</div>' +
'  </td>\n' +
'  <td style="text-align:center;padding:10px 14px;">' +
'    <div style="display:inline-block;padding:8px 22px;border-radius:4px;font-weight:700;font-size:13px;' +
'background:' + overallBg + ';color:' + overallClr + ';border:2px solid ' + overallBdr + ';">' +
esc(data.overallResult || '—') + '</div>' +
'  </td>\n' +
'</tr>\n' +
'</table>\n\n' +

// ══ SECTION 1 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('1','IGU Sampling Requirements','AS 4666:2012 · §5.8.2 & §6.0') +
rowAct('1.1','Number of IGU units to sample from production lot (Table 5.1, AS 4666)',data.s1) +
'<tr><td colspan="3" style="padding:4px 14px 4px 50px;background:#f8fafb;font-size:10px;">' +
'<span style="font-weight:700;color:#2E86AB;text-transform:uppercase;letter-spacing:0.5px;">Units to check this shift:</span>&nbsp;' +
esc(data.s1_actual_units || '—') + '&emsp;' +
'<span style="font-weight:700;color:#2E86AB;text-transform:uppercase;letter-spacing:0.5px;">Lot size range:</span>&nbsp;' +
esc(data.s1_actual_lot || '—') +
'</td></tr>' +
rowNotes(data.s11_notes, data.s11_sig) +
'</table>\n\n' +

// ══ SECTION 2 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('2','Atmospheric Conditions','AS 4666:2012 · §5.2 — Every 4 hours') +
'<tr><td colspan="3" style="padding:5px 14px 5px 50px;background:#f8fafb;border-bottom:1px solid #dce8ed;">' +
'<span style="font-size:9px;font-weight:700;color:#2E86AB;text-transform:uppercase;letter-spacing:0.5px;">Shift Start</span>&nbsp;' +
'<span style="font-size:11px;color:#1a2a36;margin-right:20px;">' + val(data.shiftStart) + '</span>' +
'<span style="font-size:9px;font-weight:700;color:#2E86AB;text-transform:uppercase;letter-spacing:0.5px;">Shift Finish</span>&nbsp;' +
'<span style="font-size:11px;color:#1a2a36;margin-right:20px;">' + val(data.shiftFinish) + '</span>' +
'<span style="font-size:9px;font-weight:700;color:#2E86AB;text-transform:uppercase;letter-spacing:0.5px;">Total Units</span>&nbsp;' +
'<span style="font-size:11px;color:#1a2a36;">' + val(data.totalUnits) + '</span>' +
'</td></tr>' +
rowAct('2.1','Record Atmospheric Conditions at start of production shift and every four (4) hours until the end of production shift',data.s2) +
(function(){
  var rows = [1,2,3,4,5,6].map(function(i){
    var t=data['atm_time_'+i], c=data['atm_temp_'+i],
        h=data['atm_hum_'+i],  p=data['atm_pres_'+i];
    if(!t&&!c&&!h&&!p) return '';
    return '<tr>' +
      '<td width="25%" style="padding:3px 8px;border:1px solid #dce8ed;">' + val(t) + '</td>' +
      '<td width="25%" style="padding:3px 8px;border:1px solid #dce8ed;">' + val(c) + '</td>' +
      '<td width="25%" style="padding:3px 8px;border:1px solid #dce8ed;">' + val(h) + '</td>' +
      '<td width="25%" style="padding:3px 8px;border:1px solid #dce8ed;">' + val(p) + '</td>' +
      '</tr>';
  }).join('');
  if(!rows) return '';
  return '<tr><td colspan="3" style="padding:4px 10px 6px 50px;background:#f8fafb;">' +
    '<table width="100%" cellpadding="4" cellspacing="0" border="1" style="border-collapse:collapse;font-size:10px;">' +
    '<tr style="background:#f0f4f6;">' +
    '<th width="25%" style="text-align:left;border:1px solid #dce8ed;color:#5a7a8a;font-size:9px;text-transform:uppercase;letter-spacing:0.5px;">Time</th>' +
    '<th width="25%" style="text-align:left;border:1px solid #dce8ed;color:#5a7a8a;font-size:9px;text-transform:uppercase;letter-spacing:0.5px;">Temp (°C)</th>' +
    '<th width="25%" style="text-align:left;border:1px solid #dce8ed;color:#5a7a8a;font-size:9px;text-transform:uppercase;letter-spacing:0.5px;">Humidity (%)</th>' +
    '<th width="25%" style="text-align:left;border:1px solid #dce8ed;color:#5a7a8a;font-size:9px;text-transform:uppercase;letter-spacing:0.5px;">Pressure (hPa)</th>' +
    '</tr>' + rows +
    '</table></td></tr>';
})() +
rowNotes(data.s21_notes, data.s21_sig) +
'</table>\n\n' +

// ══ SECTION 3 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('3','Glazing Materials','AS 4666:2012 · §5.3.1–5.3.7') +
rowPF('3.1','GLASS: Compliance with order – Traceability – Types of Substrates',data.s31) +
rowNotes(data.s31_notes, data.s31_sig) +
rowPF('3.2','GLASS: Edge Characteristics. A visual inspection shall be conducted in accordance with Section 5.3.4 and Table 5.2',data.s32) +
rowNotes(data.s32_notes, data.s32_sig) +
rowPF('3.3','GLASS: Dimensional Properties: When measured in accordance with Appendix A, AS 4666 – THICKNESS tolerance monolithic glass see, Table 5.1, FLATNESS (Bow) see Table 5.3, HEIGHT/WIDTH, Refer to Section 5.3.5.2, Residential Units, Section 5.3.5.3, Commercial Units.',data.s33) +
rowNotes(data.s33_notes, data.s33_sig) +
rowPF('3.4','GLASS: Cleanliness – Dryness, Stains/Films, Fingerprints',data.s34) +
rowNotes(data.s34_notes, data.s34_sig) +
rowPF('3.5','GLASS: A visual inspection shall be conducted for Scratches, Blemishes, Marks and Inclusions. Refer to AS 4666, Table 5.5 and Table 5.6',data.s35) +
rowNotes(data.s35_notes, data.s35_sig) +
'</table>\n\n' +

// ══ SECTION 4 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('4','Spacer Bar & Connectors','AS 4666:2012 · §5.4.1–5.4.8') +
rowPF('4.1','Spacer Bar: Dimensional Properties & Stock Record',data.s41) +
(
  (data.s41_manufacturer || data.s41_batch || data.s41_product)
  ? detailBlock([
      detailRow([['Manufacturer',data.s41_manufacturer],['Batch No.',data.s41_batch],['Product / Code',data.s41_product]]),
      detailRow([['Size (L×W×H)',data.s41_size],['Wall Thickness',data.s41_wall_thickness],['Quantity',data.s41_quantity]]),
      detailRow([['Grade',data.s41_grade],['Finish / Colour',data.s41_finish],['Date of Receipt',data.s41_date_receipt]]),
    ])
  : ''
) +
photoRow(data.photo41,'Spacer Bar Label Photo') +
rowNotes(data.s41_notes, data.s41_sig) +
rowPF('4.2','Spacer Bar: Customer Stock Record Traceability',data.s42) +
rowNotes(data.s42_notes, data.s42_sig) +
rowPF('4.3','Rigid Spacer Bar: Visual Inspection',data.s43) +
rowNotes(data.s43_notes, data.s43_sig) +
rowPF('4.4','Corner Keys & Straight Line Connectors',data.s44) +
rowNotes(data.s44_notes, data.s44_sig) +
rowPF('4.5','Spacer Bar Cleanliness: Free from Dirt, Oil, Grease, Contaminants',data.s45) +
rowNotes(data.s45_notes, data.s45_sig) +
rowPF('4.6','Dimensional Properties: Compliance with Manufacturer Specifications',data.s46) +
rowNotes(data.s46_notes, data.s46_sig) +
rowPF('4.7','Post Assembly Inspection: No splitting, tearing or flare-out on bent corners',data.s47) +
rowNotes(data.s47_notes, data.s47_sig) +
'</table>\n\n' +

// ══ SECTION 5 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('5','Desiccant','AS 4666:2012 · §5.5.2–5.5.9') +
rowPF('5.1','Desiccant: Origin, Compliance with Order & Shelf Life',data.s51) +
(
  (data.s51_manufacturer || data.s51_batch || data.s51_product)
  ? detailBlock([
      detailRow([['Manufacturer',data.s51_manufacturer],['Batch No.',data.s51_batch],['Product / Code',data.s51_product]]),
      detailRow([['Molecular Sieve',data.s51_mol_sieve],['Date of Receipt',data.s51_date_receipt],['Expiry Date',data.s51_expiry]]),
    ])
  : ''
) +
photoRow(data.photo51,'Desiccant Label Photo') +
rowNotes(data.s51_notes, data.s51_sig) +
rowPF('5.2','Desiccant Traceability: Traceable back to supplier for each customer order',data.s52) +
rowNotes(data.s52_notes, data.s52_sig) +
rowPF('5.3','Desiccant Effectiveness / Fit for Purpose: Tested per manufacturer specification before each production run',data.s53) +
rowNotes(data.s53_notes, data.s53_sig) +
rowPF('5.4','Desiccant Volume: Adequate volume employed in spacer bar cavity',data.s54) +
rowNotes(data.s54_notes, data.s54_sig) +
rowPF('5.5',exposureLabel,data.s55) +
rowNotes(data.s55_notes, data.s55_sig) +
rowPF('5.6','Desiccant Suitability: Compatible with gas type used in IGU assembly (pore size < gas molecular sieve)',data.s56) +
rowNotes(data.s56_notes, data.s56_sig) +
'</table>\n\n' +

// ══ SECTION 6 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('6','Sealants','AS 4666:2012 · §5.6.1–5.6.10') +
rowPF('6.1','Sealant: Origin, Compliance with Order & Shelf Life',data.s61) +
sealantDetail +
photoRow(data.photo61,'Sealant Label Photo') +
rowNotes(data.s61_notes, data.s61_sig) +
// 6.1 — Traceability has two separate pf-rows (primary + secondary)
'<tr><td colspan="3" style="padding:4px 14px;background:#eef4f7;border-bottom:1px solid #dce8ed;">' +
'<span style="font-size:9px;font-weight:700;color:#1C2E3D;text-transform:uppercase;letter-spacing:0.8px;">6.2 — Sealant Traceability: Primary &amp; Secondary Sealant Records</span>' +
'</td></tr>' +
'<tr>' +
'<td style="' + NUM_TD + '">6.2</td>' +
'<td style="' + TXT_TD + '">Primary Sealant – Manufacturer / Batch / Expiry</td>' +
'<td style="' + BADGE_TD + '">' + pfBadge(data.s62) + '</td>' +
'</tr>' +
rowNotes(data.s62p_notes, data.s62p_sig) +
'<tr>' +
'<td style="' + NUM_TD + '">6.2b</td>' +
'<td style="' + TXT_TD + '">Secondary Sealant – Manufacturer / Batch / Expiry</td>' +
'<td style="' + BADGE_TD + '">' + pfBadge(data.s62b) + '</td>' +
'</tr>' +
rowNotes(data.s62f_notes, data.s62f_sig) +
rowPF('6.3','Minimum Sealant Dimensions: Verified to Table 5.7. Dual or Single Seal unit.',data.s63) +
rowNotes(data.s63_notes, data.s63_sig) +
rowPF('6.4','Sealant Compatibility with Glass: Confirmed as compatible and approved for use',data.s64) +
rowNotes(data.s64_notes, data.s64_sig) +
rowPF('6.5','Secondary Adhesion: To glass surface and spacer bar. Records maintained.',data.s65) +
rowNotes(data.s65_notes, data.s65_sig) +
rowPF('6.6','Mix Ratio: Two-part secondary sealant mix ratio tested. Records maintained.',data.s66) +
rowNotes(data.s66_notes, data.s66_sig) +
rowPF('6.7','Mix Homogeneity: Tested once per shift. Results recorded.',data.s67) +
rowNotes(data.s67_notes, data.s67_sig) +
rowPF('6.8','Sealant Cure Rate: Single or secondary sealants tested. Records maintained.',data.s68) +
rowNotes(data.s68_notes, data.s68_sig) +
rowPF('6.9','Sealant Storage: Stored per manufacturer specification.',data.s69) +
rowNotes(data.s69_notes, data.s69_sig) +
'</table>\n\n' +

// ══ SECTION 7 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('7','Gas / IGU Gassing','AS 4666:2012 · §5.7.1–5.7.6') +
rowPF('7.1','Gas: Origin, Compliance with Order & Gas Type',data.s71) +
rowNotes(data.s71_notes, data.s71_sig) +
rowPF('7.2','Gas Traceability: Traceable to supplier for each customer order',data.s72) +
rowNotes(data.s72_notes, data.s72_sig) +
rowPF('7.3','Desiccant & Gas Compatibility: Desiccant pore size smaller than gas molecular sieve size',data.s73) +
rowNotes(data.s73_notes, data.s73_sig) +
rowPF('7.4','Gas Fill Volume: Minimum 90% of IGU cavity volume. Calculated or measured per Appendix O.',data.s74) +
rowNotes(data.s74_notes, data.s74_sig) +
'</table>\n\n' +

// ══ SECTION 8 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('8','Compliance Markings','§4') +
rowPF('8.1','Compliance Marking: Company name/logo, CSi ID 7709, AS 4666, Date of manufacture',data.s81) +
'<tr><td colspan="4" style="padding:6px 8px;font-size:10px;line-height:1.6">' +
cbList(['m1','m2','m3','m4'],
['Company Name or Logo','Manufacturing Site Identifier (CSi ID 7709)',
'Manufacturing Standard: AS 4666','Date of Manufacture (Month and Year)'], data) +
'</td></tr>' +
rowNotes(data.s81_notes, data.s81_sig) +
rowAct('8.1','Compliance Marking — Activity Completed',data.s81_act) +
rowAct('8.2','Work Order / Customer Order Label Affixed to Finished Product',data.s82) +
'</table>\n\n' +

// ══ SECTION 9 ════════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('9','Final Product Check','AS 4666:2012 · §5.8.2–5.8.12') +
rowAct('9.1','Test Format: Random samples (Table 6.1) before shipment. Non-conforming product rectified or replaced.',data.s91) +
rowNotes(data.s91_notes, data.s91_sig) +
rowPF('9.2','Sealant Cured: Visual & mechanical hardness check. No uncured sealant on units ready for shipment.',data.s92) +
rowNotes(data.s92_notes, data.s92_sig) +
rowPF('9.3','Dimensional Properties: Height × Width, Squareness, Flatness per §5.8.4.2–5.8.4.4 and Table 5.8',data.s93) +
rowNotes(data.s93_notes, data.s93_sig) +
rowPF('9.4','Uniformity of Spacer Depth: Minimum sealant depth per Table 5.7',data.s94) +
rowNotes(data.s94_notes, data.s94_sig) +
rowPF('9.5','Voids in Sealants: Visual inspection. Random samples per Table 6.1.',data.s95) +
rowNotes(data.s95_notes, data.s95_sig) +
rowPF('9.6','Glass Alignment / Offset: Visual inspection. Random samples per Table 6.1.',data.s96) +
rowNotes(data.s96_notes, data.s96_sig) +
rowPF('9.7','Cleanliness: Max 1mm proud on face, max 2mm on edge, max 3mm on corner (±20mm).',data.s97) +
rowNotes(data.s97_notes, data.s97_sig) +
rowPF('9.8','Marks, Scratches & Inclusions: Per Table 5.5. IGU may have twice the amount of individual pane.',data.s98) +
rowNotes(data.s98_notes, data.s98_sig) +
rowPF('9.9','Moisture Content (Dew Point Test): Conducted on previous day production lot. Random samples per Table 6.1.',data.s99) +
rowNotes(data.s99_notes, data.s99_sig) +
rowPF('9.10','Storage & Handling: Support (§5.8.12.1), Angle (§5.8.12.2), Spacing/Interleaving (§5.8.12.3). Lifting to thickest face.',data.s910) +
rowNotes(data.s910_notes, data.s910_sig) +
'</table>\n\n' +

// ══ SECTION 10 ═══════════════════════════════════════════════════════════
'<table ' + STBL + '>' +
secBanner('10','Equipment Availability','Appendix A · §5') +
rowPF('10.1','Mandatory Equipment for Dimensional Properties',data.s101) +
'<tr><td colspan="4" style="padding:6px 8px;font-size:10px;line-height:1.6">' +
cbList(['eq0','eq1','eq2','eq3','eq4','eq5'],
['Point Micrometre* 0.01mm','Plate Micrometre* 0.01mm','Vernier Calliper* 0.01mm','150mm Steel Rule or Steel Tapered Rule* 0.5mm','Standard Measuring Tape* 1.0mm','Straight Edge*'], data) +
'</td></tr>' +
rowNotes(data.s101_notes, data.s101_sig) +
rowPF('10.2','Additional Equipment for Compliance Checks',data.s102) +
'<tr><td colspan="4" style="padding:6px 8px;font-size:10px;line-height:1.6">' +
cbList(['aeq0','aeq1','aeq2','aeq3','aeq4','aeq5','aeq6','aeq7','aeq8','aeq9','aeq10','aeq11','aeq12','aeq13','aeq14'],
['Durometer Type A2 (0-100 increments)','Teflon Mould (127mm x 38mm x 6mm)','Digital Thermometer','Stopwatch','Barometric Pressure - Humidity/Temperature Gauge/Datalogger','Dry Ice Maker (Snowpac Device)','Crucible to hold dry ice','Calculator','Digital Scales','One litre beaker with sight scale in 50ml graduations','Handheld ultra-violet light source','Delta T test kit','Sparklite (Used to test IG gas concentration)','Non-powdered latex gloves','Disposable paper towel'], data) +
'</td></tr>' +
rowNotes(data.s102_notes, data.s102_sig) +
rowPF('10.3','Equipment Calibrated and Functional: Within tolerance of 0.01mm. Records maintained.',data.s103) +
rowNotes(data.s103_notes, data.s103_sig) +
'</table>\n\n' +

// ══ FOOTER ═══════════════════════════════════════════════════════════════
'<table ' + TBL + ' style="border-top:2px solid #1C2E3D;margin-top:16px;">\n' +
'<tr>\n' +
'  <td style="padding:8px 0;font-size:9px;color:#555;">' +
'    <strong style="color:#1C2E3D;text-transform:uppercase;letter-spacing:1px;font-size:9px;">' +
'    PRINTED COPIES OF THIS DOCUMENT ARE UNCONTROLLED</strong><br>' +
'    TRF-IGU-Test-Report-Alum-Spacer-Bar &nbsp;&middot;&nbsp; Northern Glass Solutions Pty Ltd' +
'    &nbsp;&middot;&nbsp; CSi ID 7709 &nbsp;&middot;&nbsp; AS 4666:2012' +
'  </td>\n' +
'  <td style="padding:8px 0;font-size:9px;color:#888;text-align:right;white-space:nowrap;">' +
'    Generated: ' + generated +
'  </td>\n' +
'</tr>\n' +
'</table>\n\n' +

'</body>\n</html>';
}

// ─── DEBUG / TEST ─────────────────────────────────────────────────────────────

/**
 * Run from the Apps Script editor (▶ Run) to test the full doPost flow
 * without a browser form submission. Check View → Executions for logs.
 */
function testDoPost() {
  const mockData = {
    date: '2026-05-13', applicableCert: 'CERT-2026-001',
    shiftStart: '07:00', shiftFinish: '15:30',
    totalUnits: '45', productionLine: 'Line 1', completedBy: 'Paul Flores',
    overallResult: 'ALL PASS', passCount: 22, failCount: 0,
    s31:'Pass', s32:'Pass', s33:'Pass', s34:'Pass', s35:'Pass',
    s41:'Pass', s42:'Pass', s43:'Pass', s44:'Pass', s45:'Pass', s46:'Pass', s47:'Pass',
    s51:'Pass', s52:'Pass', s53:'Pass', s54:'Pass', s55:'Pass', s56:'N/A',
    s61:'Pass', s62:'Pass', s62b:'Pass', s63:'Pass', s64:'Pass',
    s65:'Pass', s66:'Pass', s67:'Pass', s68:'Pass', s69:'Pass',
    s71:'Pass', s72:'Pass', s73:'Pass', s74:'Pass',
    s81:'Pass', s81_act:'Yes', s82:'Yes',
    s91:'Yes', s92:'Pass', s93:'Pass', s94:'Pass', s95:'Pass',
    s96:'Pass', s97:'Pass', s98:'Pass', s99:'Pass', s910:'Pass',
    s101:'Pass', s102:'Pass', s103:'Pass',
    s41_manufacturer:'Aluminco', s41_batch:'BATCH-001', s41_product:'AlSB-16',
    s41_size:'6m × 16mm × 10mm', s41_wall_thickness:'1.5 mm',
    s41_quantity:'100', s41_grade:'6063-T5', s41_finish:'Mill Finish', s41_date_receipt:'2026-01-15',
    s51_manufacturer:'Grace', s51_batch:'DSC-789', s51_product:'4A Molecular Sieve',
    s51_mol_sieve:'4A', s51_date_receipt:'2026-01-15', s51_expiry:'2026-12-31',
    s61_manufacturer:'Tremco', s61_batch:'PIB-456', s61_product:'Tremflex 834',
    s61_date_receipt:'2026-02-01', s61_expiry:'2026-09-01',
    desiccantFillTime:'07:30', desiccantAssemTime:'08:10',
    photo41:'', photo51:'', photo61:'',
  };

  const result = doPost({ postData: { contents: JSON.stringify(mockData) } });
  const parsed = JSON.parse(result.getContent());
  console.log('testDoPost result:', JSON.stringify(parsed));

  if (parsed.status === 'ok') {
    SpreadsheetApp.getUi().alert('✅ Report created!\n\n' + parsed.driveUrl);
  } else {
    SpreadsheetApp.getUi().alert('❌ Error:\n\n' + parsed.message);
  }
}
