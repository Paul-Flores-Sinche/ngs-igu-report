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
  SHEET_NAME: 'IGU Reports',
};

// ─── WEB APP ENTRY POINTS ─────────────────────────────────────────────────────

function doPost(e) {
  try {
    console.log('doPost triggered, payload size:', e.postData.contents.length);
    const data = JSON.parse(e.postData.contents);
    console.log('Parsed OK — date:', data.date, '| line:', data.productionLine);

    logToSheet(data);
    console.log('Sheet logged');

    const pdfFile = generateAndSavePDF(data);
    console.log('PDF saved:', pdfFile.getUrl());

    return jsonOut({ status: 'ok', driveUrl: pdfFile.getUrl(), fileName: pdfFile.getName() });
  } catch (err) {
    console.error('doPost FAILED:', err.toString(), err.stack);
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  return jsonOut({ status: 'ok', message: 'NGS IGU Report API running' });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GOOGLE SHEETS LOGGING ────────────────────────────────────────────────────

function logToSheet(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = [
      'Timestamp','Date','Shift Start','Shift Finish','Total Units','Production Line',
      'Completed By','Overall Result','Pass Count','Fail Count',
      'S3.1','S3.2','S3.3','S3.4','S3.5',
      'S4.1','S4.2','S4.3','S4.4','S4.5','S4.6','S4.7',
      'S5.1','S5.2','S5.3','S5.4','S5.5','S5.6',
      'S6.1','S6.2','S6.3','S6.4','S6.5','S6.6','S6.7','S6.8','S6.9',
      'S7.1','S7.2','S7.3','S7.4',
      'S8.1 PF','S8.2 Act',
      'S9.1 Act','S9.2','S9.3','S9.4','S9.5','S9.6','S9.7','S9.8','S9.9','S9.10',
      'S10.1','S10.2','S10.3',
      'Spacer Manufacturer','Spacer Batch','Spacer Product','Spacer Size',
      'Spacer Wall','Spacer Qty','Spacer Grade','Spacer Finish','Spacer Receipt',
      'Desic. Manufacturer','Desic. Batch','Desic. Product',
      'Mol Sieve','Desic. Receipt','Desic. Expiry',
      'Sealant Manufacturer','Sealant Batch','Sealant Product',
      'Sealant Receipt','Sealant Expiry',
      'Desic Fill Time','Desic Assem Time',
    ];
    sheet.appendRow(headers);
    const hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground('#0b2e40').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(), data.date||'', data.shiftStart||'', data.shiftFinish||'',
    data.totalUnits||'', data.productionLine||'', data.completedBy||'',
    data.overallResult||'', data.passCount||0, data.failCount||0,
    data.s31||'', data.s32||'', data.s33||'', data.s34||'', data.s35||'',
    data.s41||'', data.s42||'', data.s43||'', data.s44||'', data.s45||'', data.s46||'', data.s47||'',
    data.s51||'', data.s52||'', data.s53||'', data.s54||'', data.s55||'', data.s56||'',
    data.s61||'', data.s62||'', data.s63||'', data.s64||'', data.s65||'', data.s66||'', data.s67||'', data.s68||'', data.s69||'',
    data.s71||'', data.s72||'', data.s73||'', data.s74||'',
    data.s81||'', data.s82||'',
    data.s91||'', data.s92||'', data.s93||'', data.s94||'', data.s95||'', data.s96||'', data.s97||'', data.s98||'', data.s99||'', data.s910||'',
    data.s101||'', data.s102||'', data.s103||'',
    data.s41_manufacturer||'', data.s41_batch||'', data.s41_product||'', data.s41_size||'',
    data.s41_wall_thickness||'', data.s41_quantity||'', data.s41_grade||'', data.s41_finish||'', data.s41_date_receipt||'',
    data.s51_manufacturer||'', data.s51_batch||'', data.s51_product||'',
    data.s51_mol_sieve||'', data.s51_date_receipt||'', data.s51_expiry||'',
    data.s61_manufacturer||'', data.s61_batch||'', data.s61_product||'',
    data.s61_date_receipt||'', data.s61_expiry||'',
    data.desiccantFillTime||'', data.desiccantAssemTime||'',
  ]);
}

// ─── REPORT GENERATION (DriveApp + Utilities only — no external API calls) ────

function generateAndSavePDF(data) {
  const tz         = Session.getScriptTimeZone();
  const dateStr    = data.date || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const safeDate   = dateStr.replace(/\//g, '-');
  const safeAuthor = (data.completedBy || 'report').substring(0, 20).replace(/[^a-zA-Z0-9]/g, '_');
  const fileName   = 'IGU_Report_' + safeDate + '_' + safeAuthor + '.html';

  // Build the HTML report string
  const html = buildHtmlReport(data, dateStr, tz);
  console.log('HTML built, length:', html.length);

  // Save directly to Drive as an HTML file — no DocumentApp, no external calls
  const blob   = Utilities.newBlob(html, MimeType.HTML, fileName);
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const file   = folder.createFile(blob);
  console.log('Report saved:', file.getId());

  return file;
}

// ─── HTML REPORT BUILDER ──────────────────────────────────────────────────────

function buildHtmlReport(data, dateStr, tz) {
  const generated = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  const totalChecked = (parseInt(data.passCount) || 0) + (parseInt(data.failCount) || 0);
  const hasFails     = (parseInt(data.failCount) || 0) > 0;
  const overallColor = hasFails ? '#c0392b' : '#155724';
  const overallBg    = hasFails ? '#fde8e8' : '#d4edda';

  // ── Inline CSS ────────────────────────────────────────────────────────────
  const css = `
    body  { font-family: Arial, sans-serif; font-size: 11pt; color: #222; margin: 0; padding: 24px; }
    h1    { color: #1B6B8A; font-size: 18pt; margin: 0 0 4px; }
    .sub  { color: #666; font-size: 9pt; font-style: italic; margin-bottom: 6px; }
    hr    { border: none; border-top: 2px solid #1B6B8A; margin: 10px 0 14px; }
    h2    { color: #1B6B8A; font-size: 11pt; margin: 18px 0 6px;
            border-left: 4px solid #4fc3a1; padding-left: 8px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
    td, th { padding: 5px 8px; border: 1px solid #dce8ed; font-size: 9pt; vertical-align: top; }
    td.lbl { background: #f0f5f7; font-weight: bold; color: #1B6B8A; width: 22%; white-space: nowrap; }
    td.lbl2{ background: #f0f5f7; font-weight: bold; color: #4a6370; width: 18%; white-space: nowrap; }
    .item { font-size: 10pt; margin: 2px 0; padding: 3px 6px; }
    .pass { color: #155724; font-weight: bold; }
    .fail { color: #c0392b; font-weight: bold; }
    .na   { color: #888; }
    .none { color: #aaa; }
    .sum  { padding: 10px 14px; border: 1px solid #dce8ed; margin-bottom: 14px; font-size: 11pt; }
    .photo{ max-width: 200px; max-height: 150px; margin: 6px 0; border: 1px solid #dce8ed; }
    .cap  { font-size: 8pt; color: #888; font-style: italic; }
    .foot { font-size: 8pt; color: #888; border-top: 1px solid #dce8ed; margin-top: 24px; padding-top: 6px; }
  `;

  // ── Helper: pass/fail badge HTML ──────────────────────────────────────────
  function pf(result) {
    if (!result || result === 'Not checked') return '<span class="none">○ Not recorded</span>';
    if (result === 'Pass') return '<span class="pass">✓ Pass</span>';
    if (result === 'Fail') return '<span class="fail">✗ Fail</span>';
    return '<span class="na">' + esc(result) + '</span>';
  }
  function act(result) {
    if (!result || result === 'Not completed' || result === '-') return '<span class="none">○ Not recorded</span>';
    if (result === 'Yes') return '<span class="pass">✓ Yes</span>';
    if (result === 'No')  return '<span class="fail">✗ No</span>';
    return '<span class="na">' + esc(result) + '</span>';
  }
  function esc(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }
  function val(s) { return esc(s) || '<span class="none">—</span>'; }
  function item(num, title, pfResult, actResult) {
    const badge = pfResult !== undefined ? pf(pfResult) : act(actResult);
    return '<p class="item"><b>' + esc(num) + '</b>&nbsp;&nbsp;' + esc(title) + '&nbsp;&nbsp;&nbsp;' + badge + '</p>';
  }
  function detailRow(l1, v1, l2, v2) {
    const c2 = (l2 !== undefined)
      ? '<td class="lbl2">' + esc(l2) + '</td><td>' + val(v2) + '</td>'
      : '<td colspan="2"></td>';
    return '<tr><td class="lbl2">' + esc(l1) + '</td><td>' + val(v1) + '</td>' + c2 + '</tr>';
  }
  function photo(dataUrl, caption) {
    if (!dataUrl || dataUrl.length < 200) return '';
    return '<img class="photo" src="' + dataUrl + '"><br><span class="cap">↑ ' + esc(caption) + '</span><br>';
  }

  // ── HTML assembly ─────────────────────────────────────────────────────────
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' + css + '</style></head><body>' +

  // Header
  '<h1>NGS NT – IGU Daily Test Report</h1>' +
  '<p class="sub">Northern Glass Solutions Pty Ltd &nbsp;·&nbsp; AS 4666:2012 &nbsp;·&nbsp; CSi ID 7709 &nbsp;·&nbsp; Aluminium Spacer Bar</p>' +
  '<hr>' +

  // Report Details
  '<h2>Report Details</h2>' +
  '<table>' +
    '<tr><td class="lbl">Date</td><td>' + val(dateStr) + '</td>' +
        '<td class="lbl">Shift Start</td><td>' + val(data.shiftStart) + '</td></tr>' +
    '<tr><td class="lbl">Shift Finish</td><td>' + val(data.shiftFinish) + '</td>' +
        '<td class="lbl">Production Line</td><td>' + val(data.productionLine) + '</td></tr>' +
    '<tr><td class="lbl">Total Units</td><td>' + val(data.totalUnits) + '</td>' +
        '<td class="lbl">Completed By</td><td>' + val(data.completedBy) + '</td></tr>' +
  '</table>' +

  // Summary
  '<h2>Summary</h2>' +
  '<div class="sum" style="background:' + overallBg + ';color:' + overallColor + ';font-weight:bold;">' +
    'Checked: ' + totalChecked + '&nbsp;&nbsp;|&nbsp;&nbsp;' +
    'Pass: ' + (data.passCount || 0) + '&nbsp;&nbsp;|&nbsp;&nbsp;' +
    'Fail: ' + (data.failCount || 0) + '&nbsp;&nbsp;|&nbsp;&nbsp;' +
    esc(data.overallResult || '—') +
  '</div>' +

  // Section 1
  '<h2>1. IGU Sampling Requirements (§5.8.2 &amp; §6.0)</h2>' +
  item('1.1', 'Number of IGU units to sample from production lot (Table 6.1)', undefined, data.s1) +

  // Section 2
  '<h2>2. Atmospheric Conditions (§5.2)</h2>' +
  item('2.1', 'Record atmospheric conditions at start of shift and every 4 hours', undefined, data.s2) +

  // Section 3
  '<h2>3. Glazing Materials (§5.3.1–5.3.7)</h2>' +
  item('3.1', 'Glass: Compliance with Order, Traceability & Types of Substrates', data.s31) +
  item('3.2', 'Glass: Edge Characteristics', data.s32) +
  item('3.3', 'Glass: Dimensional Properties (Thickness, Flatness, Height/Width)', data.s33) +
  item('3.4', 'Glass: Cleanliness – Dryness, Stains/Films, Fingerprints', data.s34) +
  item('3.5', 'Glass: Scratches, Blemishes, Marks & Inclusions', data.s35) +

  // Section 4
  '<h2>4. Spacer Bar &amp; Connectors (§5.4.1–5.4.8)</h2>' +
  item('4.1', 'Spacer Bar: Dimensional Properties & Stock Record', data.s41) +
  (
    (data.s41_manufacturer || data.s41_batch || data.s41_product || data.s41_size) ?
    '<table>' +
      detailRow('Manufacturer', data.s41_manufacturer, 'Batch No.', data.s41_batch) +
      detailRow('Product / Code', data.s41_product, 'Size (L×W×H)', data.s41_size) +
      detailRow('Wall Thickness', data.s41_wall_thickness, 'Quantity', data.s41_quantity) +
      detailRow('Grade', data.s41_grade, 'Finish / Colour', data.s41_finish) +
      detailRow('Date of Receipt', data.s41_date_receipt) +
    '</table>' : ''
  ) +
  photo(data.photo41, 'Spacer Bar Label Photo') +
  item('4.2', 'Spacer Bar: Customer Stock Record Traceability', data.s42) +
  item('4.3', 'Rigid Spacer Bar: Visual Inspection', data.s43) +
  item('4.4', 'Corner Keys & Straight Line Connectors', data.s44) +
  item('4.5', 'Spacer Bar Cleanliness', data.s45) +
  item('4.6', 'Dimensional Properties: Compliance with Manufacturer Specifications', data.s46) +
  item('4.7', 'Post Assembly Inspection: No splitting, tearing or flare-out on bent corners', data.s47) +

  // Section 5
  '<h2>5. Desiccant (§5.5.2–5.5.9)</h2>' +
  item('5.1', 'Desiccant: Origin, Compliance with Order & Shelf Life', data.s51) +
  (
    (data.s51_manufacturer || data.s51_batch || data.s51_product) ?
    '<table>' +
      detailRow('Manufacturer', data.s51_manufacturer, 'Batch No.', data.s51_batch) +
      detailRow('Product / Code', data.s51_product, 'Molecular Sieve', data.s51_mol_sieve) +
      detailRow('Date of Receipt', data.s51_date_receipt, 'Expiry Date', data.s51_expiry) +
    '</table>' : ''
  ) +
  photo(data.photo51, 'Desiccant Label Photo') +
  item('5.2', 'Desiccant Traceability: Traceable back to supplier for each customer order', data.s52) +
  item('5.3', 'Desiccant Effectiveness / Fit for Purpose', data.s53) +
  item('5.4', 'Desiccant Volume: Adequate volume in spacer bar cavity', data.s54) +
  item('5.5', 'Desiccant Exposure  [Fill: ' + esc(data.desiccantFillTime || '—') +
              '  →  Assembly: ' + esc(data.desiccantAssemTime || '—') + ']', data.s55) +
  item('5.6', 'Desiccant Suitability: Compatible with gas type (N/A if air-filled)', data.s56) +

  // Section 6
  '<h2>6. Sealants (§5.6.1–5.6.10)</h2>' +
  item('6.1', 'Sealant: Origin, Compliance with Order & Shelf Life', data.s61) +
  (
    (data.s61_manufacturer || data.s61_batch || data.s61_product) ?
    '<table>' +
      detailRow('Manufacturer', data.s61_manufacturer, 'Batch No.', data.s61_batch) +
      detailRow('Product / Code', data.s61_product, 'Date of Receipt', data.s61_date_receipt) +
      detailRow('Expiry Date', data.s61_expiry) +
    '</table>' : ''
  ) +
  photo(data.photo61, 'Sealant Label Photo') +
  item('6.2', 'Sealant Traceability', data.s62) +
  item('6.3', 'Sealant Type: Primary/Secondary', data.s63) +
  item('6.4', 'Primary Sealant Application', data.s64) +
  item('6.5', 'Secondary Sealant Application', data.s65) +
  item('6.6', 'Secondary Sealant Depth', data.s66) +
  item('6.7', 'Sealant Temperature', data.s67) +
  item('6.8', 'Sealant Cure Rate: Single or secondary sealants tested', data.s68) +
  item('6.9', 'Sealant Storage: Stored per manufacturer specification', data.s69) +

  // Section 7
  '<h2>7. Gas / IGU Gassing (§5.7.1–5.7.6)</h2>' +
  item('7.1', 'Gas: Origin, Compliance with Order & Gas Type', data.s71) +
  item('7.2', 'Gas Traceability: Traceable to supplier for each customer order', data.s72) +
  item('7.3', 'Desiccant & Gas Compatibility', data.s73) +
  item('7.4', 'Gas Fill Volume: Minimum 90% of IGU cavity volume', data.s74) +

  // Section 8
  '<h2>8. Compliance Markings (§4)</h2>' +
  item('8.1', 'Compliance Marking: Company name/logo, CSi ID 7709, AS 4666, Date of manufacture', data.s81) +
  item('8.2', 'Work Order / Customer Order Label Affixed to Finished Product', undefined, data.s82) +

  // Section 9
  '<h2>9. Final Product Check (§5.8.2–5.8.12)</h2>' +
  item('9.1',  'Test Format: Random samples (Table 6.1) before shipment', undefined, data.s91) +
  item('9.2',  'Sealant Cured: Visual & mechanical hardness check', data.s92) +
  item('9.3',  'Dimensional Properties: Height × Width, Squareness, Flatness', data.s93) +
  item('9.4',  'Uniformity of Spacer Depth: Minimum sealant depth per Table 5.7', data.s94) +
  item('9.5',  'Voids in Sealants: Visual inspection', data.s95) +
  item('9.6',  'Glass Alignment / Offset: Visual inspection', data.s96) +
  item('9.7',  'Cleanliness: Protrusion limits per §5.8.9', data.s97) +
  item('9.8',  'Marks, Scratches & Inclusions: Per Table 5.5', data.s98) +
  item('9.9',  'Moisture Content (Dew Point Test): Conducted on previous day production lot', data.s99) +
  item('9.10', 'Storage & Handling: Support, angle, spacing/interleaving', data.s910) +

  // Section 10
  '<h2>10. Equipment Availability (Appendix A)</h2>' +
  item('10.1', 'Mandatory Equipment for Dimensional Properties', data.s101) +
  item('10.2', 'Additional Equipment for Compliance Checks', data.s102) +
  item('10.3', 'Equipment Calibration: Within tolerance of 0.01mm and functional', data.s103) +

  // Footer
  '<p class="foot">TRF-IGU-Test-Report-Alum-Spacer-Bar &nbsp;·&nbsp; Northern Glass Solutions Pty Ltd' +
  ' &nbsp;·&nbsp; CSi ID 7709 &nbsp;·&nbsp; AS 4666:2012 &nbsp;·&nbsp; Generated: ' + generated + '</p>' +

  '</body></html>';
}

// ─── DEBUG / TEST ─────────────────────────────────────────────────────────────

/**
 * Run from the Apps Script editor (▶ Run) to test the full doPost flow
 * without a browser form submission. Check View → Executions for logs.
 */
function testDoPost() {
  const mockData = {
    date: '2026-05-13', shiftStart: '07:00', shiftFinish: '15:30',
    totalUnits: '45', productionLine: 'Line 1', completedBy: 'Paul Flores',
    overallResult: 'ALL PASS', passCount: 20, failCount: 0,
    s31:'Pass', s32:'Pass', s33:'Pass', s34:'Pass', s35:'Pass',
    s41:'Pass', s42:'Pass', s43:'Pass', s44:'Pass', s45:'Pass', s46:'Pass', s47:'Pass',
    s51:'Pass', s52:'Pass', s53:'Pass', s54:'Pass', s55:'Pass', s56:'N/A',
    s61:'Pass', s62:'Pass', s63:'Pass', s64:'Pass', s65:'Pass',
    s66:'Pass', s67:'Pass', s68:'Pass', s69:'Pass',
    s71:'Pass', s72:'Pass', s73:'Pass', s74:'Pass',
    s81:'Pass', s82:'Yes',
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
    SpreadsheetApp.getUi().alert('✅ PDF created!\n\n' + parsed.driveUrl);
  } else {
    SpreadsheetApp.getUi().alert('❌ Error:\n\n' + parsed.message);
  }
}
