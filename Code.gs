// ================================================================
// STAGE Release Calendar — Google Apps Script
// Sheet ID: 1awfYySdnnxPql4Bj1C7KJkjFcBGExWuki0yAI3xxhZI
//
// Deploy as: Web App → Execute as Me → Anyone can access
// ================================================================

const SPREADSHEET_ID = '1awfYySdnnxPql4Bj1C7KJkjFcBGExWuki0yAI3xxhZI';

// ---- Column indexes (0-based) — adjust if sheet columns change ----
// Expected columns: Month | Date | Title | Genre | Dialect | Format | Type | Runtime | Episodes | Release Pattern | Purpose | Dubbing | Orig Dialect | Comment
const COL = {
  month:          0,   // A: "March 2026" / "April 2026"
  date:           1,   // B: "1 March 2026"
  title:          2,   // C: Content Title
  genre:          3,   // D: Genre
  dialectRaw:     4,   // E: Universal / Regional / HR,RJ,BH,GJ
  format:         5,   // F: Long Series / Feature Film / Micro Drama / etc.
  ctype:          6,   // G: STAGE Original / STAGE Dubbed
  runtime:        7,   // H: Runtime in minutes
  episodes:       8,   // I: Episode count (optional)
  releasePattern: 9,   // J: Release pattern text (optional)
  purpose:        10,  // K: Retention / Acquisition / Engagement / Re-engagement (optional)
  dubbing:        11,  // L: Done / Recommended / Not Recommended / TBD (optional)
  origDialect:    12,  // M: Original dialect (optional)
  comment:        13,  // N: Comment (optional)
};

// Sheet GIDs for each month tab
const SHEET_GIDS = { march: 1674310550, april: 1186980063, may: 866138588, june: 404386009, july: 741448647, august: 1133677531 };

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const month = (body.month || '').toLowerCase();
    const releases = body.releases;
    if (!month || !releases) throw new Error('Missing month or releases');
    saveReleases(month, releases);
    const output = ContentService.createTextOutput(JSON.stringify({ ok: true }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  } catch (err) {
    const output = ContentService.createTextOutput(JSON.stringify({ error: err.toString() }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}

function saveReleases(month, releases) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const targetGid = SHEET_GIDS[month];
  if (!targetGid) throw new Error('Unknown month: ' + month);

  let sheet = null;
  for (const s of ss.getSheets()) {
    if (s.getSheetId() === targetGid) { sheet = s; break; }
  }
  if (!sheet) throw new Error('Sheet tab not found for: ' + month);

  const allData = sheet.getDataRange().getValues();
  // Preserve first 2 header rows; rebuild data rows from releases
  const headerRows = allData.length >= 2 ? allData.slice(0, 2) : allData.slice(0, 1);

  const dataRows = releases.map(r => {
    const dialects = Array.isArray(r.dialects) ? r.dialects.join(', ') : (r.dialects || '');
    return [
      '',                  // A: format group (preserved separately)
      r.title || '',       // B: Title
      r.purpose || '',     // C: Purpose
      r.genre || '',       // D: Genre
      dialects,            // E: Dialects
      r.date || '',        // F: Date
      r.releasePattern || '', // G: Release Pattern
      '',                  // H: (blank)
      r.ctype === 'Original' ? 'STAGE Original' : 'STAGE Acquired', // I: Type
      r.origDialect || '', // J: Orig Dialect
      r.episodes || '',    // K: Episodes
      r.runtime || '',     // L: Runtime
      r.dubbing || '',     // M: Dubbing
      '',                  // N: (blank)
      r.comment || '',     // O: Comment
    ];
  });

  const newData = [...headerRows, ...dataRows];
  sheet.clearContents();
  if (newData.length > 0) {
    sheet.getRange(1, 1, newData.length, 15).setValues(newData);
  }
}

function doGet(e) {
  try {
    const result = readReleases();
    const output = ContentService.createTextOutput(JSON.stringify(result));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  } catch (err) {
    const output = ContentService.createTextOutput(JSON.stringify({ error: err.toString() }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}

function readReleases() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const marchReleases = [];
  const aprilReleases = [];
  const mayReleases = [];
  const juneReleases = [];

  // Try to find sheets by name, fallback to all sheets
  const sheetNames = ss.getSheets().map(s => s.getName());
  Logger.log('Available sheets: ' + sheetNames.join(', '));

  // Read from ALL sheets and classify by month column
  for (const sheet of ss.getSheets()) {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    let currentMonth = '';

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip blank rows
      const title = String(row[COL.title] || '').trim();
      if (!title) continue;

      // Track current month from column A
      const monthCell = String(row[COL.month] || '').trim();
      if (monthCell) currentMonth = monthCell.toLowerCase();

      // Build release object
      const release = {
        title:          title,
        genre:          String(row[COL.genre] || '').trim(),
        dialects:       mapDialects(row[COL.dialectRaw]),
        origDialect:    String(row[COL.origDialect] || '').trim(),
        date:           parseDate(row[COL.date]),
        releasePattern: String(row[COL.releasePattern] || '').trim(),
        format:         mapFormat(row[COL.format]),
        ctype:          mapType(row[COL.ctype]),
        episodes:       parseInt(row[COL.episodes]) || 0,
        runtime:        parseInt(row[COL.runtime]) || 0,
        purpose:        mapPurpose(row[COL.purpose], row[COL.format]),
        dubbing:        String(row[COL.dubbing] || 'Done').trim(),
        comment:        String(row[COL.comment] || '').trim(),
      };

      // Classify by month
      if (currentMonth.includes('march')) {
        marchReleases.push(release);
      } else if (currentMonth.includes('april')) {
        aprilReleases.push(release);
      } else if (currentMonth.includes('may')) {
        mayReleases.push(release);
      } else if (currentMonth.includes('june')) {
        juneReleases.push(release);
      }
    }
  }

  return { march: marchReleases, april: aprilReleases, may: mayReleases, june: juneReleases, updatedAt: new Date().toISOString() };
}

// ---- Helper: parse date like "1 March 2026" or "2026-03-01" ----
function parseDate(val) {
  if (!val) return 'TBD';
  const s = String(val).trim();
  if (!s) return 'TBD';

  // Already ISO format
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // If it's a Date object from Sheets
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // "1 March 2026" or "1 March" format
  const MONTHS = { january:'01', february:'02', march:'03', april:'04', may:'05', june:'06',
                   july:'07', august:'08', september:'09', october:'10', november:'11', december:'12' };
  const parts = s.split(' ');
  if (parts.length >= 2) {
    const day = String(parseInt(parts[0])).padStart(2, '0');
    const month = MONTHS[parts[1].toLowerCase()] || '01';
    const year = parts[2] || '2026';
    return `${year}-${month}-${day}`;
  }

  return 'TBD';
}

// ---- Helper: map dialect field to array ----
function mapDialects(val) {
  const s = String(val || '').trim().toLowerCase();

  if (!s || s === 'universal' || s === 'all dialects') return ['HR','RJ','BH','GJ'];
  if (s === 'other dialect') return ['RJ','BH','GJ'];
  if (s === 'regional' || s === 'regional dialect') return ['HR'];

  // Parse comma-separated codes: "HR, RJ, BH"
  const codes = s.split(',').map(x => x.trim().toUpperCase()).filter(x => ['HR','RJ','BH','GJ'].includes(x));
  return codes.length > 0 ? codes : ['HR','RJ','BH','GJ'];
}

// ---- Helper: map format string ----
function mapFormat(val) {
  const f = String(val || '').trim().toLowerCase();
  if (f.includes('long series')) return 'Long';
  if (f.includes('binge') || f.includes('web series')) return 'Binge';
  if (f.includes('feature film') || f.includes('feature')) return 'Feature';
  if (f.includes('mini film') || f.includes('mini') || f.includes('short film')) return 'Mini';
  if (f.includes('micro drama') || f.includes('micro')) return 'Micro';
  return 'Feature';
}

// ---- Helper: map content type ----
function mapType(val) {
  const t = String(val || '').trim().toLowerCase();
  if (t.includes('original')) return 'Original';
  if (t.includes('dubbed') || t.includes('acquired') || t.includes('dub')) return 'Acquired';
  return 'Original';
}

// ---- Helper: map purpose (infer if blank) ----
function mapPurpose(val, formatVal) {
  const p = String(val || '').trim();
  if (p) return p;
  // Infer from format
  const f = mapFormat(formatVal);
  if (f === 'Long') return 'Retention';
  if (f === 'Micro') return 'Re-engagement';
  return 'Acquisition';
}

// ---- Test function — run manually in GAS editor to check output ----
function test() {
  const result = readReleases();
  Logger.log('March releases: ' + result.march.length);
  Logger.log('April releases: ' + result.april.length);
  Logger.log(JSON.stringify(result.march[0], null, 2));
}
