/**
 * Google Apps Script – Alt-Fuel Activity Scoring (multi-form version)
 * --------------------------------------------------------------------
 * Builds / refreshes a composite scoring sheet for every Google Form
 * listed in FORMS[], and writes a run-summary to the LOGS sheet.
 * 
 * Changes:
 * - Persists user-editable columns per Activity:
 *   "New Description", "Project Partners", "Deliverables",
 *   "Other Relevant Activities", "Notes"
 * - Inserts two auto-filled columns (between Composite and Notes):
 *   "Original Description" (Activities.Description),
 *   "Workstream" (Activities.Workstream)
 * - Joins Activities by 'form_title'
 */

/* --------- GLOBAL CONFIGURATION ---------------------------------- */
const FORMS = [
  { raw: 'Bunker'     , out: 'Composite_Bunker'     },
  { raw: 'Safety'     , out: 'Composite_Safety'     },
  { raw: 'Digital'    , out: 'Composite_Digital'    },
  { raw: 'Specs'      , out: 'Composite_Specs'      },
  { raw: 'Production' , out: 'Composite_Production' },
  { raw: 'Vessels'    , out: 'Composite_Vessels'    },
];

const CATALOG_SHEET = 'Activities';   // catalogue with form_title, Workstream, Description, etc.
const LOG_SHEET     = 'LOGS';         // central run-summary sheet

// Weights must sum to 1.00
const WEIGHTS = {
  'Impact on Goal'       : 0.30,
  'Time-to-impact'       : 0.25,
  'Perceived Cost'       : 0.25,
  'Risk and Uncertainty' : 0.20,
};

// Choose 'mean' or 'median'
const AGG_METHOD = 'median';

// User-editable, persistent columns (keyed by Activity)
const PERSIST_COLS = [
  'New Description',
  'Project Partners',
  'Deliverables',
  'Other Relevant Activities',
  'Notes'
];

// Auto-filled cols to insert between Composite and Notes
const AUTO_BETWEEN = [
  'Original Description', // Activities.Description
  'Workstream'           // Activities.Workstream
];
/* ------------------------------------------------------------------ */


/**
 * MAIN entry – attach a single “On form submit” trigger to this.
 */
function scoreAllForms() {
  const ss         = SpreadsheetApp.getActive();
  const CAT_MAP    = buildCatalogMap_(ss); // { form_title : { description, workstream } }
  const timestamp  = new Date();
  const logEntries = [];

  FORMS.forEach(cfg => {
    if (!ss.getSheetByName(cfg.raw)) {
      Logger.log(`Sheet “${cfg.raw}” not found – skipped.`);
      return;
    }
    const stats = buildComposite_(ss, cfg.raw, cfg.out, CAT_MAP);
    if (stats) logEntries.push(stats);
  });

  writeLogs_(ss, logEntries, timestamp);
}


/* =============  HELPER FUNCTIONS  ================================= */

/**
 * Reads the Activities catalogue and returns:
 * { form_title : { description: <string>, workstream: <string> } }
 *
 * Columns required in Activities:
 * - "form_title" (join key)
 * - "Description"
 * - "Workstream"
 */
function buildCatalogMap_(ss) {
  const cat = ss.getSheetByName(CATALOG_SHEET);
  if (!cat) throw `Sheet “${CATALOG_SHEET}” not found`;

  const data = cat.getDataRange().getValues();
  const hdr  = data.shift().map(h => String(h).trim());

  const ftCol   = hdr.findIndex(h => /^form[_\s-]*title$/i.test(h));
  const descCol = hdr.findIndex(h => /^description$/i.test(h));
  const wsCol   = hdr.findIndex(h => /^workstream$/i.test(h));

  if (ftCol === -1)   throw `Column “form_title” not found in “${CATALOG_SHEET}”`;
  if (descCol === -1) throw `Column “Description” not found in “${CATALOG_SHEET}”`;
  if (wsCol === -1)   throw `Column “Workstream” not found in “${CATALOG_SHEET}”`;

  const map = {};
  data.forEach(r => {
    const key = String(r[ftCol] || '').trim();
    if (!key) return;
    map[key] = {
      description: r[descCol] || '',
      workstream : r[wsCol]   || ''
    };
  });
  return map;
}


/**
 * Builds / refreshes one composite sheet and returns run-stats.
 *
 * @returns {Object}  { rawSheet, compSheet, rawCount, processedCount }
 */
function buildComposite_(ss, rawName, outName, CAT_MAP) {

  /* --- 0. PREP OUTPUT SHEET & SALVAGE USER-EDITABLE FIELDS ----- */
  let out = ss.getSheetByName(outName);

  // Read any existing persisted values by Activity
  const salvage = {}; // { Activity : { 'New Description':..., 'Project Partners':..., ... } }

  if (out) {
    const old = out.getDataRange().getValues();
    const hdr = old[2] || []; // header row at row 3 (0-based index 2)
    // Find indices for Activity + all persisted columns
    const actIdx = hdr.indexOf('Activity');

    const persistIdx = {};
    PERSIST_COLS.forEach(c => persistIdx[c] = hdr.indexOf(c));

    if (actIdx !== -1) {
      for (let r = 3; r < old.length; r++) { // data start at row 4 (0-based 3)
        const row = old[r];
        const act = row[actIdx];
        if (!act) continue;
        salvage[act] = salvage[act] || {};
        PERSIST_COLS.forEach(c => {
          const idx = persistIdx[c];
          if (idx !== -1) salvage[act][c] = row[idx];
        });
      }
    }
    out.clear();
  } else {
    out = ss.insertSheet(outName);
  }

  /* --- 1. READ RESPONSES -------------------------------------- */
  const raw  = ss.getSheetByName(rawName);
  const data = raw.getDataRange().getValues();
  const hdr  = data.shift();
  const RESP_COUNT = data.length;

  // Parse headers like: "<Criterion> [<form_title>]"
  // Build: { activity(form_title) : { criterion : [scores] } }
  const scores = {};
  hdr.forEach((h, colIdx) => {
    const m = String(h).match(/^(Impact on Goal|Time-to-impact|Perceived Cost|Risk and Uncertainty)\s*\[(.*)]$/);
    if (!m) return;
    const crit = m[1];
    const activityKey = String(m[2] || '').trim(); // this should equal Activities.form_title

    data.forEach(r => {
      const v = Number(r[colIdx]);
      if (!v && v !== 0) return;       // skip blanks/non-numbers
      (scores[activityKey]       = scores[activityKey] || {});
      (scores[activityKey][crit] = scores[activityKey][crit] || []).push(v);
    });
  });

  /* --- 2. AGGREGATORS ---------------------------------------- */
  function agg(arr){
    if (!arr || !arr.length) return NaN;
    if (AGG_METHOD === 'mean'){
      return arr.reduce((a,b)=>a+b,0)/arr.length;
    } else {                          // median
      const s = arr.slice().sort((a,b)=>a-b);
      const mid = Math.floor(s.length/2);
      return s.length % 2 ? s[mid] : (s[mid-1]+s[mid])/2;
    }
  }
  const COL_LABEL = AGG_METHOD === 'median' ? 'Median' : 'Avg';

  /* --- 3. BUILD OUTPUT ARRAY ---------------------------------- */
  // Final column order:
  // Activity, Impact <label>, Time <label>, Cost <label>, Risk <label>, Composite,
  // Original Description, Workstream, New Description, Project Partners,
  // Deliverables, Other Relevant Activities, Notes

  const header = [
    'Activity',
    'Impact ' + COL_LABEL,
    'Time '   + COL_LABEL,
    'Cost '   + COL_LABEL,
    'Risk '   + COL_LABEL,
    'Composite',
    ...AUTO_BETWEEN,   // Original Description, Workstream (auto-filled)
    ...PERSIST_COLS    // user-editable persistent fields
  ];

  const body = Object.entries(scores)
    .map(([act, critObj]) => {
      const imp  = agg(critObj['Impact on Goal']);
      const time = agg(critObj['Time-to-impact']);
      const cost = agg(critObj['Perceived Cost']);
      const risk = agg(critObj['Risk and Uncertainty']);
      const comp = (
        (isNaN(imp)  ? 0 : imp)  * WEIGHTS['Impact on Goal']       +
        (isNaN(time) ? 0 : time) * WEIGHTS['Time-to-impact']       +
        (isNaN(cost) ? 0 : cost) * WEIGHTS['Perceived Cost']       +
        (isNaN(risk) ? 0 : risk) * WEIGHTS['Risk and Uncertainty']
      );

      const cat = CAT_MAP[act] || { description: '', workstream: '' };
      const origDesc = cat.description || '';
      const workstream = cat.workstream || '';

      const keep = salvage[act] || {};
      const rnd = v => (typeof v === 'number' && !isNaN(v)) ? Math.round(v*100)/100 : v;

      return [
        act,
        rnd(imp), rnd(time), rnd(cost), rnd(risk),
        rnd(comp),
        origDesc,
        workstream,
        keep['New Description'] || '',
        keep['Project Partners'] || '',
        keep['Deliverables'] || '',
        keep['Other Relevant Activities'] || '',
        keep['Notes'] || ''
      ];
    })
    .sort((a,b) => {
      // Composite column index = 6th (0-based 5)
      const ia = a[5], ib = b[5];
      if (isNaN(ia) && isNaN(ib)) return 0;
      if (isNaN(ia)) return 1;
      if (isNaN(ib)) return -1;
      return ib - ia; // descending
    });

  /* --- 4. WRITE TO SHEET -------------------------------------- */
  // Leave rows 1-2 blank (visual buffer). Header row at row 3.
  out.getRange(3, 1, body.length + 1, header.length)
     .setValues([header, ...body]);

  // Make all persistent columns plain text for easy editing
  if (body.length > 0) {
    const startCol = header.indexOf(PERSIST_COLS[0]) + 1; // 1-based
    const persistCount = PERSIST_COLS.length;
    if (startCol > 0) {
      out.getRange(4, startCol, body.length, persistCount).setNumberFormat('@STRING@');
    }
  }

  // Auto columns can also be treated as plain text (optional)
  const autoStart = header.indexOf(AUTO_BETWEEN[0]) + 1;
  if (body.length > 0 && autoStart > 0) {
    out.getRange(4, autoStart, body.length, AUTO_BETWEEN.length).setNumberFormat('@STRING@');
  }

  /* --- 5. RETURN STATS ---------------------------------------- */
  return {
    rawSheet       : rawName,
    compSheet      : outName,
    rawCount       : RESP_COUNT,
    processedCount : body.length
  };
}


/**
 * Writes / rewrites the central LOGS sheet with run-summary rows.
 */
function writeLogs_(ss, entries, stamp) {
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) log = ss.insertSheet(LOG_SHEET);
  log.clear();

  const header = ['Timestamp', 'Form Sheet', 'Composite Sheet',
                  'Raw Submissions', 'Scores Processed'];
  const rows = entries.map(e => [
    stamp,
    e.rawSheet,
    e.compSheet,
    e.rawCount,
    e.processedCount
  ]);

  log.getRange(1, 1, rows.length + 1, header.length)
     .setValues([header, ...rows]);

  if (rows.length) {
    log.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
  }
}
