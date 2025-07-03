/**
 * Google Apps Script – Alt-Fuel Activity Scoring  (multi-form version)
 * --------------------------------------------------------------------
 * Builds / refreshes a composite scoring sheet for every Google Form
 * listed in FORMS[], and writes a run-summary to the LOGS sheet.
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

const CATALOG_SHEET = 'Activities';   // catalogue of names + descriptions
const LOG_SHEET     = 'LOGS';         // NEW: central run-summary sheet

// Weights must sum to 1.00
const WEIGHTS = {
  'Impact on Goal'       : 0.30,
  'Time-to-impact'       : 0.25,
  'Perceived Cost'       : 0.25,
  'Risk and Uncertainty' : 0.20,
};

// Choose 'mean' or 'median'
const AGG_METHOD = 'median';
/* ------------------------------------------------------------------ */


/**
 * MAIN entry – attach a single “On form submit” trigger to this.
 */
function scoreAllForms() {
  const ss         = SpreadsheetApp.getActive();
  const DESC_MAP   = buildDescriptionMap_(ss);
  const timestamp  = new Date();               // one stamp for the whole run
  const logEntries = [];                       // gather per-form stats

  FORMS.forEach(cfg => {
    if (!ss.getSheetByName(cfg.raw)) {
      Logger.log(`Sheet “${cfg.raw}” not found – skipped.`);
      return;
    }
    const stats = buildComposite_(ss, cfg.raw, cfg.out, DESC_MAP);
    if (stats) logEntries.push(stats);
  });

  writeLogs_(ss, logEntries, timestamp);
}


/* =============  HELPER FUNCTIONS  ================================= */

/**
 * Reads the Activities catalogue and returns { activityName : description }
 */
function buildDescriptionMap_(ss) {
  const cat = ss.getSheetByName(CATALOG_SHEET);
  if (!cat) throw `Sheet “${CATALOG_SHEET}” not found`;

  const data = cat.getDataRange().getValues();
  const hdr  = data.shift();

  const nameCol = hdr.findIndex(h => /^activity\s*name$/i.test(h));
  const descCol = hdr.findIndex(h => /^description$/i.test(h));
  if (nameCol === -1 || descCol === -1)
    throw `Columns “Activity Name” and/or “Description” not found in “${CATALOG_SHEET}”`;

  const map = {};
  data.forEach(r => {
    const name = r[nameCol];
    if (name) map[name] = r[descCol] || '';
  });
  return map;
}


/**
 * Builds / refreshes one composite sheet and returns run-stats.
 *
 * @returns {Object}  { rawSheet, compSheet, rawCount, processedCount }
 */
function buildComposite_(ss, rawName, outName, DESC_MAP) {

  /* --- 0. PREP OUTPUT SHEET & SALVAGE NOTES -------------------- */
  let out = ss.getSheetByName(outName);
  const NOTES_BY_ACTIVITY = {};

  if (out) {                                       // sheet exists ⇒ read notes
    const old     = out.getDataRange().getValues();
    const hdrRow  = old[2] || [];                  // header is row-index 2
    const noteCol = hdrRow.indexOf('Notes');
    if (noteCol !== -1) {
      for (let r = 3; r < old.length; r++) {       // data start row-index 3
        const act  = old[r][0];
        const note = old[r][noteCol];
        if (act) NOTES_BY_ACTIVITY[act] = note;
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

  // {activity : {criterion : [scores]}}
  const scores = {};
  hdr.forEach((h, colIdx) => {
    const m = h.match(/^(Impact on Goal|Time-to-impact|Perceived Cost|Risk and Uncertainty) \[(.*)]$/);
    if (!m) return;
    const crit = m[1], activity = m[2];
    data.forEach(r => {
      const v = Number(r[colIdx]);
      if (!v) return;
      (scores[activity]       = scores[activity] || {});
      (scores[activity][crit] = scores[activity][crit] || []).push(v);
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
  const header = [
    'Activity', 'Description',
    'Impact ' + COL_LABEL,
    'Time '   + COL_LABEL,
    'Cost '   + COL_LABEL,
    'Risk '   + COL_LABEL,
    'Composite',
    'Notes'
  ];

  const body = Object.entries(scores)
    .map(([act, critObj]) => {
      const imp  = agg(critObj['Impact on Goal']);
      const time = agg(critObj['Time-to-impact']);
      const cost = agg(critObj['Perceived Cost']);
      const risk = agg(critObj['Risk and Uncertainty']);
      const comp = (
        imp  * WEIGHTS['Impact on Goal']       +
        time * WEIGHTS['Time-to-impact']       +
        cost * WEIGHTS['Perceived Cost']       +
        risk * WEIGHTS['Risk and Uncertainty']
      );
      const desc  = DESC_MAP[act] || '';
      const note  = NOTES_BY_ACTIVITY[act] || '';
      const rnd   = v => isNaN(v) ? v : Math.round(v*100)/100;

      return [act, desc, rnd(imp), rnd(time), rnd(cost), rnd(risk), rnd(comp), note];
    })
    .sort((a,b) => b[6] - a[6]);                 // sort by Composite score

  /* --- 4. WRITE TO SHEET -------------------------------------- */
  // (Rows 1-2 intentionally left blank to preserve header offset)
  out.getRange(3, 1, body.length + 1, header.length)
     .setValues([header, ...body]);

  // Ensure Notes column is plain text so users can edit freely
  if (body.length > 0) {
    out.getRange(4, header.length, body.length, 1)
       .setNumberFormat('@STRING@');
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
 *
 * @param {Array<Object>} entries – one object per form
 * @param {Date}          stamp   – timestamp for this run
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

     // --- pretty-print the Timestamp column (col A, rows 2-last) ---
  if (rows.length) {
    log.getRange(2, 1, rows.length, 1)          // skip header row
       .setNumberFormat('yyyy-mm-dd HH:mm:ss'); // or 'm/d/yy h:mm:ss AM/PM'
  }
}
