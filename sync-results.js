/**
 * Sync results → output_manual.xlsx
 * ──────────────────────────────────────────────
 * Run this AFTER fetch-worker.js and/or fetch-finnhub.js,
 * with output_manual.xlsx closed in Excel:
 *   node sync-results.js
 *
 * Merges:
 *   results_manual.json  → is_correct column   (from fetch-worker.js)
 *   finnhub_manual.json  → finhub_symbol column (from fetch-finnhub.js)
 * Both are optional — whichever exists will be applied.
 */

const fs   = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const EXCEL_FILE    = "./volume/New stocks to be added to the universe.xlsx";
const SHEET_NAME    = "filtered by exchange";
const JSON_FILE     = "./volume/results/results_manual.json";
const FINNHUB_FILE  = "./volume/results/finnhub_manual.json";
// const FINNHUB_FILE = null;
const OUTPUT_FILE   = "./volume/output_manual.xlsx";

function log(level, message) {
  const time   = new Date().toISOString().slice(11, 19);
  const prefix = { info: "ℹ", ok: "✔", warn: "⚠", error: "✘" }[level] ?? "·";
  console.log(`[${time}] ${prefix}  ${message}`);
}

function main() {
  console.log("━".repeat(60));
  log("info", "Syncing results JSON → Excel");
  console.log("━".repeat(60));

  // 1. Load GTN results (is_correct) — optional
  let resultsMap = new Map();
  if (fs.existsSync(JSON_FILE)) {
    const results = JSON.parse(fs.readFileSync(JSON_FILE, "utf-8"));
    resultsMap = new Map(results.map((r) => [r.rowNumber, r]));
    log("ok", `Loaded ${results.length} GTN entries  (is_correct)  ← ${JSON_FILE}`);
  } else {
    log("warn", `GTN results not found — is_correct will be null  (${JSON_FILE})`);
  }

  // 2. Load Finnhub results (finhub_symbol) — optional
  let finnhubMap = new Map();
  if (fs.existsSync(FINNHUB_FILE)) {
    const finnhub = JSON.parse(fs.readFileSync(FINNHUB_FILE, "utf-8"));
    finnhubMap = new Map(finnhub.map((r) => [r.rowNumber, r]));
    log("ok", `Loaded ${finnhub.length} Finnhub entries (finhub_symbol) ← ${FINNHUB_FILE}`);
  } else {
    log("warn", `Finnhub results not found — finhub_symbol will be null  (${FINNHUB_FILE})`);
  }

  if (resultsMap.size === 0 && finnhubMap.size === 0) {
    log("error", "No result files found at all. Run fetch-worker.js and/or fetch-finnhub.js first.");
    process.exit(1);
  }

  // 3. Load source Excel
  if (!fs.existsSync(EXCEL_FILE)) {
    log("error", `Excel file not found: ${EXCEL_FILE}`);
    process.exit(1);
  }
  const wb = XLSX.readFile(EXCEL_FILE);
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) {
    log("error", `Sheet "${SHEET_NAME}" not found. Available: ${wb.SheetNames.join(", ")}`);
    process.exit(1);
  }
  const allRows = XLSX.utils.sheet_to_json(ws);
  log("ok", `Loaded ${allRows.length} rows from source Excel`);

  // 4. Apply is_correct + finhub_symbol to every row
  let syncedGtn = 0, syncedFinnhub = 0, pending = 0;
  const outputRows = allRows.map((row, idx) => {
    const gtn     = resultsMap.get(idx);
    const finnhub = finnhubMap.get(idx);
    const out     = { ...row };

    if (gtn !== undefined) {
      out.is_correct = gtn.is_correct;
      syncedGtn++;
    } else {
      out.is_correct = null;
    }

    if (finnhub !== undefined) {
      out.finhub_symbol = finnhub.finhub_symbol ?? null;
      syncedFinnhub++;
    } else {
      out.finhub_symbol = null;
    }

    if (gtn === undefined && finnhub === undefined) pending++;
    return out;
  });

  // 5. Check output file isn't locked
  if (fs.existsSync(OUTPUT_FILE)) {
    try {
      const fd = fs.openSync(OUTPUT_FILE, "r+");
      fs.closeSync(fd);
    } catch {
      log("error", `Output file is locked/open in another program: ${OUTPUT_FILE}`);
      log("info",  "Close the file in Excel and run this script again.");
      process.exit(1);
    }
  }

  // 6. Write output Excel
  const outWs = XLSX.utils.json_to_sheet(outputRows);
  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, outWs, "results");
  XLSX.writeFile(outWb, OUTPUT_FILE);
  log("ok", `Output Excel written → ${OUTPUT_FILE}`);

  // 7. Summary
  console.log("");
  console.log("━".repeat(60));
  log("info", `Total rows          : ${allRows.length}`);
  log("ok",   `is_correct synced   : ${syncedGtn}`);
  log("ok",   `finhub_symbol synced: ${syncedFinnhub}`);
  log("warn",  `Fully pending       : ${pending}  (neither column set)`);
  console.log("━".repeat(60));
}

main();
