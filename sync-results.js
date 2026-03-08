/**
 * Sync results_manual.json → output_manual.xlsx
 * ──────────────────────────────────────────────
 * Run this AFTER fetch-worker.js, with output_manual.xlsx closed in Excel:
 *   node sync-results.js
 *
 * Reads all accumulated results from the JSON file and writes the full
 * original sheet + is_correct column to output_manual.xlsx.
 */

const fs   = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const EXCEL_FILE  = "./volume/New stocks to be added to the universe.xlsx";
const SHEET_NAME  = "filtered by exchange";
const JSON_FILE   = "./volume/results/results_manual.json";
const OUTPUT_FILE = "./volume/output_manual.xlsx";

function log(level, message) {
  const time   = new Date().toISOString().slice(11, 19);
  const prefix = { info: "ℹ", ok: "✔", warn: "⚠", error: "✘" }[level] ?? "·";
  console.log(`[${time}] ${prefix}  ${message}`);
}

function main() {
  console.log("━".repeat(60));
  log("info", "Syncing results JSON → Excel");
  console.log("━".repeat(60));

  // 1. Load results JSON
  if (!fs.existsSync(JSON_FILE)) {
    log("error", `Results file not found: ${JSON_FILE}`);
    log("info",  "Run fetch-worker.js first to generate results.");
    process.exit(1);
  }
  const results = JSON.parse(fs.readFileSync(JSON_FILE, "utf-8"));
  const resultsMap = new Map(results.map((r) => [r.rowNumber, r]));
  log("ok", `Loaded ${results.length} entries from ${JSON_FILE}`);

  // 2. Load source Excel
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

  // 3. Apply is_correct to every row
  let synced = 0, pending = 0;
  const outputRows = allRows.map((row, idx) => {
    const result = resultsMap.get(idx);
    if (result !== undefined) {
      synced++;
      return { ...row, is_correct: result.is_correct };
    }
    pending++;
    return { ...row, is_correct: null }; // not yet fetched
  });

  // 4. Check output file isn't locked
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

  // 5. Write output Excel
  const outWs = XLSX.utils.json_to_sheet(outputRows);
  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, outWs, "results");
  XLSX.writeFile(outWb, OUTPUT_FILE);
  log("ok", `Output Excel written → ${OUTPUT_FILE}`);

  // 6. Summary
  console.log("");
  console.log("━".repeat(60));
  log("info", `Total rows   : ${allRows.length}`);
  log("ok",   `Synced rows  : ${synced}  (have is_correct value)`);
  log("warn",  `Pending rows : ${pending}  (not yet fetched — is_correct = null)`);

  const trueCount  = results.filter((r) => r.is_correct === true).length;
  const falseCount = results.filter((r) => r.is_correct === false).length;
  log("ok",   `  ✔ Correct    : ${trueCount}`);
  log("warn",  `  ✘ Not found  : ${falseCount}`);
  console.log("━".repeat(60));
}

main();
