/**
 * Finnhub symbol lookup worker (Node.js)
 * ────────────────────────────────────────
 * Reads the Excel sheet, looks up each row's ISIN against the Finnhub API,
 * and saves the returned finhub_symbol to a persistent JSON file.
 *
 * Edit START_ROW / END_ROW below, then run:
 *   node fetch-finnhub.js
 *
 * Results are appended/updated in:
 *   volume/results/finnhub_manual.json
 *
 * Run sync-results.js after to write output_manual.xlsx with finhub_symbol column.
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────────

const START_ROW = 0;   // inclusive, 0-based (first data row after header)
const END_ROW   = 50;  // exclusive  → processes rows 0..49

const EXCEL_FILE  = "./volume/New stocks to be added to the universe.xlsx";
const SHEET_NAME  = "filtered by exchange";
const RESULTS_DIR = "./volume/results";
const JSON_FILE   = "finnhub_manual.json";

const FINNHUB_TOKEN = "d1qgq9hr01qo4qd75oe0d1qgq9hr01qo4qd75oeg";
const FINNHUB_URL   = "https://finnhub.io/api/v1/search";

const DELAY_MS        = 300;   // delay between normal requests (ms)
const RATE_LIMIT_WAIT = 60000; // wait time on 429 (ms)

// ─── DEPENDENCIES ────────────────────────────────────────────────────────────

const fs   = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function log(level, message) {
  const time   = new Date().toISOString().slice(11, 19); // HH:MM:SS
  const prefix = { info: "ℹ", ok: "✔", warn: "⚠", error: "✘" }[level] ?? "·";
  console.log(`[${time}] ${prefix}  ${message}`);
}

/**
 * Fetch Finnhub symbol by ISIN.
 * Automatically retries once after a 60-second wait on 429.
 * Returns { displaySymbol, symbol, description, type } or null if not found.
 */
async function fetchFinnhubSymbol(isin, isRetry = false) {
  const url = new URL(FINNHUB_URL);
  url.searchParams.set("token", FINNHUB_TOKEN);
  url.searchParams.set("q", isin);

  try {
    const res = await fetch(url.toString(), {
      headers: { Accept: "application/json" },
    });

    if (res.status === 429) {
      if (isRetry) {
        log("error", `Rate limited again after wait — skipping ISIN ${isin}`);
        return null;
      }
      log("warn", `Rate limited by Finnhub. Waiting ${RATE_LIMIT_WAIT / 1000}s before retry...`);
      await sleep(RATE_LIMIT_WAIT);
      return fetchFinnhubSymbol(isin, true);
    }

    if (!res.ok) {
      log("warn", `HTTP ${res.status} for ISIN ${isin}`);
      return null;
    }

    const data = await res.json();

    if (data.count > 0 && data.result?.length > 0) {
      const r = data.result[0];
      return {
        displaySymbol: r.displaySymbol,
        symbol:        r.symbol,
        description:   r.description,
        type:          r.type,
      };
    }

    return null; // no results
  } catch (err) {
    log("error", `Network error for ISIN ${isin}: ${err.message}`);
    return null;
  }
}

// ─── MAIN ────────────────────────────────────────────────────────────────────

async function main() {
  console.log("━".repeat(60));
  log("info", "Finnhub Symbol Lookup — Node.js manual runner");
  log("info", `Rows: ${START_ROW} → ${END_ROW - 1}  (${END_ROW - START_ROW} total)`);
  log("info", `Excel: ${EXCEL_FILE}`);
  console.log("━".repeat(60));

  // 1. Load workbook
  log("info", "Loading Excel file...");
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
  log("ok", `Loaded ${allRows.length} rows from sheet "${SHEET_NAME}"`);

  // 2. Slice batch
  const batch = allRows.slice(START_ROW, END_ROW);
  if (batch.length === 0) {
    log("warn", "No rows in the specified range. Check START_ROW / END_ROW.");
    process.exit(0);
  }
  log("info", `Processing ${batch.length} rows (${START_ROW}–${START_ROW + batch.length - 1})`);
  console.log("");

  // 3. Load existing JSON (append/update, never overwrite previous runs)
  fs.mkdirSync(RESULTS_DIR, { recursive: true });
  const jsonPath = path.join(RESULTS_DIR, JSON_FILE);
  let existing = [];
  if (fs.existsSync(jsonPath)) {
    existing = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
    log("info", `Found existing file with ${existing.length} entries — will update in place`);
  } else {
    log("info", "No existing file found — creating new one");
  }
  const resultsMap = new Map(existing.map((r) => [r.rowNumber, r]));
  console.log("");

  // 4. Process each row
  let countFound = 0, countNotFound = 0, countSkipped = 0;

  for (let i = 0; i < batch.length; i++) {
    const row       = batch[i];
    const rowNumber = START_ROW + i;
    const progress  = `[${i + 1}/${batch.length}]`;

    // Get ISIN and sheet symbol for comparison
    const isin        = String(row.isin   ?? row.ISIN   ?? "").trim();
    const sheetSymbol = String(row.symbol ?? row.Symbol ?? "").trim();

    if (!isin || isin.toLowerCase() === "nan") {
      log("warn", `${progress} Row ${rowNumber} — missing ISIN, skipping`);
      resultsMap.set(rowNumber, { rowNumber, isin: "", sheet_symbol: sheetSymbol, finhub_symbol: null, reason: "missing_isin" });
      countSkipped++;
      continue;
    }

    log("info", `${progress} Row ${rowNumber} — ISIN=${isin}  sheet_symbol=${sheetSymbol}  looking up Finnhub...`);

    const result = await fetchFinnhubSymbol(isin);

    if (!result) {
      log("warn", `${progress} Row ${rowNumber} — NOT FOUND  (ISIN=${isin}  sheet_symbol=${sheetSymbol})`);
      resultsMap.set(rowNumber, { rowNumber, isin, sheet_symbol: sheetSymbol, finhub_symbol: null, found: false });
      countNotFound++;
    } else {
      const match = result.displaySymbol === sheetSymbol ? "MATCH" : "DIFF";
      log("ok", `${progress} Row ${rowNumber} — FOUND ✔  sheet_symbol=${sheetSymbol}  finhub_symbol=${result.displaySymbol}  [${match}]  desc="${result.description}"`);
      resultsMap.set(rowNumber, {
        rowNumber,
        isin,
        sheet_symbol:    sheetSymbol,
        finhub_symbol:   result.displaySymbol,
        finnhub_raw_sym: result.symbol,
        description:     result.description,
        type:            result.type,
        found:           true,
      });
      countFound++;
    }

    if (i < batch.length - 1) await sleep(DELAY_MS);
  }

  // 5. Save merged JSON
  const merged = Array.from(resultsMap.values()).sort((a, b) => a.rowNumber - b.rowNumber);
  fs.writeFileSync(jsonPath, JSON.stringify(merged, null, 2));
  log("ok", `Results file updated → ${jsonPath}  (${merged.length} total entries across all runs)`);

  // 6. Summary
  console.log("");
  console.log("━".repeat(60));
  log("info", "Summary:");
  log("ok",   `  Found      : ${countFound}`);
  log("warn",  `  Not found  : ${countNotFound}`);
  log("warn",  `  Skipped    : ${countSkipped}  (missing ISIN)`);
  console.log("━".repeat(60));
  console.log("");
  log("info", "Run  node sync-results.js  to write the Excel output.");
}

main().catch((err) => {
  log("error", `Unhandled error: ${err.message}`);
  console.error(err);
  process.exit(1);
});
