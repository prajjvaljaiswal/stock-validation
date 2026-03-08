/**
 * Manual GTN symbol-validation worker (Node.js)
 * ─────────────────────────────────────────────
 * Edit START_ROW / END_ROW below, then run:
 *   node fetch-worker.js
 *
 * Results are written to:
 *   volume/results/results_manual.json   (raw per-row data)
 *   volume/output_manual.xlsx            (original sheet + is_correct column)
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────────

const START_ROW = 100;   // inclusive, 0-based (first data row after header)
const END_ROW   = 149;  // exclusive  → processes rows 100..148

const EXCEL_FILE  = "./volume/New stocks to be added to the universe.xlsx";
const SHEET_NAME  = "filtered by exchange";
const RESULTS_DIR = "./volume/results";

const BEARER_TOKEN = 
"O0AeyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJHVE4iLCJyb2xlIjoiY3VzdG9tZXIiLCJodWIiOiJESUZDIiwicHJvdmlkZXIiOiJHVE4iLCJjaGFubmVsIjoiNDYiLCJpbnN0Q29kZSI6Ik5VUUkgTUFVUklUSVVTIiwiY3VzdG9tZXJOdW1iZXIiOiI5MzgzMjM2NTIiLCJ2ZXJzaW9uIjoidjEiLCJleHAiOjE3NzI5NjYwOTUsImlhdCI6MTc3Mjk2MjE5NSwianRpIjoiZWVmZDZlOTQtZjE1NS00NzMzLWI2MzQtZjQ1ZjQ0NzNlNTc2In0.15lpu7xqkkLenQI8y9_X398L-vmU6zD68CYAxk33at8";
const THROTTLE_KEY = "ptmMH4AY8_LZsW5KbpvtUZzQNSAa";

const DELAY_MS = 300; // delay between requests to avoid rate-limiting

// ─── DEPENDENCIES ────────────────────────────────────────────────────────────

const fs   = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function log(level, message) {
  const time = new Date().toISOString().slice(11, 19); // HH:MM:SS
  const prefix = { info: "ℹ", ok: "✔", warn: "⚠", error: "✘" }[level] ?? "·";
  console.log(`[${time}] ${prefix}  ${message}`);
}

/**
 * Call GTN symbol-search API for a given exchange + symbol.
 * Returns parsed JSON or null on failure.
 */
async function fetchSymbol(exchange, symbol) {
  const url = new URL("https://api-mena.globaltradingnetwork.com/market-data/symbol-search");
  url.searchParams.set("keys", symbol);
  url.searchParams.set("search-symbol-code", "true");
  url.searchParams.set("exchanges", exchange);
  url.searchParams.set("lang", "EN");

  try {
    const res = await fetch(url.toString(), {
      headers: {
        Authorization: `Bearer ${BEARER_TOKEN}`,
        "Throttle-Key": THROTTLE_KEY,
      },
    });

    if (!res.ok) {
      log("warn", `HTTP ${res.status} for ${exchange}~${symbol}`);
      return null;
    }

    return await res.json();
  } catch (err) {
    log("error", `Network error for ${exchange}~${symbol}: ${err.message}`);
    return null;
  }
}

/**
 * Determine is_correct from the API response.
 * True if any returned doc matches by ISIN (primary) or SYMBOL+EXCHANGE (fallback).
 */
function checkCorrect(row, response) {
  if (!response) return false;

  const docs = response?.response?.docs ?? [];
  if (docs.length === 0) return false;

  const rowIsin     = String(row.isin     ?? row.ISIN     ?? "").trim().toUpperCase();
  const rowSymbol   = String(row.symbol   ?? "").trim().toUpperCase();
  const rowExchange = String(row.exchange ?? "").trim().toUpperCase();

  for (const doc of docs) {
    const docIsin     = String(doc.ISIN_CODE ?? "").trim().toUpperCase();
    const docSymbol   = String(doc.SYMBOL    ?? "").trim().toUpperCase();
    const docExchange = String(doc.EXCHANGE  ?? "").trim().toUpperCase();

    if (rowIsin && docIsin && docIsin === rowIsin) return true;
    if (docSymbol === rowSymbol && docExchange === rowExchange) return true;
  }

  return false;
}

// ─── MAIN ────────────────────────────────────────────────────────────────────

async function main() {
  console.log("━".repeat(60));
  log("info", `GTN Symbol Validator — Node.js manual runner`);
  log("info", `Rows: ${START_ROW} → ${END_ROW - 1}  (${END_ROW - START_ROW} total)`);
  log("info", `Excel: ${EXCEL_FILE}`);
  console.log("━".repeat(60));

  // 1. Load workbook
  log("info", "Loading Excel file...");
  if (!fs.existsSync(EXCEL_FILE)) {
    log("error", `Excel file not found: ${EXCEL_FILE}`);
    process.exit(1);
  }
  const wb   = XLSX.readFile(EXCEL_FILE);
  const ws   = wb.Sheets[SHEET_NAME];
  if (!ws) {
    log("error", `Sheet "${SHEET_NAME}" not found. Available: ${wb.SheetNames.join(", ")}`);
    process.exit(1);
  }
  const allRows = XLSX.utils.sheet_to_json(ws);
  log("ok", `Loaded ${allRows.length} rows from sheet "${SHEET_NAME}"`);

  // 2. Slice the batch
  const batch = allRows.slice(START_ROW, END_ROW);
  if (batch.length === 0) {
    log("warn", "No rows in the specified range. Check START_ROW / END_ROW.");
    process.exit(0);
  }
  log("info", `Processing ${batch.length} rows (${START_ROW}–${START_ROW + batch.length - 1})`);
  console.log("");

  // 3. Load existing results file (if any) so we append/update, not overwrite
  fs.mkdirSync(RESULTS_DIR, { recursive: true });
  const jsonPath = path.join(RESULTS_DIR, "results_manual.json");
  let existingResults = [];
  if (fs.existsSync(jsonPath)) {
    existingResults = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
    log("info", `Found existing results file with ${existingResults.length} entries — will update in place`);
  } else {
    log("info", "No existing results file found — creating new one");
  }
  // Build a map keyed by rowNumber for fast lookup/update
  const resultsMap = new Map(existingResults.map((r) => [r.rowNumber, r]));

  // 4. Process each row
  const results = [];
  let countTrue = 0, countFalse = 0, countSkipped = 0;

  for (let i = 0; i < batch.length; i++) {
    const row       = batch[i];
    const rowNumber = START_ROW + i;
    const key       = String(row.key ?? "").trim();
    const progress  = `[${i + 1}/${batch.length}]`;

    if (!key || key.toLowerCase() === "nan" || !key.includes("~")) {
      log("warn", `${progress} Row ${rowNumber} — missing/invalid key "${key}", skipping`);
      const entry = { rowNumber, key, is_correct: false, reason: "invalid_key" };
      resultsMap.set(rowNumber, entry);
      results.push(entry);
      countSkipped++;
      continue;
    }

    const [exchange, symbol] = key.split("~").map((s) => s.trim());
    log("info", `${progress} Row ${rowNumber} — fetching  exchange=${exchange}  symbol=${symbol}`);

    const response = await fetchSymbol(exchange, symbol);

    if (!response) {
      log("error", `${progress} Row ${rowNumber} — no response, marking false`);
      const entry = { rowNumber, key, exchange, symbol, is_correct: false, reason: "no_response" };
      resultsMap.set(rowNumber, entry);
      results.push(entry);
      countFalse++;
    } else {
      const numFound   = response?.response?.numFound ?? 0;
      const isCorrect  = checkCorrect(row, response);
      const tag        = isCorrect ? "ok" : "warn";
      const label      = isCorrect ? "CORRECT ✔" : "NOT FOUND ✘";
      log(tag, `${progress} Row ${rowNumber} — ${label}  (numFound=${numFound})`);

      const entry = { rowNumber, key, exchange, symbol, is_correct: isCorrect, numFound };
      resultsMap.set(rowNumber, entry);
      results.push(entry);
      isCorrect ? countTrue++ : countFalse++;
    }

    if (i < batch.length - 1) await sleep(DELAY_MS);
  }

  // 5. Save merged JSON results (all runs combined, sorted by rowNumber)
  const mergedResults = Array.from(resultsMap.values()).sort((a, b) => a.rowNumber - b.rowNumber);
  fs.writeFileSync(jsonPath, JSON.stringify(mergedResults, null, 2));
  log("ok", `Results file updated → ${jsonPath}  (${mergedResults.length} total entries across all runs)`);

  // 6. Summary
  console.log("");
  console.log("━".repeat(60));
  log("info", `Summary:`);
  log("ok",   `  Correct   : ${countTrue}`);
  log("warn",  `  Not found : ${countFalse}`);
  log("warn",  `  Skipped   : ${countSkipped}`);
  console.log("━".repeat(60));
}

main().catch((err) => {
  log("error", `Unhandled error: ${err.message}`);
  console.error(err);
  process.exit(1);
});
