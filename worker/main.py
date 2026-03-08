import os
import json
import time
import requests
import pandas as pd
from pathlib import Path

EXCEL_FILE = "/data/New stocks to be added to the universe.xlsx"
SHEET_NAME = "filtered by exchange"
RESULTS_DIR = "/data/results"

START_ROW   = int(os.environ["START_ROW"])
END_ROW     = int(os.environ["END_ROW"])
WORKER_ID   = os.environ.get("WORKER_ID", "0")
BEARER      = os.environ["BEARER_TOKEN"]
THROTTLE    = os.environ["THROTTLE_KEY"]
PROXY_URL   = os.environ.get("PROXY_URL", "").strip() or None

BASE_URL = "https://api-mena.globaltradingnetwork.com/market-data/symbol-search"
HEADERS = {
    "Authorization": f"Bearer {BEARER}",
    "Throttle-Key": THROTTLE,
}
PROXIES = {"http": PROXY_URL, "https": PROXY_URL} if PROXY_URL else None


def fetch_symbol(exchange: str, symbol: str) -> dict | None:
    params = {
        "keys": symbol,
        "search-symbol-code": "true",
        "exchanges": exchange,
        "lang": "EN",
    }
    try:
        resp = requests.get(
            BASE_URL, params=params, headers=HEADERS,
            proxies=PROXIES, timeout=30
        )
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print(f"[worker-{WORKER_ID}] ERROR fetching {exchange}~{symbol}: {e}")
        return None


def check_correct(row: dict, response: dict | None) -> bool:
    if not response:
        return False
    docs = response.get("response", {}).get("docs", [])
    if not docs:
        return False

    row_isin     = str(row.get("isin") or row.get("ISIN") or "").strip().upper()
    row_symbol   = str(row.get("symbol") or "").strip().upper()
    row_exchange = str(row.get("exchange") or "").strip().upper()

    for doc in docs:
        doc_isin     = str(doc.get("ISIN_CODE", "")).strip().upper()
        doc_symbol   = str(doc.get("SYMBOL", "")).strip().upper()
        doc_exchange = str(doc.get("EXCHANGE", "")).strip().upper()

        # Primary check: ISIN match
        if row_isin and doc_isin and doc_isin == row_isin:
            return True
        # Fallback: exact symbol + exchange match
        if doc_symbol == row_symbol and doc_exchange == row_exchange:
            return True

    return False


def main():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    batch = df.iloc[START_ROW:END_ROW]

    Path(RESULTS_DIR).mkdir(parents=True, exist_ok=True)

    results = []
    total = len(batch)

    for i, (idx, row) in enumerate(batch.iterrows(), 1):
        key = str(row.get("key", "")).strip()
        print(f"[worker-{WORKER_ID}] {i}/{total} | row {idx} | key={key}")

        if not key or key.lower() == "nan" or "~" not in key:
            results.append({"row_index": int(idx), "is_correct": False})
            continue

        exchange, symbol = key.split("~", 1)
        exchange, symbol = exchange.strip(), symbol.strip()

        response = fetch_symbol(exchange, symbol)
        correct  = check_correct(row.to_dict(), response)
        results.append({"row_index": int(idx), "is_correct": correct})

        time.sleep(0.3)  # gentle pacing per container

    output_path = f"{RESULTS_DIR}/results_{WORKER_ID}.json"
    with open(output_path, "w") as f:
        json.dump(results, f)

    print(f"[worker-{WORKER_ID}] Done → {output_path}")


if __name__ == "__main__":
    main()
