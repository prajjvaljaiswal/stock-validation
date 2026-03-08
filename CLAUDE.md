# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Parallel Docker-based tool that validates stock symbols against the GTN MENA market-data API. It reads `volume/New stocks to be added to the universe.xlsx` (sheet: "filtered by exchange", 667 rows), fans work out across 14 worker containers (50 rows each), then merges results into `volume/output_with_results.xlsx` with an added `is_correct` column.

## Commands

```bash
# Build images
docker compose build

# Run all workers + auto-merge when all succeed
docker compose up

# Run only a single worker (useful for debugging)
docker compose run --rm worker_0

# Run merger manually (if workers already completed)
docker compose run --rm merger
```

## Architecture

```
docker-compose.yml
├── worker_0 … worker_13   (14 containers, each 50 rows)
│   └── worker/main.py     reads Excel → calls GTN API → writes /data/results/results_N.json
└── merger                 worker/merge.py → reads all JSONs → writes output_with_results.xlsx
```

**Volume mount:** `./volume:/data` — the Excel file must be at `volume/New stocks to be added to the universe.xlsx`.

**Key column format:** `EXCHANGE~SYMBOL` (e.g. `ADSM~ADNH`). Splits on `~`, calls:
```
GET /market-data/symbol-search?keys=<SYMBOL>&exchanges=<EXCHANGE>&search-symbol-code=true&lang=EN
```

**is_correct logic (`worker/main.py → check_correct`):**
1. `numFound > 0` in response
2. Any doc's `ISIN_CODE` matches row's `isin` column (primary), or `SYMBOL`+`EXCHANGE` match exactly (fallback)

**Rate limiting / IP isolation:** Each worker accepts a `PROXY_URL` env var (set in `.env` as `PROXY_URL_0` … `PROXY_URL_13`). Without proxies all containers share the host's outgoing IP.

## Credentials

Stored in `.env` — `BEARER_TOKEN` and `THROTTLE_KEY`. Token has a short TTL; update `.env` when it expires.
