"""
Run after all workers complete:
  docker compose run --rm merger
Reads all results_*.json files from /data/results/ and writes
/data/output_with_results.xlsx with the original data plus is_correct column.
"""

import json
import glob
import pandas as pd
from pathlib import Path

EXCEL_FILE   = "/data/New stocks to be added to the universe.xlsx"
SHEET_NAME   = "filtered by exchange"
RESULTS_DIR  = "/data/results"
OUTPUT_FILE  = "/data/output_with_results.xlsx"


def main():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    df["is_correct"] = None  # default: not processed

    result_files = sorted(glob.glob(f"{RESULTS_DIR}/results_*.json"))
    if not result_files:
        print("No result files found in", RESULTS_DIR)
        return

    print(f"Merging {len(result_files)} result file(s)...")
    for rf in result_files:
        with open(rf) as f:
            records = json.load(f)
        for rec in records:
            df.at[rec["row_index"], "is_correct"] = rec["is_correct"]

    # Report counts
    true_count  = df["is_correct"].sum()
    false_count = (df["is_correct"] == False).sum()
    null_count  = df["is_correct"].isna().sum()
    print(f"is_correct: True={true_count}, False={false_count}, unprocessed={null_count}")

    # Copy original sheets + add results sheet
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="results", index=False)

    print(f"Output written → {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
