"""
Cars TAM Agent
Collects Israeli annual car sales (2020-2026) for electrified models
and writes results to a new Excel file preserving the original format.

Usage:
    python agent.py                  # full run
    python agent.py --dry-run        # load models, skip scraping (zeros)
    python agent.py --model "Toyota Camry"   # run single model only
"""

import argparse
import json
import logging
import os
import sys
from pathlib import Path

import openpyxl

from config import SOURCE_FILE, OUTPUT_FILE, YEARS
from excel_handler import read_models, build_output
from scraper import get_sales_data, MODEL_PATHS

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

CHECKPOINT_FILE = Path(__file__).parent / "output" / "checkpoint.json"


def load_checkpoint() -> dict:
    if CHECKPOINT_FILE.exists():
        with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_checkpoint(data: dict):
    CHECKPOINT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def run(dry_run: bool = False, only_model: str = None):
    # 1. Load source models
    logger.info(f"Reading source file: {SOURCE_FILE}")
    src_wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)
    models = read_models(src_wb)
    logger.info(f"Found {len(models)} electrified models (HEV/PHEV/BEV, MHEV→HEV)")

    if only_model:
        query = only_model.lower()
        models = [m for m in models
                  if query in f"{m['vendor']} {m['model']}".lower()]
        if not models:
            logger.error(f"No model matched '{only_model}'")
            sys.exit(1)
        logger.info(f"Filtered to {len(models)} model(s) matching '{only_model}'")

    # 2. Collect sales data (with checkpoint resume)
    checkpoint = load_checkpoint()
    results = []
    total = len(models)

    for idx, m in enumerate(models, 1):
        key = str(m["row_num"])
        logger.info(f"[{idx}/{total}] {m['vendor']} {m['model']} ({m['type']})")

        if key in checkpoint:
            logger.info(f"  → loaded from checkpoint")
            sales = {int(k): v for k, v in checkpoint[key].items()}
        elif dry_run:
            sales = {year: None for year in YEARS}
        else:
            try:
                sales = get_sales_data(m["row_num"])
            except Exception as e:
                logger.warning(f"  → scrape error: {e}  (stored as empty)")
                sales = {year: None for year in YEARS}

            checkpoint[key] = sales
            save_checkpoint(checkpoint)

        results.append({**m, "sales": sales})

    # 3. Build output Excel
    logger.info("Building output Excel...")
    out_path = build_output(results)
    logger.info(f"Done → {out_path}")

    # 4. Summary
    _print_summary(results)


def _print_summary(results: list[dict]):
    print("\n" + "=" * 70)
    print(f"{'#':<4} {'Vendor':<14} {'Model':<25} {'Type':<6} {'Years found'}")
    print("-" * 70)
    for r in results:
        found = sum(1 for v in r["sales"].values() if v is not None)
        years_str = ", ".join(
            str(y) for y in sorted(YEARS)
            if r["sales"].get(y) is not None
        ) or "none"
        print(f"{r['row_num']:<4} {r['vendor']:<14} {str(r['model']):<25} {r['type']:<6} {years_str}")
    print("=" * 70)
    total = len(results)
    full = sum(1 for r in results if all(r["sales"].get(y) is not None for y in YEARS))
    partial = sum(1 for r in results if any(r["sales"].get(y) is not None for y in YEARS)) - full
    empty = total - full - partial
    print(f"\nTotal: {total}  |  Full data: {full}  |  Partial: {partial}  |  No data: {empty}")
    print(f"Output: {OUTPUT_FILE}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Cars TAM Agent")
    parser.add_argument("--dry-run", action="store_true",
                        help="Skip web scraping, write empty sales columns")
    parser.add_argument("--model", type=str, default=None,
                        help="Run only for a specific model (substring match)")
    parser.add_argument("--clear-checkpoint", action="store_true",
                        help="Delete checkpoint and start fresh")
    args = parser.parse_args()

    if args.clear_checkpoint and CHECKPOINT_FILE.exists():
        CHECKPOINT_FILE.unlink()
        logger.info("Checkpoint cleared")

    run(dry_run=args.dry_run, only_model=args.model)
