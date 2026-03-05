#!/usr/bin/env python3
"""
ExcelMaster — Clean messy business datasets.

Features:
    - Load CSV or Excel files
    - Remove duplicate rows
    - Standardize date columns
    - Normalize currency values
    - Handle missing values
    - Export cleaned dataset

Usage:
    python scripts/excel_master.py
    python scripts/excel_master.py --input examples/messy_data.csv
"""

import json
import logging
import re
import sys
from pathlib import Path
from typing import Optional

import pandas as pd

# Ensure Unicode output works on Windows consoles (cp1252, etc.)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Project root — always resolve relative to this script's parent
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def load_config(config_path: Optional[Path] = None) -> dict:
    """Load settings from config/config.json."""
    path = config_path or PROJECT_ROOT / "config" / "config.json"
    try:
        with open(path, "r", encoding="utf-8") as fh:
            config = json.load(fh)
        logger.info("Configuration loaded from %s", path)
        return config
    except FileNotFoundError:
        logger.error("Config file not found at %s", path)
        sys.exit(1)
    except json.JSONDecodeError as exc:
        logger.error("Invalid JSON in config file: %s", exc)
        sys.exit(1)


def load_data(file_path: Path) -> pd.DataFrame:
    """Load a CSV or Excel file into a DataFrame."""
    suffix = file_path.suffix.lower()
    try:
        if suffix in (".xlsx", ".xls"):
            df = pd.read_excel(file_path, engine="openpyxl")
        elif suffix == ".csv":
            df = pd.read_csv(file_path)
        else:
            logger.error("Unsupported file format: %s", suffix)
            sys.exit(1)
        logger.info("Loaded %d rows from %s", len(df), file_path.name)
        return df
    except Exception as exc:
        logger.error("Failed to load data: %s", exc)
        sys.exit(1)


def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Remove duplicate rows and report count."""
    before = len(df)
    df = df.drop_duplicates()
    removed = before - len(df)
    if removed:
        print(f"  ✔ {removed} duplicate row(s) removed")
    else:
        print("  ✔ No duplicate rows found")
    logger.info("Duplicates removed: %d", removed)
    return df


def _parse_date(value: str) -> Optional[str]:
    """Attempt to parse a date string into YYYY-MM-DD."""
    if pd.isna(value):
        return value
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d",
                "%b %d %Y", "%B %d %Y", "%b %d, %Y", "%d-%m-%Y"):
        try:
            return pd.to_datetime(str(value), format=fmt).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue
    # Fallback — let pandas guess
    try:
        return pd.to_datetime(str(value)).strftime("%Y-%m-%d")
    except Exception:
        return value


def standardize_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Detect date-like columns and standardize to YYYY-MM-DD."""
    date_cols = [c for c in df.columns if "date" in c.lower()]
    for col in date_cols:
        df[col] = df[col].apply(_parse_date)
        print(f"  ✔ Date column '{col}' standardized to YYYY-MM-DD")
    if not date_cols:
        print("  ℹ No date columns detected")
    return df


def normalize_currency(df: pd.DataFrame, symbol: str = "₹") -> pd.DataFrame:
    """Strip currency symbols and thousands separators; convert to float."""
    currency_cols = [c for c in df.columns if "amount" in c.lower()
                     or "price" in c.lower() or "cost" in c.lower()]
    for col in currency_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(symbol, "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")
        print(f"  ✔ Currency column '{col}' normalized to plain numbers")
    if not currency_cols:
        print("  ℹ No currency columns detected")
    return df


def handle_missing_values(df: pd.DataFrame) -> pd.DataFrame:
    """Fill missing values with sensible defaults."""
    missing_before = int(df.isna().sum().sum())
    # Numeric columns → 0, string columns → "N/A"
    for col in df.columns:
        if df[col].dtype in ("float64", "int64"):
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("N/A")
    missing_after = int(df.isna().sum().sum())
    filled = missing_before - missing_after
    print(f"  ✔ {filled} missing value(s) filled")
    logger.info("Missing values filled: %d", filled)
    return df


def main() -> None:
    """Entry-point: clean a data file end-to-end."""
    import argparse

    parser = argparse.ArgumentParser(description="ExcelMaster — Clean messy datasets")
    parser.add_argument("--input", type=str, help="Path to the input CSV/Excel file")
    parser.add_argument("--output", type=str, help="Path for the cleaned output file")
    args = parser.parse_args()

    config = load_config()
    input_folder = PROJECT_ROOT / config.get("input_folder", "examples/")
    output_folder = PROJECT_ROOT / config.get("output_folder", "output/")
    output_folder.mkdir(parents=True, exist_ok=True)
    currency_symbol: str = config.get("currency_symbol", "₹")

    # Resolve input file
    if args.input:
        input_file = Path(args.input)
        if not input_file.is_absolute():
            input_file = PROJECT_ROOT / input_file
    else:
        input_file = input_folder / "messy_data.csv"

    output_file = Path(args.output) if args.output else output_folder / "cleaned_data.csv"

    print("\n🔧 ExcelMaster — Data Cleaning Tool")
    print("=" * 42)
    print(f"  Input : {input_file}")
    print(f"  Output: {output_file}\n")

    df = load_data(input_file)

    print("▸ Removing duplicates …")
    df = remove_duplicates(df)

    print("▸ Standardizing dates …")
    df = standardize_dates(df)

    print("▸ Normalizing currency …")
    df = normalize_currency(df, symbol=currency_symbol)

    print("▸ Handling missing values …")
    df = handle_missing_values(df)

    # Export
    df.to_csv(output_file, index=False)
    print(f"\n✔ Cleaned data saved to {output_file}")
    print(f"  Rows: {len(df)}  |  Columns: {len(df.columns)}")
    logger.info("Output written to %s", output_file)


if __name__ == "__main__":
    main()
