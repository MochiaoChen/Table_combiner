import argparse
import logging
import os
import re
import sys
from typing import List, Tuple

import pandas as pd

# ----------- Logging -----------
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s"
)

# ----------- Helpers -----------

EXCEL_EXTS = {".xlsx", ".xls"}
CSV_EXTS = {".csv", ".tsv", ".txt"}
MAX_SHEETNAME_LEN = 31


def list_input_files(folder: str) -> List[str]:
    all_files = []
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext in EXCEL_EXTS | CSV_EXTS:
            all_files.append(path)
    # Sort for deterministic order
    all_files.sort(key=lambda p: os.path.basename(p).lower())
    return all_files


def suffix_after_last_underscore(stem: str) -> str:
    parts = stem.split("_")
    return parts[-1] if len(parts) > 1 else stem


def sanitize_sheet_name(name: str) -> str:
    """
    Remove characters illegal for Excel sheet names and trim spaces.
    """
    name = re.sub(r'[:\\/\?\*\[\]]', " ", name).strip()
    # Replace consecutive spaces with single space
    name = re.sub(r"\s+", " ", name)
    return name or "Sheet"


def uniquify_and_truncate(names: List[str], max_len: int = MAX_SHEETNAME_LEN) -> List[str]:
    seen = {}
    result = []
    for raw in names:
        base = sanitize_sheet_name(raw)[:max_len]
        candidate = base if base else "Sheet"
        idx = 1
        while candidate.lower() in seen:
            suffix = f"_{idx}"
            candidate = (base[: max_len - len(suffix)] + suffix) if len(base) + len(suffix) > max_len else base + suffix
            idx += 1
        seen[candidate.lower()] = True
        result.append(candidate)
    return result


def read_csv_like(path: str) -> pd.DataFrame:
    # Try utf-8 first, fallback to gbk
    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        logging.warning(f"UTF-8 failed for {os.path.basename(path)}, falling back to GBK.")
        return pd.read_csv(path, encoding="gbk")


def read_excel_all_sheets(path: str) -> List[Tuple[str, pd.DataFrame]]:
    xls = pd.ExcelFile(path)
    items = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        items.append((sheet, df))
    return items


def collect_tables(folder: str) -> Tuple[List[pd.DataFrame], List[str]]:
    """
    Returns:
        dfs: list of DataFrames, in deterministic order
        planned_names: proposed sheet names before uniquify/limit
    """
    files = list_input_files(folder)
    if not files:
        logging.error("No supported files (.xlsx/.xls/.csv/.tsv/.txt) found in the folder.")
        sys.exit(1)

    dfs: List[pd.DataFrame] = []
    planned_names: List[str] = []

    for path in files:
        fname = os.path.basename(path)
        stem, ext = os.path.splitext(fname)
        ext = ext.lower()

        if ext in EXCEL_EXTS:
            sheets = read_excel_all_sheets(path)
            if len(sheets) == 1:
                # Single-sheet Excel -> use filename as sheet name
                original_name, df = sheets[0]
                dfs.append(df)
                planned_names.append(stem)
                logging.info(f"[Excel-1] {fname} -> sheet '{stem}'")
            else:
                # Multi-sheet Excel -> use original sheet names
                for original_name, df in sheets:
                    dfs.append(df)
                    planned_names.append(original_name)
                logging.info(f"[Excel-N] {fname} -> {len(sheets)} sheets preserved by original names")

        elif ext in CSV_EXTS:
            # Single sheet per CSV
            df = read_csv_like(path)
            dfs.append(df)
            planned_names.append(stem)
            logging.info(f"[CSV] {fname} -> sheet '{stem}'")

    # Enforce uniqueness + length limits
    final_names = uniquify_and_truncate(planned_names, MAX_SHEETNAME_LEN)

    # Log any renames due to collisions/length
    for raw, final in zip(planned_names, final_names):
        if sanitize_sheet_name(raw)[:MAX_SHEETNAME_LEN] != final:
            logging.warning(f"Sheet name adjusted: '{raw}' -> '{final}'")

    return dfs, final_names


def save_to_excel(dfs: List[pd.DataFrame], sheet_names: List[str], out_path: str) -> None:
    # Ensure parent folder exists
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for df, name in zip(dfs, sheet_names):
            # Guarantee at least one column to avoid writer issues
            if df.shape[1] == 0:
                df = df.copy()
                df["(empty)"] = []
            df.to_excel(writer, sheet_name=name, index=False)
    logging.info(f"Saved merged workbook -> {out_path}")


# ----------- CLI -----------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Merge all tables in a folder into a single Excel workbook with controlled sheet naming."
    )
    parser.add_argument(
        "-i", "--input-folder", required=True, help="Path to the folder containing tables (.xlsx/.xls/.csv/.tsv/.txt)."
    )
    parser.add_argument(
        "-o", "--output-name", required=True, help="Output Excel filename (e.g., 问题3.xlsx). The file is saved into the input folder."
    )
    parser.add_argument(
        "-q", "--quiet", action="store_true", help="Suppress info logs (only warnings/errors)."
    )
    return parser.parse_args()


def main():
    args = parse_args()
    folder = os.path.abspath(args.input_folder)
    if args.quiet:
        logging.getLogger().setLevel(logging.WARNING)

    if not os.path.isdir(folder):
        logging.error(f"Input folder not found: {folder}")
        sys.exit(1)

    dfs, names = collect_tables(folder)
    out_path = os.path.join(folder, args.output_name)
    save_to_excel(dfs, names, out_path)


if __name__ == "__main__":
    main()
