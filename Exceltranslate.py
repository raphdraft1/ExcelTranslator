"""
ExcelTranslator.py
------------------
Translates selected columns in an Excel file using Google Translate.
Supports auto language detection, merged cell preservation, caching,
retries, and optional PowerPoint translation.

Usage:
    python ExcelTranslator.py --input file.xlsx --output ./translated/
    python ExcelTranslator.py --input file.pptx --output ./translated/
    python ExcelTranslator.py --input file.xlsx --src zh-CN --dest en --all-columns
"""

import os
import time
import argparse
import logging

import pandas as pd
from tqdm import tqdm
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from deep_translator import GoogleTranslator
from deep_translator.exceptions import RequestError, TooManyRequests

# ── logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger(__name__)

# ── translator / cache ────────────────────────────────────────────────────────

_cache: dict[str, str] = {}


def translate_text(text: str, src: str, dest: str, retries: int = 3) -> str:
    """
    Translate a single string. Returns original on failure.
    Caches results to avoid re-translating identical strings.
    """
    if not isinstance(text, str) or not text.strip():
        return text

    cache_key = f"{src}:{dest}:{text}"
    if cache_key in _cache:
        return _cache[cache_key]

    for attempt in range(1, retries + 1):
        try:
            time.sleep(0.25)   # polite throttle
            result = GoogleTranslator(source=src, target=dest).translate(text)
            translated = result if result else text
            _cache[cache_key] = translated
            return translated

        except TooManyRequests:
            wait = 2 ** attempt          # exponential backoff: 2s, 4s, 8s
            log.warning("Rate limited — waiting %ds (attempt %d/%d)", wait, attempt, retries)
            time.sleep(wait)

        except RequestError as e:
            log.warning("Request error: %s (attempt %d/%d)", e, attempt, retries)
            time.sleep(1)

        except Exception as e:
            log.warning("Unexpected error: %s (attempt %d/%d)", e, attempt, retries)
            time.sleep(1)

    log.error("All retries failed for: %.60s", text)
    _cache[cache_key] = text
    return text


# ── Excel translation ─────────────────────────────────────────────────────────

def _select_columns(df: pd.DataFrame, all_columns: bool) -> list[str]:
    """Interactive column selection, or return all string columns."""
    string_cols = [c for c in df.columns if df[c].dtype == "object"]

    if all_columns:
        return string_cols

    print("\nAvailable columns:")
    for i, col in enumerate(df.columns, 1):
        tag = "(text)" if col in string_cols else "(non-text, skipped)"
        print(f"  {i:>2}. {col}  {tag}")

    raw = input("\nEnter column numbers to translate (comma-separated, e.g. 1,3): ")
    indices = [int(x.strip()) - 1 for x in raw.split(",") if x.strip().isdigit()]
    selected = [df.columns[i] for i in indices if 0 <= i < len(df.columns)]

    invalid = [df.columns[i] for i in indices
               if 0 <= i < len(df.columns) and df.columns[i] not in string_cols]
    if invalid:
        log.warning("Skipping non-text columns: %s", invalid)

    return [c for c in selected if c in string_cols]


def _translate_dataframe(df: pd.DataFrame, columns: list[str], src: str, dest: str) -> pd.DataFrame:
    out = df.copy()
    for col in tqdm(columns, desc="Translating columns"):
        out[col] = df[col].apply(
            lambda x: translate_text(x, src, dest) if isinstance(x, str) else x
        )
    return out


def translate_excel(input_path: str, output_dir: str, src: str, dest: str, all_columns: bool) -> str:
    """
    Translate an Excel file, preserving merged cells and formatting.
    Returns path to the saved output file.
    """
    log.info("Loading workbook: %s", input_path)
    excel_data = pd.ExcelFile(input_path)

    # Sheet selection
    print("\nAvailable sheets:")
    for i, name in enumerate(excel_data.sheet_names, 1):
        print(f"  {i:>2}. {name}")

    raw = input("\nEnter sheet number (or press Enter for all sheets): ").strip()
    if raw == "":
        sheet_names = excel_data.sheet_names
    else:
        idx = int(raw) - 1
        if not (0 <= idx < len(excel_data.sheet_names)):
            raise ValueError(f"Invalid sheet number: {idx + 1}")
        sheet_names = [excel_data.sheet_names[idx]]

    # Use openpyxl to preserve merged cells and formatting
    wb = load_workbook(input_path)

    for sheet_name in sheet_names:
        log.info("Processing sheet: %s", sheet_name)
        ws = wb[sheet_name]

        # Record merged cell ranges before we touch anything
        merged_ranges = list(ws.merged_cells.ranges)

        # Temporarily unmerge so we can iterate all cells
        for merge_range in merged_ranges:
            ws.unmerge_cells(str(merge_range))

        # Collect all string cells
        cells_to_translate = [
            (row, col, ws.cell(row=row, column=col).value)
            for row in range(1, ws.max_row + 1)
            for col in range(1, ws.max_column + 1)
            if isinstance(ws.cell(row=row, column=col).value, str)
        ]

        if not all_columns:
            # If not --all-columns, do column selection via pandas for the sheet
            df = excel_data.parse(sheet_name)
            selected = _select_columns(df, all_columns=False)
            # Map selected column names to 1-based indices
            selected_indices = set()
            header_row = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
            for col_name in selected:
                if col_name in header_row:
                    selected_indices.add(header_row.index(col_name) + 1)
            cells_to_translate = [(r, c, v) for r, c, v in cells_to_translate if c in selected_indices]

        # Translate and write back
        for row, col, value in tqdm(cells_to_translate, desc=f"  {sheet_name}"):
            ws.cell(row=row, column=col).value = translate_text(value, src, dest)

        # Re-merge
        for merge_range in merged_ranges:
            ws.merge_cells(str(merge_range))

    os.makedirs(output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(output_dir, f"{base}_translated.xlsx")
    wb.save(out_path)
    log.info("Saved: %s", out_path)
    return out_path


# ── PowerPoint translation ────────────────────────────────────────────────────

def translate_pptx(input_path: str, output_dir: str, src: str, dest: str) -> str:
    """
    Translate all text in a PowerPoint file, preserving formatting.
    Returns path to the saved output file.
    """
    try:
        from pptx import Presentation
        from pptx.util import Pt
    except ImportError:
        raise ImportError("python-pptx is required for PowerPoint support: pip install python-pptx")

    log.info("Loading presentation: %s", input_path)
    prs = Presentation(input_path)

    slides = list(prs.slides)
    for slide_num, slide in enumerate(tqdm(slides, desc="Translating slides"), 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        run.text = translate_text(run.text, src, dest)

    os.makedirs(output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(output_dir, f"{base}_translated.pptx")
    prs.save(out_path)
    log.info("Saved: %s", out_path)
    return out_path


# ── CLI ───────────────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Translate Excel (.xlsx) or PowerPoint (.pptx) files via Google Translate.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    p.add_argument("--input",  "-i", required=True,  help="Path to input file (.xlsx or .pptx)")
    p.add_argument("--output", "-o", default="./translated", help="Output directory (default: ./translated)")
    p.add_argument("--src",    "-s", default="auto", help="Source language code (default: auto-detect)")
    p.add_argument("--dest",   "-d", default="en",   help="Target language code (default: en)")
    p.add_argument("--all-columns", action="store_true",
                   help="Translate all text columns without prompting (Excel only)")
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        parser.error(f"File not found: {args.input}")

    ext = os.path.splitext(args.input)[1].lower()
    start = time.time()

    if ext in (".xlsx", ".xls", ".xlsm"):
        out = translate_excel(args.input, args.output, args.src, args.dest, args.all_columns)
    elif ext in (".pptx", ".ppt"):
        out = translate_pptx(args.input, args.output, args.src, args.dest)
    else:
        parser.error(f"Unsupported file type: {ext}  (supported: .xlsx, .xls, .xlsm, .pptx)")

    elapsed = time.time() - start
    print(f"\n✓  Done in {elapsed:.1f}s  →  {out}")
    print(f"   Cache hits saved ~{len(_cache)} repeat translations.")


if __name__ == "__main__":
    main()
