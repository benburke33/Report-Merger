#!/usr/bin/env python3
"""Merge weekly CSV/Excel reports into one formatted Excel summary.

Usage:
    python merge_reports.py weekly_reports/ -o summary.xlsx
    python merge_reports.py week1.xlsx week2.csv week3.csv --group-by region
    python merge_reports.py reports/ --group-by "region,product"

Output workbook:
  * Combined : every row from every file, plus a source_file column, totals row
  * Summary  : totals of each numeric column (per group if --group-by is given)
  * By File  : totals per source file, so you can compare weeks
Requires: pandas, openpyxl.
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}
HEADER_FILL = PatternFill("solid", start_color="1F4E78")
TOTAL_FILL = PatternFill("solid", start_color="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")


def collect_files(inputs):
    files = []
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            files += sorted(f for f in p.iterdir()
                            if f.suffix.lower() in EXCEL_EXT | {".csv"} and not f.name.startswith(("~", ".")))
        elif p.exists():
            files.append(p)
        else:
            print(f"Skipping missing path: {p}", file=sys.stderr)
    return files


def read_any(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in EXCEL_EXT:
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    df.insert(0, "source_file", path.name)
    return df


def style_sheet(ws, df, total_row=False, number_format="#,##0.00"):
    ncols, nrows = df.shape[1], df.shape[0]
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for idx, col in enumerate(df.columns, start=1):
        if pd.api.types.is_numeric_dtype(df[col]):
            values = df[col].dropna().astype(float)
            is_int = bool((values % 1 == 0).all())        # whole numbers (e.g. units) -> no decimals
            fmt = "#,##0" if is_int else number_format
            decimals = 0 if is_int else 2
            # measure the text Excel will actually display (commas included), incl. the total
            shown = [f"{v:,.{decimals}f}" for v in values.head(500)]
            if total_row and len(values):
                shown.append(f"{values.sum():,.{decimals}f}")
            for r in range(2, nrows + 2 + (1 if total_row else 0)):
                ws.cell(row=r, column=idx).number_format = fmt
        else:
            shown = [str(v) for v in df[col].head(500)]
        # header gets +4 so the filter dropdown arrow doesn't cover the title
        width = max([len(str(col)) + 4, 8] + [len(s) for s in shown]) + 2
        ws.column_dimensions[get_column_letter(idx)].width = min(width, 45)
    if total_row:
        for c in range(1, ncols + 1):
            cell = ws.cell(row=nrows + 2, column=c)
            cell.font = Font(bold=True)
            cell.fill = TOTAL_FILL
            cell.border = Border(top=THIN)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{nrows + 1}"


def add_sum_row(ws, df):
    """Append a TOTAL row using live Excel SUM formulas for numeric columns."""
    last = df.shape[0] + 1
    total_r = last + 1
    ws.cell(row=total_r, column=1, value="TOTAL")
    for idx, col in enumerate(df.columns, start=1):
        if pd.api.types.is_numeric_dtype(df[col]) and idx > 1:
            L = get_column_letter(idx)
            ws.cell(row=total_r, column=idx, value=f"=SUM({L}2:{L}{last})")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+", help="files and/or folders")
    ap.add_argument("-o", "--output", default="merged_summary.xlsx")
    ap.add_argument("--group-by", help="comma-separated column(s) to group the Summary sheet by")
    args = ap.parse_args()

    files = collect_files(args.inputs)
    if not files:
        sys.exit("No CSV/Excel files found.")
    out_path = Path(args.output).resolve()
    files = [f for f in files if f.resolve() != out_path]

    frames = [read_any(f) for f in files]
    combined = pd.concat(frames, ignore_index=True, sort=False)

    # Coerce columns that are mostly numeric (e.g. text-formatted numbers)
    for col in combined.columns:
        if col == "source_file" or pd.api.types.is_numeric_dtype(combined[col]):
            continue
        cleaned = combined[col].astype(str).str.replace(r"[\$,\s]", "", regex=True)
        nums = pd.to_numeric(cleaned.where(combined[col].notna()), errors="coerce")
        if combined[col].notna().sum() and nums.notna().sum() / combined[col].notna().sum() >= 0.9:
            combined[col] = nums

    numeric_cols = [c for c in combined.columns if pd.api.types.is_numeric_dtype(combined[c])]
    if not numeric_cols:
        sys.exit("No numeric columns found to total.")

    keys = [k.strip() for k in args.group_by.split(",")] if args.group_by else []
    missing = [k for k in keys if k not in combined.columns]
    if missing:
        sys.exit(f"--group-by column(s) not found: {missing}. Available: {list(combined.columns)}")

    if keys:
        summary = combined.groupby(keys, dropna=False)[numeric_cols].sum().reset_index()
    else:
        summary = combined[numeric_cols].sum().to_frame("total").reset_index()
        summary.columns = ["metric", "total"]
    by_file = combined.groupby("source_file")[numeric_cols].sum().reset_index()

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        combined.to_excel(xw, sheet_name="Combined", index=False)
        summary.to_excel(xw, sheet_name="Summary", index=False)
        by_file.to_excel(xw, sheet_name="By File", index=False)

        add_sum_row(xw.sheets["Combined"], combined)
        style_sheet(xw.sheets["Combined"], combined, total_row=True)
        if keys:
            add_sum_row(xw.sheets["Summary"], summary)
            style_sheet(xw.sheets["Summary"], summary, total_row=True)
        else:
            style_sheet(xw.sheets["Summary"], summary)
        add_sum_row(xw.sheets["By File"], by_file)
        style_sheet(xw.sheets["By File"], by_file, total_row=True)
        xw.book.calculation.fullCalcOnLoad = True   # make Excel calculate the TOTAL formulas on open

    print(f"Merged {len(files)} file(s), {len(combined)} rows -> {out_path}")


if __name__ == "__main__":
    main()