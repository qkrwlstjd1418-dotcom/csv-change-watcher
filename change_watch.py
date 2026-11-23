#!/usr/bin/env python3
"""change_watch.py — compare yesterday's export with today's and list what actually changed.

Free script, free to use and share. The other five scripts in the Python
Automation Pack work the same way:
    https://siljeondatalab.gumroad.com/l/python-automation-pack

Why this exists
    "Did anything move overnight?" is answered today by opening two files and
    scrolling. The three things people miss are a row that appeared, a row that
    vanished (usually the interesting one), and an amount that changed quietly.

What it does
    - Keys both snapshots on one or more columns (default: the first column).
    - Writes four sheets: new rows, removed rows, changed fields (before, after,
      difference), and a summary.
    - Appends every run to change_log.csv, so you can see when a value moved,
      not just that it did.
    - Exits with code 2 when something changed, so a scheduler can act on it.

Run
    python change_watch.py
    python change_watch.py --previous monday.csv --current tuesday.csv \\
        --key sku --compare price,stock

Needs: pandas, openpyxl  (pip install -r requirements.txt)
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CONFIG = {
    "previous": "sample_data/snapshot_monday.csv",
    "current": "sample_data/snapshot_tuesday.csv",
    "key": "",                   # empty = first column
    "compare": "",               # empty = every non-key column
    "output_file": "output/change_report_{date}.xlsx",
    "log_file": "output/change_log.csv",
    "numeric_tolerance": 0.005,  # differences below this count as unchanged
}


def read_table(path: Path) -> pd.DataFrame:
    if not path.is_file():
        sys.exit(f"File not found: {path}")
    if path.suffix.lower() == ".csv":
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(path, encoding=encoding)
            except UnicodeDecodeError:
                continue
        sys.exit(f"Could not decode {path.name}")
    return pd.read_excel(path, engine="openpyxl")


def to_number(value: object) -> float | None:
    """'1,200' / '₩1,200' / '(300)' -> float. Anything else -> None."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip().replace(",", "").replace("₩", "").replace("$", "").replace("€", "")
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def same_value(before: object, after: object, tolerance: float) -> bool:
    if (before is None or (isinstance(before, float) and pd.isna(before))) and \
       (after is None or (isinstance(after, float) and pd.isna(after))):
        return True
    if isinstance(before, str) and isinstance(after, str) and before.strip() == after.strip():
        return True
    b_num, a_num = to_number(before), to_number(after)
    if b_num is not None and a_num is not None:
        return abs(a_num - b_num) <= tolerance
    return str(before) == str(after)


def main() -> None:
    p = argparse.ArgumentParser(description="Diff two snapshots of the same export.")
    p.add_argument("--previous", default=CONFIG["previous"], help="older file (csv or xlsx)")
    p.add_argument("--current", default=CONFIG["current"], help="newer file")
    p.add_argument("--key", default=CONFIG["key"], help="comma-separated key columns")
    p.add_argument("--compare", default=CONFIG["compare"],
                   help="comma-separated columns to watch; default = all non-key columns")
    p.add_argument("--output", default=CONFIG["output_file"])
    p.add_argument("--log", default=CONFIG["log_file"])
    p.add_argument("--tolerance", type=float, default=CONFIG["numeric_tolerance"])
    args = p.parse_args()

    old = read_table(Path(args.previous).expanduser())
    new = read_table(Path(args.current).expanduser())
    if old.empty or new.empty:
        sys.exit("One of the files has no rows — refusing to report an empty diff.")

    keys = [k.strip() for k in args.key.split(",") if k.strip()] or [str(old.columns[0])]
    for key in keys:
        if key not in old.columns or key not in new.columns:
            side = "previous" if key not in old.columns else "current"
            sys.exit(f"Key column '{key}' is missing from the {side} file. Columns: {list(old.columns)}")

    watched = [c.strip() for c in args.compare.split(",") if c.strip()]
    if not watched:
        watched = [c for c in old.columns if c not in keys and c in new.columns]
    watched = [c for c in watched if c in old.columns and c in new.columns]
    if not watched:
        sys.exit("No overlapping columns to compare — check the column names.")

    dup_old = int(old.duplicated(subset=keys).sum())
    dup_new = int(new.duplicated(subset=keys).sum())
    old = old.drop_duplicates(subset=keys, keep="last")
    new = new.drop_duplicates(subset=keys, keep="last")

    merged = old[keys + watched].merge(
        new[keys + watched], on=keys, how="outer", suffixes=("_before", "_after"), indicator=True)

    new_rows = merged[merged["_merge"] == "right_only"].drop(columns=["_merge"])
    removed_rows = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])
    both = merged[merged["_merge"] == "both"]

    changes: list[dict] = []
    for _, row in both.iterrows():
        for column in watched:
            before, after = row[f"{column}_before"], row[f"{column}_after"]
            if same_value(before, after, args.tolerance):
                continue
            record = {key: row[key] for key in keys}
            b_num, a_num = to_number(before), to_number(after)
            record.update({
                "field": column,
                "before": before,
                "after": after,
                "difference": (a_num - b_num) if b_num is not None and a_num is not None else None,
            })
            changes.append(record)
    changed = pd.DataFrame(changes)

    # put the watched columns back in their original order for the new/removed sheets
    new_rows = new_rows[[c for c in new.columns if c in new_rows.columns]]
    removed_rows = removed_rows[[c for c in old.columns if c in removed_rows.columns]]

    summary = pd.DataFrame([
        {"metric": "previous file", "value": args.previous},
        {"metric": "current file", "value": args.current},
        {"metric": "rows in previous", "value": len(old)},
        {"metric": "rows in current", "value": len(new)},
        {"metric": "new rows", "value": len(new_rows)},
        {"metric": "removed rows", "value": len(removed_rows)},
        {"metric": "changed fields", "value": len(changed)},
        {"metric": "duplicate keys in previous", "value": dup_old},
        {"metric": "duplicate keys in current", "value": dup_new},
        {"metric": "run at", "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ])

    out = Path(args.output.replace("{date}", datetime.now().strftime("%Y%m%d_%H%M")))
    out.parent.mkdir(parents=True, exist_ok=True)
    head_fill = PatternFill("solid", fgColor="1F2A44")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        changed.to_excel(writer, sheet_name="changed", index=False)
        new_rows.to_excel(writer, sheet_name="new", index=False)
        removed_rows.to_excel(writer, sheet_name="removed", index=False)
        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.fill, cell.font = head_fill, Font(color="FFFFFF", bold=True, size=10)
                cell.alignment = Alignment(vertical="center")
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for idx, column in enumerate(ws.columns, start=1):
                width = max((len(str(c.value)) for c in column if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(idx)].width = min(max(10, width + 2), 40)

    log_path = Path(args.log)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame([{
        "run_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "previous": Path(args.previous).name,
        "current": Path(args.current).name,
        "new": len(new_rows), "removed": len(removed_rows), "changed": len(changed),
        "report": out.name,
    }]).to_csv(log_path, mode="a", header=not log_path.exists(), index=False, encoding="utf-8-sig")

    print(f"  {len(new_rows)} new · {len(removed_rows)} removed · {len(changed)} changed field(s)")
    print(f"  report: {out}")
    if len(changed):
        print(changed.head(5)[keys + ["field", "before", "after", "difference"]].to_string(index=False))

    if len(new_rows) or len(removed_rows) or len(changed):
        sys.exit(2)      # "something moved" — a scheduler can act on this


if __name__ == "__main__":
    main()
