#!/usr/bin/env python3
"""Convert Rayyan RIS export to Excel spreadsheet.

Usage:
    python process_ris.py articles.ris output.xlsx

The script parses specific fields and expands reviewer decisions and
exclusion reasons as described by the user.
"""

import argparse
import json
import re
from collections import defaultdict
from typing import List, Tuple

import pandas as pd


def parse_ris_file(path: str) -> List[dict]:
    """Parse a RIS file into a list of records.

    Each record is represented as a dictionary mapping RIS tags to lists of
    values. Multiple occurrences of the same tag (e.g., AU) are preserved.
    """
    records: List[dict] = []
    record: defaultdict = defaultdict(list)

    with open(path, encoding="utf-8") as fh:
        for raw_line in fh:
            line = raw_line.rstrip("\n")
            if not line:
                continue
            if line.startswith("ER  -"):
                if record:
                    records.append(record)
                record = defaultdict(list)
                continue
            tag = line[:2]
            if line[2:6].strip() == "-":
                value = line[6:].strip()
                record[tag].append(value)
    return records


def parse_n1(n1_values: List[str]) -> Tuple[str, str, str, List[str]]:
    """Parse the N1 field into reviewer decisions and exclusion reasons."""
    text = " | ".join(n1_values)
    rita = jules = ""
    agreement = ""
    reasons: List[str] = []

    match = re.search(r"RAYYAN-INCLUSION:\s*({.*?})", text)
    if match:
        decisions_str = match.group(1).replace("=>", ":")
        try:
            decisions = json.loads(decisions_str)
        except json.JSONDecodeError:
            decisions = {}
        rita = decisions.get("Rita", "")
        jules = decisions.get("Jules", "")
        if rita and jules:
            agreement = "Yes" if rita == jules else "No"

    match = re.search(r"RAYYAN-EXCLUSION-REASONS:\s*([^|]*)", text)
    if match:
        reasons = [r.strip() for r in match.group(1).split(",") if r.strip()]

    return rita, jules, agreement, reasons


def build_dataframe(records: List[dict]) -> pd.DataFrame:
    """Convert parsed RIS records into a pandas DataFrame."""
    all_reasons = set()
    parsed_rows = []

    for rec in records:
        rita, jules, agreement, reasons = parse_n1(rec.get("N1", []))
        all_reasons.update(reasons)
        authors = rec.get("AU", [])
        first_author = authors[0] if authors else ""
        other_authors = "; ".join(authors[1:])
        parsed_rows.append(
            {
                "TI": " ".join(rec.get("TI", [])),
                "T2": " ".join(rec.get("T2", [])),
                "Y2": " ".join(rec.get("Y2", [])),
                "Y3": " ".join(rec.get("Y3", [])),
                "First Author": first_author,
                "Other Authors": other_authors,
                "AB": " ".join(rec.get("AB", [])),
                "DO": " ".join(rec.get("DO", [])),
                "AN": " ".join(rec.get("AN", [])),
                "Rita Decision": rita,
                "Jules Decision": jules,
                "Agreement": agreement,
                "_reasons": reasons,
            }
        )

    df = pd.DataFrame(parsed_rows)
    for reason in sorted(all_reasons):
        df[reason] = df["_reasons"].apply(lambda r: "Yes" if reason in r else "")
    return df.drop(columns="_reasons")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert Rayyan RIS export to Excel")
    parser.add_argument("input_ris", help="Path to the articles.ris file")
    parser.add_argument("output_xlsx", help="Output Excel file path")
    args = parser.parse_args()

    records = parse_ris_file(args.input_ris)
    df = build_dataframe(records)
    df.to_excel(args.output_xlsx, index=False)


if __name__ == "__main__":
    main()
