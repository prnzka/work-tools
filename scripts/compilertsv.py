import os
import argparse
import pandas as pd
import csv
from tkinter import Tk, filedialog

# Columns to extract (final output format)
TARGET_COLUMNS = [
    "First Name",
    "Last Name",
    "Email",
    "Phone Number",
    "Company Name",
    "Website",
    "LinkedIn Profile",
    "Location",
    "Status",
    "Category",
    "ESP Type",
    "Current Sequence Number",
    "title",
    "emailType",
    "informalAddress",
    "informalIndustry",
]

# Deduplication keys
DEDUPLICATE_BY = ["email"]

# Header aliases (TSV → CSV mapping)
ALIASES = {
    "Phone Number": ["businessPhone"],
    "Website": ["businessUrl", "businessURL"],
    "Location": ["businessAddress"],
    "Company Name": ["businessName"],
}


def parse_args():
    p = argparse.ArgumentParser(description="Merge CSV/TSV files with header normalization")
    p.add_argument(
        "--output",
        "-o",
        help="Output path (default: ~/Downloads/merged_output.csv)",
        default=os.path.expanduser("~/Downloads/merged_output.csv"),
    )
    p.add_argument(
        "--files",
        "-f",
        nargs="*",
        help="List of CSV/TSV files (skip file dialog).",
    )
    return p.parse_args()


def _normalize(s: str) -> str:
    return "".join(ch for ch in s.lower().strip() if ch.isalnum())


def _detect_sep(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".tsv":
        return "\t"
    if ext == ".csv":
        return ","

    try:
        with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
            sample = f.read(8192)
        dialect = csv.Sniffer().sniff(sample, delimiters=[",", "\t", ";", "|"])
        return dialect.delimiter
    except Exception:
        return ","


def _build_expected_tokens():
    tokens = set()
    for t in TARGET_COLUMNS:
        tokens.add(_normalize(t))
    for alias_list in ALIASES.values():
        for a in alias_list:
            tokens.add(_normalize(a))
    return tokens


EXPECTED_TOKENS = _build_expected_tokens()


def _detect_header_row(path: str, sep: str, max_lines: int = 50) -> int:
    """
    Detect correct header row if file contains extra lines before actual header.
    """
    best_idx = 0
    best_score = -1

    try:
        with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
            for i in range(max_lines):
                line = f.readline()
                if not line:
                    break

                parts = [p.strip().strip('"').strip("'") for p in line.split(sep)]
                norms = [_normalize(p) for p in parts if p.strip()]

                score = sum(1 for n in norms if n in EXPECTED_TOKENS)

                if len(parts) >= 5:
                    score += 1

                if score > best_score:
                    best_score = score
                    best_idx = i

        return best_idx if best_score >= 2 else 0

    except Exception:
        return 0


def _read_table(path: str) -> pd.DataFrame:
    sep = _detect_sep(path)
    header_row = _detect_header_row(path, sep)

    try:
        return pd.read_csv(
            path,
            sep=sep,
            header=header_row,
            dtype=str,
            low_memory=False,
            on_bad_lines="skip",
            engine="c",
            encoding="utf-8-sig",
        )
    except Exception:
        return pd.read_csv(
            path,
            sep=sep,
            header=header_row,
            dtype=str,
            on_bad_lines="skip",
            engine="python",
            encoding="utf-8-sig",
        )


def _find_column(df: pd.DataFrame, target: str, col_map: dict) -> str | None:
    # Exact
    if target in df.columns:
        return target

    # Case-insensitive
    for c in df.columns:
        if c.lower() == target.lower():
            return c

    # Alias
    for alias in ALIASES.get(target, []):
        for c in df.columns:
            if c.lower() == alias.lower():
                return c
        alias_norm = _normalize(alias)
        if alias_norm in col_map:
            return col_map[alias_norm]

    # Normalized exact
    target_norm = _normalize(target)
    if target_norm in col_map:
        return col_map[target_norm]

    # Partial match
    for norm_col, original in col_map.items():
        if target_norm in norm_col or norm_col in target_norm:
            return original

    return None


def main():
    args = parse_args()

    output_path = os.path.expanduser(args.output)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    # File selection
    if args.files:
        file_paths = [os.path.expanduser(f) for f in args.files if os.path.exists(os.path.expanduser(f))]
    else:
        root = Tk()
        root.withdraw()
        file_paths = list(
            filedialog.askopenfilenames(
                title="Select CSV/TSV Files",
                filetypes=[
                    ("CSV or TSV files", "*.csv *.tsv"),
                    ("CSV files", "*.csv"),
                    ("TSV files", "*.tsv"),
                    ("All files", "*.*"),
                ],
            )
        )
        root.destroy()

    if not file_paths:
        print("No files selected.")
        return

    merged_data = []

    for file in file_paths:
        try:
            df = _read_table(file)
        except Exception as e:
            print(f"Error reading {file}: {e}")
            continue

        col_map = {_normalize(c): c for c in df.columns}
        extracted = pd.DataFrame()
        mapping = {}

        for target in TARGET_COLUMNS:
            chosen = _find_column(df, target, col_map)

            if chosen:
                extracted[target] = df[chosen].fillna("").astype(str).str.strip()
                mapping[target] = chosen
            else:
                extracted[target] = [""] * len(df)
                mapping[target] = None

        merged_data.append(extracted)

        mapped_info = ", ".join(f"{t}->{mapping[t] or 'MISSING'}" for t in TARGET_COLUMNS)
        print(f"Processed: {file} | {mapped_info}")

    if not merged_data:
        print("No data merged.")
        return

    final_df = pd.concat(merged_data, ignore_index=True)

    # Deduplicate (case-insensitive, trimmed)
    for d in DEDUPLICATE_BY:
        matches = [c for c in final_df.columns if c.lower() == d.lower()]
        if matches:
            for c in matches:
                final_df[c] = final_df[c].str.lower().str.strip()
            final_df.drop_duplicates(subset=matches, inplace=True)

    final_df = final_df[TARGET_COLUMNS]
    final_df.to_csv(output_path, index=False)

    print(f"\nDone! Saved to: {output_path}")


if __name__ == "__main__":
    main()