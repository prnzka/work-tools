#!/usr/bin/env python3
"""
Simple CSV email deduper.

Usage: run the script and select the base CSV, then select one or more deduper CSVs.
The script finds an "Email" column (case-insensitive / substring match), normalizes emails,
removes rows in the base file whose emails appear in any deduper file, and saves the result.
"""
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def read_csv_guess(path):
    # Try the fast C engine first with low_memory disabled (recommended for large files).
    for enc in ("utf-8", "latin1"):
        try:
            return pd.read_csv(path, encoding=enc, low_memory=False)
        except Exception:
            continue
    # Fall back to the python engine (which doesn't support low_memory).
    for enc in ("utf-8", "latin1"):
        try:
            return pd.read_csv(path, engine="python", encoding=enc)
        except Exception:
            continue
    # Final attempt letting pandas choose defaults.
    return pd.read_csv(path)


def find_email_column(df):
    cols = list(df.columns)
    # exact match ignoring case
    for c in cols:
        if c.strip().lower() == "email":
            return c
    # substring match
    for c in cols:
        if "email" in c.strip().lower():
            return c
    return None


def normalize_email_series(s):
    return s.astype(str).str.strip().str.lower().replace({"nan": None}).dropna()


def main():
    root = tk.Tk()
    root.withdraw()

    base_path = filedialog.askopenfilename(title="Select base CSV file (file 1)", filetypes=[("CSV files", "*.csv"), ("All files", "*")])
    if not base_path:
        print("No base file selected. Exiting.")
        return

    deduper_paths = filedialog.askopenfilenames(title="Select deduper CSV file(s)", filetypes=[("CSV files", "*.csv"), ("All files", "*")])
    if not deduper_paths:
        print("No deduper files selected. Exiting.")
        return

    try:
        base_df = read_csv_guess(base_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read base CSV: {e}")
        return

    base_email_col = find_email_column(base_df)
    if base_email_col is None:
        messagebox.showerror("Error", "Could not find an 'Email' column in the base file.")
        return

    base_emails = normalize_email_series(base_df[base_email_col])

    deduper_emails = set()
    for p in deduper_paths:
        try:
            df = read_csv_guess(p)
        except Exception as e:
            messagebox.showwarning("Warning", f"Failed to read {p}: {e}")
            continue
        col = find_email_column(df)
        if col is None:
            messagebox.showwarning("Warning", f"Could not find 'Email' column in {p}; skipping.")
            continue
        deduper_emails.update(normalize_email_series(df[col]).tolist())

    # filter base_df rows whose normalized email appears in deduper_emails
    def email_in_deduper(val):
        try:
            v = str(val).strip().lower()
        except Exception:
            return False
        if v == "nan" or v == "":
            return False
        return v in deduper_emails

    initial_count = len(base_df)
    mask = base_df[base_email_col].apply(lambda v: not email_in_deduper(v))
    result_df = base_df[mask]
    removed = initial_count - len(result_df)

    # Collect removed emails (original-looking, stripped)
    removed_series = base_df.loc[~mask, base_email_col].astype(str).str.strip()
    removed_emails = (
        removed_series.replace({"nan": ""})
        .loc[lambda s: s.str.len() > 0]
        .dropna()
        .astype(str)
        .tolist()
    )

    save_path = filedialog.asksaveasfilename(title="Save deduped CSV as", defaultextension=".csv", filetypes=[("CSV files", "*.csv" )])
    if not save_path:
        print("No save location chosen. Exiting without saving.")
        return

    try:
        result_df.to_csv(save_path, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save file: {e}")
        return

    messagebox.showinfo("Done", f"Saved deduped file to:\n{save_path}\n\nRows removed: {removed}")
    print(f"Saved deduped file to: {save_path}")
    if removed_emails:
        print("\nRemoved emails:")
        for e in removed_emails:
            print(e)
    else:
        print("\nNo emails were removed.")


if __name__ == "__main__":
    main()
