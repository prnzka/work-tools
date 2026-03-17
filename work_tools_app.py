#!/usr/bin/env python3
"""
Work Tools Desktop App (single-file launcher)

Run:
  python work_tools_app.py

Build EXE on Windows:
  pyinstaller --noconfirm --onefile --windowed --name WorkToolsDesktop work_tools_app.py
"""

import csv
import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

TARGET_COLUMNS = [
    "First Name", "Last Name", "Email", "Phone Number", "Company Name",
    "Website", "LinkedIn Profile", "Location", "Status", "Category",
    "ESP Type", "Current Sequence Number", "title", "emailType",
    "informalAddress", "informalIndustry", "businessRating", "businessReviewCount"
]

ALIASES = {
    "Phone Number": ["businessPhone"],
    "Website": ["businessUrl", "businessURL"],
    "Location": ["businessAddress"],
    "Company Name": ["businessName"],
}


def finder_output_cols():
    return [
        "First Name", "Last Name", "Email", "Phone Number", "Company Name",
        "Website", "LinkedIn Profile", "Location", "Status", "Category",
        "ESP Type", "Current Sequence Number", "title", "emailType",
        "informalAddress", "informalIndustry"
    ]


def normalize(s: str) -> str:
    return "".join(ch for ch in str(s).lower().strip() if ch.isalnum())


def find_column(columns, target):
    if target in columns:
        return target
    for col in columns:
        if col.lower() == target.lower():
            return col
    for alias in ALIASES.get(target, []):
        for col in columns:
            if col.lower() == alias.lower():
                return col
    target_norm = normalize(target)
    for col in columns:
        if normalize(col) == target_norm:
            return col
    for col in columns:
        n = normalize(col)
        if target_norm in n or n in target_norm:
            return col
    return None


def detect_sep(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".tsv":
        return "\t"
    if ext == ".csv":
        return ","
    with open(path, "r", encoding="utf-8-sig", errors="ignore") as fh:
        first = fh.readline()
    return "\t" if first.count("\t") > first.count(",") else ","


class WorkToolsDesktop(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Work Tools Desktop")
        self.geometry("1100x760")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.compiler_files = []
        self.base_file = None
        self.deduper_files = []
        self.finder_file = None

        self.tab_compiler = ttk.Frame(notebook)
        self.tab_words = ttk.Frame(notebook)
        self.tab_demails = ttk.Frame(notebook)
        self.tab_finder = ttk.Frame(notebook)
        self.tab_zip = ttk.Frame(notebook)

        notebook.add(self.tab_compiler, text="CSV/TSV Compiler")
        notebook.add(self.tab_words, text="Dedupe Words")
        notebook.add(self.tab_demails, text="Dedupe Emails")
        notebook.add(self.tab_finder, text="Location Finder")
        notebook.add(self.tab_zip, text="ZIP Extractor")

        self.build_compiler_tab()
        self.build_words_tab()
        self.build_demail_tab()
        self.build_finder_tab()
        self.build_zip_tab()

    def require_pandas(self):
        if pd is None:
            messagebox.showerror("Missing dependency", "pandas is required. Install with:\n\npython -m pip install pandas openpyxl")
            return False
        return True

    def build_compiler_tab(self):
        ttk.Button(self.tab_compiler, text="Select CSV/TSV Files", command=self.compiler_select_files).pack(pady=10)
        self.compiler_list = tk.Text(self.tab_compiler, height=6)
        self.compiler_list.pack(fill="x", padx=10)
        ttk.Button(self.tab_compiler, text="Run Compiler", command=self.run_compiler).pack(pady=8)
        self.compiler_status = ttk.Label(self.tab_compiler, text="")
        self.compiler_status.pack()

    def compiler_select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV/TSV", "*.csv *.tsv"), ("All files", "*.*")])
        self.compiler_files = list(files)
        self.compiler_list.delete("1.0", "end")
        self.compiler_list.insert("end", "\n".join(self.compiler_files))

    def run_compiler(self):
        if not self.require_pandas():
            return
        if not self.compiler_files:
            return messagebox.showwarning("Compiler", "Choose files first.")

        merged = []
        for file in self.compiler_files:
            try:
                sep = detect_sep(file)
                df = pd.read_csv(file, sep=sep, dtype=str, low_memory=False, on_bad_lines="skip").fillna("")
                cols = list(df.columns)
                col_map = {c: find_column(cols, c) for c in TARGET_COLUMNS}
                out = pd.DataFrame({t: df[col_map[t]].astype(str).str.strip() if col_map[t] else "" for t in TARGET_COLUMNS})
                merged.append(out)
            except Exception as e:
                messagebox.showerror("Compiler", f"Failed on {file}: {e}")
                return

        all_df = pd.concat(merged, ignore_index=True)
        all_df["Email"] = all_df["Email"].astype(str).str.strip().str.lower()
        before = len(all_df)
        deduped = all_df.drop_duplicates(subset=["Email"], keep="first")
        out = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="merged_output.csv")
        if out:
            deduped.to_csv(out, index=False, encoding="utf-8-sig")
            self.compiler_status.config(text=f"Saved {len(deduped):,} rows ({before-len(deduped):,} duplicates removed).")

    def build_words_tab(self):
        ttk.Label(self.tab_words, text="Paste one item per line:").pack(anchor="w", padx=10, pady=(10, 0))
        self.words_in = tk.Text(self.tab_words, height=18)
        self.words_in.pack(fill="both", expand=True, padx=10, pady=6)
        ttk.Button(self.tab_words, text="Deduplicate", command=self.run_words).pack(pady=8)
        self.words_out = tk.Text(self.tab_words, height=8)
        self.words_out.pack(fill="x", padx=10, pady=(0, 10))

    def run_words(self):
        items = [x.strip() for x in self.words_in.get("1.0", "end").splitlines() if x.strip()]
        seen, uniq = set(), []
        for item in items:
            if item not in seen:
                seen.add(item)
                uniq.append(item)
        self.words_out.delete("1.0", "end")
        self.words_out.insert("end", "; ".join(uniq))

    def build_demail_tab(self):
        row = ttk.Frame(self.tab_demails)
        row.pack(fill="x", pady=8)
        ttk.Button(row, text="Select Base CSV", command=self.set_base_file).pack(side="left", padx=8)
        ttk.Button(row, text="Select Deduper CSVs", command=self.set_dedupe_files).pack(side="left", padx=8)
        ttk.Button(row, text="Run Email Deduplication", command=self.run_demail).pack(side="left", padx=8)
        self.demail_info = ttk.Label(self.tab_demails, text="")
        self.demail_info.pack(anchor="w", padx=10)

    def set_base_file(self):
        f = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if f:
            self.base_file = f
            self.demail_info.config(text=f"Base: {Path(f).name}")

    def set_dedupe_files(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV", "*.csv")])
        self.deduper_files = list(files)
        if files:
            self.demail_info.config(text=f"Base: {Path(self.base_file).name if self.base_file else '-'} | Dedupers: {len(files)}")

    @staticmethod
    def find_email_col(cols):
        for c in cols:
            if c.strip().lower() == "email":
                return c
        for c in cols:
            if "email" in c.strip().lower():
                return c
        return None

    def run_demail(self):
        if not self.require_pandas():
            return
        if not self.base_file or not self.deduper_files:
            return messagebox.showwarning("Dedupe Emails", "Select base file and deduper files first.")

        base = pd.read_csv(self.base_file, dtype=str, low_memory=False, on_bad_lines="skip").fillna("")
        base_col = self.find_email_col(base.columns)
        if not base_col:
            return messagebox.showerror("Dedupe Emails", "No email column in base CSV.")

        dedupe_set = set()
        for file in self.deduper_files:
            df = pd.read_csv(file, dtype=str, low_memory=False, on_bad_lines="skip").fillna("")
            c = self.find_email_col(df.columns)
            if not c:
                continue
            dedupe_set.update(df[c].astype(str).str.strip().str.lower())

        initial = len(base)
        keep_mask = ~base[base_col].astype(str).str.strip().str.lower().isin(dedupe_set)
        out_df = base[keep_mask]
        out = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="deduped_output.csv")
        if out:
            out_df.to_csv(out, index=False, encoding="utf-8-sig")
            messagebox.showinfo("Dedupe Emails", f"Done. Removed {initial-len(out_df):,} rows.")

    def build_finder_tab(self):
        top = ttk.Frame(self.tab_finder)
        top.pack(fill="x", pady=8)
        ttk.Button(top, text="Select CSV / Excel", command=self.set_finder_file).pack(side="left", padx=8)
        self.finder_file_label = ttk.Label(top, text="No file selected")
        self.finder_file_label.pack(side="left", padx=6)

        form = ttk.Frame(self.tab_finder)
        form.pack(fill="x", padx=10, pady=6)
        ttk.Label(form, text="Location query (semicolon-separated):").grid(row=0, column=0, sticky="w")
        self.finder_loc = ttk.Entry(form, width=100)
        self.finder_loc.grid(row=1, column=0, sticky="we", pady=4)
        ttk.Label(form, text="Industry query (optional, semicolon-separated):").grid(row=2, column=0, sticky="w")
        self.finder_ind = ttk.Entry(form, width=100)
        self.finder_ind.grid(row=3, column=0, sticky="we", pady=4)

        ttk.Button(self.tab_finder, text="Search and Save Results", command=self.run_finder).pack(pady=8)
        self.finder_status = ttk.Label(self.tab_finder, text="")
        self.finder_status.pack(anchor="w", padx=10)

    def set_finder_file(self):
        f = filedialog.askopenfilename(filetypes=[("CSV / Excel", "*.csv *.xlsx *.xls")])
        if f:
            self.finder_file = f
            self.finder_file_label.config(text=Path(f).name)

    def run_finder(self):
        if not self.require_pandas():
            return
        if not self.finder_file:
            return messagebox.showwarning("Finder", "Select a file first.")

        loc_items = [x.strip() for x in self.finder_loc.get().split(";") if x.strip()]
        if not loc_items:
            return messagebox.showwarning("Finder", "Enter at least one location term.")

        ind_items = [x.strip() for x in self.finder_ind.get().split(";") if x.strip()]
        loc_pat = re.compile(r"\b(" + "|".join(re.escape(x) for x in loc_items) + r")\b", re.IGNORECASE)
        ind_pat = re.compile(r"\b(" + "|".join(re.escape(x) for x in ind_items) + r")\b", re.IGNORECASE) if ind_items else None

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="finder_results.csv")
        if not save_path:
            return

        total_scanned = 0
        total_matched = 0

        with open(save_path, "w", newline="", encoding="utf-8-sig") as out_f:
            writer = csv.DictWriter(out_f, fieldnames=finder_output_cols())
            writer.writeheader()

            suffix = Path(self.finder_file).suffix.lower()
            if suffix in [".xlsx", ".xls"]:
                df = pd.read_excel(self.finder_file, dtype=str).fillna("")
                total_scanned = len(df)
                for _, row in df.iterrows():
                    if not loc_pat.search(str(row.get("Location", ""))):
                        continue
                    if ind_pat and not ind_pat.search(str(row.get("informalIndustry", ""))):
                        continue
                    writer.writerow({c: row.get(c, "") for c in finder_output_cols()})
                    total_matched += 1
            else:
                for chunk in pd.read_csv(self.finder_file, dtype=str, chunksize=120_000, low_memory=False, on_bad_lines="skip"):
                    chunk = chunk.fillna("")
                    total_scanned += len(chunk)
                    if "Location" not in chunk.columns:
                        continue
                    loc_mask = chunk["Location"].astype(str).str.contains(loc_pat, na=False)
                    filt = chunk[loc_mask]
                    if ind_pat is not None and "informalIndustry" in filt.columns:
                        filt = filt[filt["informalIndustry"].astype(str).str.contains(ind_pat, na=False)]
                    elif ind_pat is not None:
                        filt = filt.iloc[0:0]
                    if not filt.empty:
                        total_matched += len(filt)
                        rows = [{c: row.get(c, "") for c in finder_output_cols()} for _, row in filt.iterrows()]
                        writer.writerows(rows)
                    self.update_idletasks()

        self.finder_status.config(text=f"Scanned {total_scanned:,} rows, found {total_matched:,} matches.")
        messagebox.showinfo("Finder", f"Done. Scanned {total_scanned:,} rows, found {total_matched:,} matches.")

    def build_zip_tab(self):
        ttk.Label(self.tab_zip, text="Extract ZIP codes from PDF (5-digit + ZIP+4):").pack(anchor="w", padx=10, pady=(10, 4))
        ttk.Button(self.tab_zip, text="Select PDF and Extract", command=self.run_zip_extract).pack(padx=10, pady=8, anchor="w")
        self.zip_out = tk.Text(self.tab_zip, height=14)
        self.zip_out.pack(fill="both", expand=True, padx=10, pady=8)

    def run_zip_extract(self):
        if PdfReader is None:
            return messagebox.showerror("ZIP Extractor", "Missing dependency: pypdf (install with: python -m pip install pypdf)")

        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return

        reader = PdfReader(path)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        zips = re.findall(r"\b\d{5}(?:-\d{4})?\b", text)
        unique = list(dict.fromkeys(zips))
        self.zip_out.delete("1.0", "end")
        self.zip_out.insert("end", "; ".join(unique))


if __name__ == "__main__":
    app = WorkToolsDesktop()
    app.mainloop()
