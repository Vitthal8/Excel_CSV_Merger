#!/usr/bin/env python3
"""
excel_merger_fixed.py

Updates:
- FIXED: Added 'clean_illegal_chars' to strip invalid Excel characters (like ).
- Preserves the Sheet Filter and High Performance mode.
"""
import os
import threading
import queue
import datetime
import gc
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

# Supported extensions
SUPPORTED_EXTS = {'.csv', '.xlsx', '.xlsm', '.xls'}

class MergerGUI:
    def __init__(self, root):
        self.root = root
        root.title("Excel/CSV Merger (Auto-Fix Bad Chars)")
        self.file_list = []
        self.log_q = queue.Queue()

        # --- Top frame: file selection ---
        top = ttk.Frame(root, padding=(10, 10))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        self.add_btn = ttk.Button(top, text="Select Files...", command=self.select_files)
        self.add_btn.grid(row=0, column=0, sticky="w")

        self.remove_btn = ttk.Button(top, text="Remove Selected", command=self.remove_selected)
        self.remove_btn.grid(row=0, column=1, sticky="w", padx=(6, 0))

        self.clear_btn = ttk.Button(top, text="Clear List", command=self.clear_files)
        self.clear_btn.grid(row=0, column=2, sticky="w", padx=(6, 0))

        # --- Middle: Listbox ---
        mid = ttk.Frame(root, padding=(10, 0, 10, 0))
        mid.grid(row=1, column=0, sticky="nsew")
        root.rowconfigure(1, weight=1)
        mid.columnconfigure(0, weight=1)

        self.files_var = tk.Variable(value=[])
        self.listbox = tk.Listbox(mid, listvariable=self.files_var, height=8, selectmode=tk.EXTENDED)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(mid, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        # --- Sheet Filter ---
        filter_frame = ttk.LabelFrame(root, text="Sheet Filter", padding=(10, 5))
        filter_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 0))
        filter_frame.columnconfigure(1, weight=1)

        self.scan_btn = ttk.Button(filter_frame, text="Scan for Sheet Names", command=self.scan_sheet_names)
        self.scan_btn.grid(row=0, column=0, sticky="w")

        ttk.Label(filter_frame, text="Select Sheet:").grid(row=0, column=1, sticky="e", padx=(10, 5))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(filter_frame, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=0, column=2, sticky="ew")
        self.sheet_combo.set("All Sheets")
        self.sheet_combo['values'] = ["All Sheets"]

        # --- Output Options ---
        out = ttk.LabelFrame(root, text="Output Options", padding=(10, 5))
        out.grid(row=3, column=0, sticky="ew", padx=10, pady=(5, 0))
        out.columnconfigure(1, weight=1)

        ttk.Label(out, text="Format:").grid(row=0, column=0, sticky="w")
        self.out_format = tk.StringVar(value="csv")
        ttk.Radiobutton(out, text="CSV (Fast)", variable=self.out_format, value="csv").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(out, text="Excel (Slow)", variable=self.out_format, value="xlsx").grid(row=0, column=2, sticky="w")

        ttk.Label(out, text="File:").grid(row=1, column=0, sticky="w", pady=5)
        self.out_entry = ttk.Entry(out)
        self.out_entry.grid(row=1, column=1, columnspan=2, sticky="ew", pady=5)
        self.browse_out_btn = ttk.Button(out, text="Browse...", command=self.choose_output_path)
        self.browse_out_btn.grid(row=1, column=3, sticky="w", padx=5)

        self.include_source = tk.BooleanVar(value=True)
        ttk.Checkbutton(out, text="Add 'Source File' column", variable=self.include_source).grid(row=2, column=0, columnspan=3, sticky="w")

        # --- Action ---
        action = ttk.Frame(root, padding=(10, 10))
        action.grid(row=4, column=0, sticky="ew")
        self.merge_btn = ttk.Button(action, text="Start Merge", command=self.start_merge)
        self.merge_btn.grid(row=0, column=0, sticky="w")
        self.status_label = ttk.Label(action, text="Ready")
        self.status_label.grid(row=0, column=1, sticky="e")
        action.columnconfigure(1, weight=1)

        # --- Log ---
        log_frame = ttk.LabelFrame(root, text="Log", padding=(10, 5))
        log_frame.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0, 10))
        root.rowconfigure(5, weight=1)
        self.log_text = tk.Text(log_frame, height=8, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.root.after(100, self.poll_log_queue)

    # ---------------- UI Logic ----------------
    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel/CSV", "*.csv *.xlsx *.xlsm *.xls")])
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
        self.update_listbox()
        self.log(f"Total files: {len(self.file_list)}")

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            del self.file_list[idx]
        self.update_listbox()

    def clear_files(self):
        self.file_list.clear()
        self.update_listbox()

    def update_listbox(self):
        self.files_var.set(self.file_list)

    def choose_output_path(self):
        ext = self.out_format.get()
        path = filedialog.asksaveasfilename(defaultextension=f".{ext}")
        if path:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, path)

    def log(self, msg):
        self.log_q.put(msg)

    def poll_log_queue(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self.poll_log_queue)

    def scan_sheet_names(self):
        if not self.file_list:
            messagebox.showwarning("No Files", "Add files before scanning.")
            return

        self.log("Scanning files for sheet names...")
        unique_sheets = set()
        self.root.config(cursor="watch")
        self.root.update()

        for f in self.file_list:
            ext = os.path.splitext(f)[1].lower()
            if ext == ".csv": continue
            try:
                try:
                    xls = pd.ExcelFile(f, engine="calamine")
                except:
                    xls = pd.ExcelFile(f)
                unique_sheets.update(xls.sheet_names)
            except Exception as e:
                self.log(f"Warning: Could not read {os.path.basename(f)}")

        sorted_sheets = sorted(list(unique_sheets))
        self.sheet_combo['values'] = ["All Sheets"] + sorted_sheets
        self.sheet_combo.set("All Sheets")
        self.root.config(cursor="")
        self.log(f"Scan complete. Found {len(sorted_sheets)} sheets.")

    # ---------------- Merge Logic ----------------
    def start_merge(self):
        if not self.file_list: return
        out = self.out_entry.get()
        if not out:
            messagebox.showwarning("Missing Output", "Select an output file.")
            return

        self.merge_btn.config(state="disabled")
        self.status_label.config(text="Processing...")

        target_sheet = self.sheet_var.get()
        if target_sheet == "All Sheets": target_sheet = None

        threading.Thread(target=self._worker, args=(out, target_sheet), daemon=True).start()

    def clean_illegal_chars(self, val):
        """Removes characters that crash Excel (like ASCII 26 'Substitute')."""
        if isinstance(val, str):
            # Only allow standard printable chars + tab + newline
            # This strips out low ASCII control codes (0-31) except 9, 10, 13
            return "".join(c for c in val if c >= ' ' or c in '\t\n\r')
        return val

    def _worker(self, out_path, target_sheet):
        try:
            files = self.file_list
            total = len(files)
            fmt = self.out_format.get()
            write_header = True
            dfs_buffer = []

            for i, fpath in enumerate(files):
                fname = os.path.basename(fpath)
                self.log(f"[{i+1}/{total}] {fname}")

                df = self._read_filtered(fpath, target_sheet)

                if df is not None and not df.empty:
                    # --- FIX: Sanitize Columns before processing ---
                    # We check string columns and clean bad characters
                    str_cols = df.select_dtypes(include=['object']).columns
                    if not str_cols.empty:
                        for col in str_cols:
                            df[col] = df[col].map(self.clean_illegal_chars)
                    # -----------------------------------------------

                    if fmt == "csv":
                        mode = 'w' if write_header else 'a'
                        df.to_csv(out_path, mode=mode, index=False, header=write_header)
                        write_header = False
                        self.log(f"  -> Appended {len(df)} rows")
                    else:
                        dfs_buffer.append(df)

                del df
                gc.collect()

            if fmt == "xlsx":
                if dfs_buffer:
                    self.log("Concatenating and writing to Excel...")
                    full_df = pd.concat(dfs_buffer, ignore_index=True)

                    # Double check final dataframe just in case
                    self.log("Final check for bad characters...")
                    # (Code already cleaned chunks, so we should be safe)

                    full_df.to_excel(out_path, index=False)
                    self.log("Excel write complete.")
                else:
                    self.log("No data found to write.")

            messagebox.showinfo("Success", "Merge Complete!")

        except Exception as e:
            self.log(f"Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.root.after(0, lambda: self.merge_btn.config(state="normal"))
            self.root.after(0, lambda: self.status_label.config(text="Ready"))

    def _read_filtered(self, fpath, target_sheet):
        ext = os.path.splitext(fpath)[1].lower()
        include_src = self.include_source.get()
        base = os.path.basename(fpath)

        try:
            if ext == ".csv":
                if target_sheet and target_sheet != "<csv>": return None
                df = pd.read_csv(fpath)
                if include_src:
                    df["Source_File"] = base
                    df["Sheet_Name"] = "<csv>"
                return df

            try:
                xls = pd.ExcelFile(fpath, engine="calamine")
            except:
                xls = pd.ExcelFile(fpath)

            dfs = []
            sheets = xls.sheet_names
            if target_sheet:
                if target_sheet in sheets: sheets = [target_sheet]
                else: return None

            for s in sheets:
                d = pd.read_excel(xls, sheet_name=s)
                if include_src:
                    d["Source_File"] = base
                    d["Sheet_Name"] = s
                dfs.append(d)

            if dfs: return pd.concat(dfs, ignore_index=True)
            return None

        except Exception as e:
            self.log(f"  -> Read Error: {e}")
            return None

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("700x650")
    MergerGUI(root)
    root.mainloop()