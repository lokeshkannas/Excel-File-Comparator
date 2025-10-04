\
from __future__ import annotations

import json
import os
import sys
from datetime import datetime
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from comparison_engine import compare_workbooks
from report_generator import generate_report


APP_TITLE = "Excel Workbook Comparator"
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")


def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"numeric_tolerance": 1e-9, "report_dir": "Documents/Excel_Comparisons", "theme": "dark"}


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.cfg = load_config()
        theme = self.cfg.get("theme", "dark")
        ctk.set_appearance_mode(theme)
        ctk.set_default_color_theme("dark-blue")

        self.title(APP_TITLE)
        self.geometry("860x640")
        self.minsize(840, 600)

        self.ssrs_path = tk.StringVar()
        self.pbi_path = tk.StringVar()
        default_name = f"Comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        self.report_name = tk.StringVar(value=default_name)

        # Top frame (file pickers)
        top = ctk.CTkFrame(self, corner_radius=16)
        top.pack(fill="x", padx=16, pady=(16, 8))

        self._build_picker(top, "SSRS Excel", self.ssrs_path, "#3b82f6")
        self._build_picker(top, "Power BI Excel", self.pbi_path, "#06b6d4")

        # Report name frame
        name_frame = ctk.CTkFrame(self, corner_radius=16)
        name_frame.pack(fill="x", padx=16, pady=8)
        ctk.CTkLabel(name_frame, text="Report file name:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=12, pady=12)
        name_entry = ctk.CTkEntry(name_frame, textvariable=self.report_name, width=420)
        name_entry.pack(side="left", padx=8, pady=12)
        ctk.CTkLabel(name_frame, text="(e.g., Sales_Sept_Check.xlsx)").pack(side="left", padx=8)

        # Buttons
        btns = ctk.CTkFrame(self, corner_radius=16)
        btns.pack(fill="x", padx=16, pady=8)
        self.compare_btn = ctk.CTkButton(btns, text="▶ Compare Files", command=self.on_compare, height=40)
        self.compare_btn.pack(side="left", padx=8, pady=8)
        self.open_btn = ctk.CTkButton(btns, text="📂 Open Report Folder", command=self.open_report_folder, height=40)
        self.open_btn.pack(side="left", padx=8, pady=8)

        # Results text area
        self.result_box = ctk.CTkTextbox(self, height=360, corner_radius=16)
        self.result_box.pack(fill="both", expand=True, padx=16, pady=(8,16))
        self._write_result("Ready. Select two Excel files and click 'Compare Files'.\n")

    def _build_picker(self, parent, label, var, accent):
        frame = ctk.CTkFrame(parent, corner_radius=16)
        frame.pack(fill="x", padx=12, pady=12)

        ctk.CTkLabel(frame, text=label, font=("Segoe UI", 14, "bold")).pack(side="left", padx=12, pady=12)
        entry = ctk.CTkEntry(frame, textvariable=var, width=520)
        entry.pack(side="left", padx=8, pady=12)
        btn = ctk.CTkButton(frame, text="Browse", command=lambda: self._browse(var))
        btn.pack(side="left", padx=8, pady=12)

    def _browse(self, var):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"), ("All files", "*.*")]
        )
        if path:
            var.set(path)

    def _write_result(self, text: str):
        self.result_box.insert("end", text)
        self.result_box.see("end")

    def open_report_folder(self):
        report_dir = os.path.join(os.path.expanduser("~"), self.cfg.get("report_dir", "Documents/Excel_Comparisons"))
        os.makedirs(report_dir, exist_ok=True)
        # Open folder in file explorer
        if sys.platform.startswith("win"):
            os.startfile(report_dir)  # type: ignore
        elif sys.platform == "darwin":
            os.system(f'open "{report_dir}"')
        else:
            os.system(f'xdg-open "{report_dir}"')

    def on_compare(self):
        ssrs = self.ssrs_path.get().strip()
        pbi = self.pbi_path.get().strip()
        if not ssrs or not pbi:
            messagebox.showwarning(APP_TITLE, "Please select both SSRS and Power BI Excel files.")
            return

        # Validate report name
        report_name = self.report_name.get().strip() or f"Comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        if not report_name.lower().endswith(".xlsx"):
            report_name += ".xlsx"

        report_dir = os.path.join(os.path.expanduser("~"), self.cfg.get("report_dir", "Documents/Excel_Comparisons"))
        os.makedirs(report_dir, exist_ok=True)
        output_path = os.path.join(report_dir, report_name)

        self.result_box.delete("1.0", "end")
        self._write_result("Comparing workbooks...\n\n")

        try:
            tol = float(self.cfg.get("numeric_tolerance", 1e-9))
            result = compare_workbooks(ssrs, pbi, numeric_tolerance=tol)

            # Generate report
            out = generate_report(output_path, result)

            summary = result.summary
            if summary["all_matched"]:
                self._write_result("✅ All checks passed. No mismatches found.\n")
            else:
                self._write_result("⚠️ Mismatches found.\n")

            self._write_result("\n--- Summary ---\n")
            self._write_result(f"Structure issues: {summary['structure_issue_count']}\n")
            self._write_result(f"Dtype issues:     {summary['dtype_issue_count']}\n")
            self._write_result(f"Value mismatches: {summary['value_mismatch_count']}\n")

            # Show a concise list of first N mismatches
            max_show = 200
            shown = 0

            if result.structure_issues:
                self._write_result("\n[Structure Issues]\n")
                for s in result.structure_issues[:max_show]:
                    self._write_result(f" - [{s.sheet}] {s.issue}: {s.detail}\n")
                    shown += 1

            if result.dtype_issues and shown < max_show:
                self._write_result("\n[Data Type Issues]\n")
                for d in result.dtype_issues[: max_show - shown]:
                    self._write_result(f" - [{d.sheet}] {d.column}: SSRS={d.ssrs_dtype} vs PowerBI={d.powerbi_dtype}\n")
                    shown += 1

            if result.value_mismatches and shown < max_show:
                self._write_result("\n[Value Mismatches]\n")
                for v in result.value_mismatches[: max_show - shown]:
                    self._write_result(f" - [{v.sheet}] Row {v.row}, Col '{v.column}': SSRS='{v.ssrs_value}' vs PowerBI='{v.powerbi_value}'\n")
                    shown += 1

            self._write_result(f"\n📄 Report saved at:\n{out}\n")
            self._write_result("Tip: Click '📂 Open Report Folder' to view the file.\n")

        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Comparison failed:\n{e}")
            self._write_result(f"\n❌ Error: {e}\n")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
