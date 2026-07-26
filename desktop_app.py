"""
desktop_app.py
A simple double-click desktop tool for non-technical daily use:

    1. Click "Select Excel File" -> pick your company's .xlsx/.xls/.csv
    2. Click "Generate & Send Report" -> PPT is built and (optionally) emailed
    3. Done -- output folder opens automatically

No command line needed once set up. Run once with:  python desktop_app.py
Or double-click "Generate Report.bat" on Windows.

Reads SMTP credentials from a local .env file (same as the scheduled pipeline) --
see .env.example. Only DATA_SOURCE / CSV_PATH / EXCEL_PATH in .env are ignored
here since you're picking the file directly from the dialog each time.
"""

import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv
from src import ingest, process, charts, report_builder
from src.email_sender import send_report

load_dotenv()

APP_BG = "#F7F8FA"
NAVY = "#1B2A4A"
TEAL = "#2E9E8F"
CORAL = "#E4622C"


class ReportApp:
    def __init__(self, root):
        self.root = root
        self.selected_file = None

        root.title("Business Report Generator")
        root.geometry("520x360")
        root.configure(bg=APP_BG)
        root.resizable(False, False)

        tk.Label(root, text="Business Report Generator", font=("Segoe UI", 16, "bold"),
                 bg=APP_BG, fg=NAVY).pack(pady=(24, 4))
        tk.Label(root, text="Select your Excel file, then generate + email the report.",
                 font=("Segoe UI", 10), bg=APP_BG, fg="#5A6472").pack(pady=(0, 20))

        self.file_label = tk.Label(root, text="No file selected", font=("Segoe UI", 10),
                                    bg="white", fg="#5A6472", relief="solid", bd=1,
                                    width=52, height=2, anchor="w", padx=10)
        self.file_label.pack(pady=(0, 12))

        btn_row = tk.Frame(root, bg=APP_BG)
        btn_row.pack(pady=(0, 10))
        tk.Button(btn_row, text="📁 Select Excel File", font=("Segoe UI", 10, "bold"),
                  bg="white", fg=NAVY, relief="solid", bd=1, padx=12, pady=6,
                  command=self.select_file, cursor="hand2").pack(side="left", padx=4)
        tk.Button(btn_row, text="📂 Use Latest from 'incoming/'", font=("Segoe UI", 10, "bold"),
                  bg="white", fg=NAVY, relief="solid", bd=1, padx=12, pady=6,
                  command=self.select_latest_from_incoming, cursor="hand2").pack(side="left", padx=4)

        self.send_email_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="Also email the report (uses .env SMTP settings)",
                        variable=self.send_email_var, bg=APP_BG, font=("Segoe UI", 9)).pack(pady=(0, 16))

        self.generate_btn = tk.Button(root, text="▶  Generate & Send Report",
                                       font=("Segoe UI", 11, "bold"), bg=TEAL, fg="white",
                                       relief="flat", padx=16, pady=10, cursor="hand2",
                                       command=self.run_pipeline_threaded, state="disabled")
        self.generate_btn.pack(pady=(0, 16))

        self.status_label = tk.Label(root, text="", font=("Segoe UI", 9), bg=APP_BG, fg="#5A6472",
                                      wraplength=460, justify="left")
        self.status_label.pack(pady=(0, 6))

        self.progress = ttk.Progressbar(root, mode="indeterminate", length=460)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select company data file",
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if path:
            self.selected_file = path
            self.file_label.config(text=os.path.basename(path))
            self.generate_btn.config(state="normal")

    def select_latest_from_incoming(self):
        incoming_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "incoming")
        os.makedirs(incoming_dir, exist_ok=True)
        valid_ext = (".xlsx", ".xls", ".csv")
        files = [os.path.join(incoming_dir, f) for f in os.listdir(incoming_dir)
                 if f.lower().endswith(valid_ext) and not f.startswith("~$")]
        if not files:
            messagebox.showwarning("No files found",
                                    "The 'incoming' folder is empty. Drop your Excel/CSV file there first.")
            return
        latest = max(files, key=os.path.getmtime)
        self.selected_file = latest
        self.file_label.config(text=f"(latest) {os.path.basename(latest)}")
        self.generate_btn.config(state="normal")

    def set_status(self, text, color="#5A6472"):
        self.status_label.config(text=text, fg=color)
        self.root.update_idletasks()

    def run_pipeline_threaded(self):
        self.generate_btn.config(state="disabled")
        self.progress.pack(pady=(0, 10))
        self.progress.start(10)
        threading.Thread(target=self.run_pipeline, daemon=True).start()

    def run_pipeline(self):
        try:
            self.set_status("Reading file...")
            raw_df = ingest.from_any_file(self.selected_file)

            self.set_status("Cleaning and calculating KPIs...")
            df = process.clean_data(raw_df)
            kpis = process.compute_kpis(df)

            self.set_status("Building charts...")
            chart_paths = {}
            region_df = process.region_summary(df)
            if not region_df.empty:
                chart_paths["region"] = charts.region_bar_chart(region_df)
            trend_df = process.trend_summary(df)
            if not trend_df.empty:
                chart_paths["trend"] = charts.trend_line_chart(trend_df)
            sp_df = process.salesperson_summary(df)
            if not sp_df.empty:
                chart_paths["leaderboard"] = charts.salesperson_leaderboard_chart(sp_df)

            self.set_status("Creating PowerPoint report...")
            output_path = report_builder.build_report(kpis, chart_paths)

            if self.send_email_var.get():
                self.set_status("Sending email...")
                send_report(output_path, subject="Business Report")

            self.progress.stop()
            self.progress.pack_forget()
            self.set_status(f"✅ Done! Report saved to {output_path}", color=TEAL)
            self.generate_btn.config(state="normal")

            messagebox.showinfo("Success", "Report generated" +
                                 (" and emailed!" if self.send_email_var.get() else "!"))
            self.open_output_folder(output_path)

        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.set_status(f"❌ Error: {e}", color=CORAL)
            self.generate_btn.config(state="normal")
            traceback.print_exc()
            messagebox.showerror("Error", str(e))

    def open_output_folder(self, output_path):
        folder = os.path.dirname(os.path.abspath(output_path))
        try:
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception:
            pass  # non-fatal if the folder can't be auto-opened


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
