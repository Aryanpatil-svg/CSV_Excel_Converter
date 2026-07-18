import os
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import converter
from logger import logger

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DataEngine Studio - CSV ⇄ Excel Pro")
        self.root.geometry("850x680")
        self.root.configure(bg="#0f172a")

        self.bg_color = "#0f172a"
        self.secondary_bg = "#1e293b"
        self.accent_color = "#0ea5e9"
        self.text_muted = "#94a3b8"
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.setup_styles()

        self.input_file_path = tk.StringVar()
        self.output_directory = tk.StringVar(value=os.path.abspath("output"))
        self.conversion_type = tk.StringVar(value="CSV to Excel")
        
        self.var_dup = tk.BooleanVar(value=False)
        self.var_missing = tk.StringVar(value="Ignore")
        self.var_date = tk.BooleanVar(value=False)

        self.create_widgets()

    def setup_styles(self):
        self.style.configure(".", background=self.bg_color, foreground="white")
        self.style.configure("TLabel", background=self.bg_color, foreground="white", font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground=self.accent_color)
        self.style.configure("Card.TFrame", background=self.secondary_bg, relief="flat")
        self.style.configure("TEntry", fieldbackground=self.secondary_bg, foreground="white", borderwidth=0)
        self.style.configure("Muted.TLabel", background=self.secondary_bg, foreground=self.text_muted, font=("Segoe UI", 10))
        self.style.configure("Cyber.Horizontal.TProgressbar", thickness=8, troughcolor=self.secondary_bg, color=self.accent_color)
        
        self.style.configure("Action.TButton", font=("Segoe UI", 10, "bold"), background=self.accent_color, foreground="white")
        self.style.map("Action.TButton", background=[('active', '#0284c7')])
        self.style.configure("TButton", font=("Segoe UI", 10), background=self.secondary_bg, foreground="white")
        self.style.map("TButton", background=[('active', '#334155')])
        
        self.style.configure("Treeview", background=self.secondary_bg, fieldbackground=self.secondary_bg, foreground="white", rowheight=25)
        self.style.configure("Treeview.Heading", background="#334155", foreground="white", font=("Segoe UI", 9, "bold"))

    def create_widgets(self):
        main_layout = tk.Frame(self.root, bg=self.bg_color, padx=20, pady=20)
        main_layout.pack(fill="both", expand=True)

        top_bar = tk.Frame(main_layout, bg=self.bg_color)
        top_bar.pack(fill="x", pady=(0, 15))
        ttk.Label(top_bar, text="DATAENGINE STUDIO PRO", style="Header.TLabel").pack(side="left")
        
        mode_combo = ttk.Combobox(top_bar, textvariable=self.conversion_type, values=["CSV to Excel", "Excel to CSV"], state="readonly", width=14)
        mode_combo.pack(side="right", padx=5)
        mode_combo.bind("<<ComboboxSelected>>", lambda e: self.input_file_path.set(""))

        io_card = ttk.Frame(main_layout, style="Card.TFrame", padding=15)
        io_card.pack(fill="x", pady=5)

        ttk.Label(io_card, text="Source File Location:", background=self.secondary_bg).grid(row=0, column=0, sticky="w")
        ttk.Entry(io_card, textvariable=self.input_file_path, width=70).grid(row=1, column=0, ipady=4, padx=(0, 10), pady=(2, 10))
        ttk.Button(io_card, text="Browse File", command=self.handle_file_browsing).grid(row=1, column=1, pady=(2, 10))

        ttk.Label(io_card, text="Target Output Folder:", background=self.secondary_bg).grid(row=2, column=0, sticky="w")
        ttk.Entry(io_card, textvariable=self.output_directory, width=70).grid(row=3, column=0, ipady=4, padx=(0, 10))
        ttk.Button(io_card, text="Select Destination", command=self.handle_dest_browsing).grid(row=3, column=1)

        middle_grid = tk.Frame(main_layout, bg=self.bg_color)
        middle_grid.pack(fill="x", pady=15)

        utilities_panel = ttk.Frame(middle_grid, style="Card.TFrame", padding=15)
        utilities_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        ttk.Label(utilities_panel, text="Data Optimization Toolkit", font=("Segoe UI", 11, "bold"), background=self.secondary_bg).pack(anchor="w", pady=(0, 10))
        tk.Checkbutton(utilities_panel, text="Remove Duplicate Rows", variable=self.var_dup, bg=self.secondary_bg, fg="white", selectcolor=self.secondary_bg, activebackground=self.secondary_bg, activeforeground="white").pack(anchor="w", pady=2)
        tk.Checkbutton(utilities_panel, text="Auto-Parse Dates & Timezones", variable=self.var_date, bg=self.secondary_bg, fg="white", selectcolor=self.secondary_bg, activebackground=self.secondary_bg, activeforeground="white").pack(anchor="w", pady=2)
        
        missing_frame = tk.Frame(utilities_panel, bg=self.secondary_bg)
        missing_frame.pack(fill="x", pady=5)
        ttk.Label(missing_frame, text="Null Operations:", background=self.secondary_bg).pack(side="left")
        ttk.Combobox(missing_frame, textvariable=self.var_missing, values=["Ignore", "Drop Dropouts", "Fill with N/A"], state="readonly", width=14).pack(side="left", padx=10)

        self.meta_panel = ttk.Frame(middle_grid, style="Card.TFrame", padding=15)
        self.meta_panel.pack(side="right", fill="both", expand=True)
        self.render_empty_meta()

        preview_title = ttk.Label(main_layout, text="Dataset Snapshot (Top 5 Rows Preview)")
        preview_title.pack(anchor="w", pady=(5, 5))

        self.preview_frame = ttk.Frame(main_layout, style="Card.TFrame")
        self.preview_frame.pack(fill="both", expand=True)
        
        self.sheet_view = ttk.Treeview(self.preview_frame, show="headings")
        self.sheet_view.pack(fill="both", expand=True)

        self.progress_bar = ttk.Progressbar(main_layout, style="Cyber.Horizontal.TProgressbar", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(15, 5))

        footer_panel = tk.Frame(main_layout, bg=self.bg_color)
        footer_panel.pack(fill="x", pady=5)

        self.status_lbl = tk.Label(footer_panel, text="System: Core Ready Engine", bg=self.bg_color, fg=self.text_muted, font=("Segoe UI", 9, "italic"))
        self.status_lbl.pack(side="left")

        self.open_dir_btn = ttk.Button(footer_panel, text="Open Output Directory Folder", command=self.open_output_folder, state="disabled")
        self.open_dir_btn.pack(side="right", padx=5)

        self.exec_btn = ttk.Button(footer_panel, text="EXECUTE ENGINE RUN", style="Action.TButton", command=self.dispatch_processing_run)
        self.exec_btn.pack(side="right", padx=5)

    def render_empty_meta(self):
        for widget in self.meta_panel.winfo_children(): widget.destroy()
        ttk.Label(self.meta_panel, text="Analytical Diagnostics", font=("Segoe UI", 11, "bold"), background=self.secondary_bg).pack(anchor="w", pady=(0, 5))
        tk.Label(self.meta_panel, text="Load target metrics to examine profiles.", fg=self.text_muted, bg=self.secondary_bg, font=("Segoe UI", 10)).pack(anchor="w")

    def handle_file_browsing(self):
        ft = [("CSV Sheets", "*.csv")] if self.conversion_type.get() == "CSV to Excel" else [("Excel Spreadsheets", "*.xlsx")]
        path = filedialog.askopenfilename(filetypes=ft)
        if path:
            self.input_file_path.set(path)
            self.load_metadata_and_preview(path)

    def handle_dest_browsing(self):
        path = filedialog.askdirectory()
        if path: self.output_directory.set(path)

    def load_metadata_and_preview(self, path):
        # 1. Fetch data metrics summary
        info = converter.get_file_info(path)
        for widget in self.meta_panel.winfo_children(): widget.destroy()
        
        ttk.Label(self.meta_panel, text="File Analysis Diagnostics", font=("Segoe UI", 11, "bold"), background=self.secondary_bg).pack(anchor="w", pady=(0, 5))
        tk.Label(self.meta_panel, text=f"• Total Data Rows: {info['rows']}", fg="white", bg=self.secondary_bg, font=("Segoe UI", 10)).pack(anchor="w", pady=1)
        tk.Label(self.meta_panel, text=f"• Column Metrics: {info['cols']}", fg="white", bg=self.secondary_bg, font=("Segoe UI", 10)).pack(anchor="w", pady=1)
        tk.Label(self.meta_panel, text=f"• Storage Space: {info['size']}", fg="white", bg=self.secondary_bg, font=("Segoe UI", 10)).pack(anchor="w", pady=1)

        # 2. Hard split structure table processing
        try:
            if path.endswith('.csv'):
                df = pd.read_csv(path, nrows=5, sep=',', engine='python', encoding='utf-8-sig')
                
                # Ultimate Force Logic: Agar pandas fail hua toh hard split invoke karenge
                if df.shape[1] == 1:
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        lines = [f.readline().strip().split(',') for _ in range(6)]
                    columns_list = lines[0]
                    rows_data = lines[1:]
                else:
                    columns_list = list(df.columns)
                    rows_data = [list(row.values) for _, row in df.iterrows()]
            else:
                df = pd.read_excel(path, nrows=5, engine='openpyxl')
                columns_list = list(df.columns)
                rows_data = [list(row.values) for _, row in df.iterrows()]
            
            # Treeview Completely Re-initialization
            self.sheet_view.delete(*self.sheet_view.get_children())
            self.sheet_view.configure(columns=columns_list)
            
            available_width = 800
            col_width = max(90, int(available_width / max(1, len(columns_list))))

            for col in columns_list:
                self.sheet_view.heading(col, text=col, anchor="center")
                self.sheet_view.column(col, width=col_width, minwidth=70, anchor="center")

            for row in rows_data:
                formatted_row = [str(val) if pd.notna(val) and val != '' else "NULL" for val in row]
                self.sheet_view.insert("", "end", values=formatted_row)
        except Exception as e:
            print(f"[DEBUG ERROR] Preview Grid Layout Build Failure: {str(e)}")
            logger.error(f"Failed grid render view: {str(e)}")

    def update_progress(self, val):
        self.progress_bar['value'] = val
        self.root.update_idletasks()

    def dispatch_processing_run(self):
        if not self.input_file_path.get():
            messagebox.showwarning("Incomplete Execution", "Load a proper file vector configuration matrix.")
            return

        self.exec_btn.config(state="disabled")
        self.status_lbl.config(text="System Processing Pipeline Active...", fg=self.accent_color)
        
        config = {
            "remove_duplicates": self.var_dup.get(),
            "handle_missing": self.var_missing.get(),
            "parse_dates": self.var_date.get()
        }

        threading.Thread(target=self.async_worker, args=(config,), daemon=True).start()

    def async_worker(self, config):
        try:
            res = converter.process_and_convert(
                self.input_file_path.get(),
                self.output_directory.get(),
                self.conversion_type.get(),
                config,
                self.update_progress
            )
            self.status_lbl.config(text="✓ Production Build Run Complete", fg="#22c55e")
            self.open_dir_btn.config(state="normal")
            messagebox.showinfo("Pipeline Completed", f"Successfully structured out target payload at:\n{res}")
        except Exception as e:
            self.status_lbl.config(text="⚠ Pipeline Interruption System Crash", fg="#ef4444")
            messagebox.showerror("System Error Traceback Exception", str(e))
        finally:
            self.exec_btn.config(state="normal")

    def open_output_folder(self):
        out_path = self.output_directory.get()
        if os.path.exists(out_path):
            if os.name == 'nt':
                subprocess.Popen(f'explorer "{out_path}"')
            else:
                subprocess.Popen(['xdg-open', out_path])