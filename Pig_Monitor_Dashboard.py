from datetime import datetime
from dateutil.relativedelta import relativedelta
import xlsxwriter
import os
import tkinter as tk
import pandas as pd
import numpy as np
import Pig_Monitor as Pg
from tkinter import messagebox, ttk


class PigMonitorDashboard(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Pig Monitor")

        # Launch in Maximized (Zoomed) Mode
        self.state("zoomed")
        self.minsize(980, 620)

        # Colour Palette & Styling Setup
        self.configure(bg="#f1f5f9")

        self.style = ttk.Style(self)
        self.tk.call("source", "forest-light.tcl")
        self.style.theme_use("forest-light")

        self._apply_custom_styles()

        # Chronologically sort available year directories
        years = sorted([item for item in os.listdir() if 'figures 20' in item and os.path.isdir(item)])

        if years:
            last_year_inputs = sorted([item for item in os.listdir(years[-1]) if 'inputs' in item])
            self.inp_dates = [f"{item[:2]}/{item[3:5]}" for item in last_year_inputs]
        else:
            self.inp_dates = []

        self.fig_dates = []
        for year in years:
            fig_files = sorted([item for item in os.listdir(year) if 'figures' in item])
            self.fig_dates += [f"{item[:2]}-{item[3:5]}" for item in fig_files]

        # Sort figure dates chronologically by month and year
        self.fig_dates.sort(key=lambda x: datetime.strptime(x, "%m-%y"))

        if self.fig_dates:
            self.date = self.fig_dates[-1]
        else:
            self.date = "12-22"

        # dates to collect data from
        date = datetime.strptime(self.date, "%m-%y")
        months_range = (int(date.strftime("%m-%y")[3:]) - 23) * 12 + 3 * (int(self.date[:2]) // 3) + 3
        dates = [date - relativedelta(months=i) for i in range(0, months_range, 3)]

        # collect data from Excel sheets
        self.data = []
        for i in [date.strftime("%m-%y") for date in dates]:
            if i[:2] == self.date[:2]:
                data_block = []
                for j in ['Input1', 'Input2', 'KPF']:
                    excel_data = pd.read_excel(os.path.join(f"figures 20{i[3:]}", f"{i} figures.xlsx"), j)
                    data_block += excel_data.values.tolist()
                if not self.data:
                    for k in range(len(data_block)):
                        self.data.append([])
                for k in range(len(data_block)):
                    self.data[k] += (data_block[k][1:])
            else:
                excel_data = pd.read_excel(os.path.join(f"figures 20{i[3:]}", f"{i} figures.xlsx"), 'KPF')
                excel_data = excel_data.values.tolist()
                for k in range(len(excel_data) - 1, -1, -1):
                    self.data[len(self.data) - k - 1] += excel_data[::-1][k][1:]

        data_block = []
        for j in ['Input1', 'Input2']:
            excel_data = pd.read_excel(os.path.join("figures 2022", "12-22 figures.xlsx"), j)
            data_block += excel_data.values.tolist()
        for k in range(len(data_block)):
            self.data[k] += data_block[k][13 - int(self.date[:2]):]

        self.data = self.data[:46] + self.data[47:69] + self.data[72:84] + self.data[86:93] + self.data[94:]
        self.data = [row[::-1] for row in self.data]

        # Build UI Structure
        self._configure_layout()
        self._build_app_header()
        self._build_sidebar()
        self._build_top_control_bar()
        self._build_kpi_cards()
        self._build_main_canvas()

        # Initial Plots Render
        self._update_left_plot()
        self._update_right_plot()

    def _apply_custom_styles(self):
        """Custom TTK styling overrides for visual contrast."""
        self.style.configure("SoftGreen.TButton", background="#dcfce7", foreground="#166534",
                             font=("Segoe UI", 9, "bold"), padding=(8, 4))
        self.style.configure("Card.TLabelframe", background="#ffffff")
        self.style.configure("Card.TLabelframe.Label", background="#ffffff", foreground="#0f172a",
                             font=("Segoe UI", 9, "bold"))

    def _configure_layout(self):
        """Configure grid weights and spacing."""
        self.columnconfigure(0, weight=0, minsize=230)
        self.columnconfigure(1, weight=1)

        self.rowconfigure(0, weight=0)  # App Header
        self.rowconfigure(1, weight=0)  # Top Control Ribbon
        self.rowconfigure(2, weight=0)  # KPI Metrics
        self.rowconfigure(3, weight=1)  # Visual Canvas Workspace

    # ---------------------------------------------------------
    # Helper: Dates Collecting Methods
    # ---------------------------------------------------------
    def _get_next_input_label(self):
        """Calculates the label for the next input sheet not in the project."""
        try:
            years = sorted([item for item in os.listdir() if 'figures 20' in item and os.path.isdir(item)])
            input_dates = []
            for year in years:
                for file in sorted(os.listdir(year)):
                    if 'inputs' in file:
                        clean_name = file.replace('_', '-').replace('/', '-')
                        date_str = clean_name[:5]
                        try:
                            dt = datetime.strptime(date_str, "%m-%y")
                            input_dates.append(dt)
                        except ValueError:
                            continue
            if input_dates:
                next_date = max(input_dates) + relativedelta(months=1)
                return f"📄 {next_date.strftime('%m/%y')} Inputs"
        except Exception:
            pass
        return "📄 01/26 Inputs"

    def _get_update_from_dates(self):
        """
        Generates monthly date strings ('MM/YY') from 01/23 up to the most recent input sheet
        whose preceding month has a corresponding figures sheet.
        """
        try:
            years = sorted([item for item in os.listdir() if 'figures 20' in item and os.path.isdir(item)])
            fig_dates = set()
            inp_dates = set()

            for year in years:
                for file in sorted(os.listdir(year)):
                    clean_name = file.replace('_', '-').replace('/', '-')
                    if len(clean_name) >= 5:
                        date_str = clean_name[:5]
                        try:
                            dt = datetime.strptime(date_str, "%m-%y")
                            if 'figures' in file:
                                fig_dates.add(dt)
                            elif 'inputs' in file:
                                inp_dates.add(dt)
                        except ValueError:
                            continue

            # Filter input dates where (month - 1) exists in figure sheets
            valid_inp_dates = [d for d in inp_dates if (d - relativedelta(months=1)) in fig_dates]

            if valid_inp_dates:
                max_date = max(valid_inp_dates)
            elif fig_dates:
                max_date = max(fig_dates) + relativedelta(months=1)
            else:
                max_date = datetime(2023, 1, 1)

            start_date = datetime(2023, 1, 1)
            if max_date < start_date:
                max_date = start_date

            date_list = []
            curr = start_date
            while curr <= max_date:
                date_list.append(curr.strftime("%m/%y"))
                curr += relativedelta(months=1)

            return date_list
        except Exception:
            return ["01/23"]

    # ---------------------------------------------------------
    # 1. App Header & Help Action
    # ---------------------------------------------------------
    def _build_app_header(self):
        header_card = ttk.Frame(self, padding=(12, 10))
        header_card.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(10, 4))

        logo_canvas = tk.Canvas(header_card, width=38, height=38, bg="#166534", highlightthickness=0)
        logo_canvas.pack(side="left", padx=(0, 12))
        logo_canvas.create_text(19, 19, text="PM", fill="#ffffff", font=("Segoe UI", 13, "bold"))

        title_box = ttk.Frame(header_card)
        title_box.pack(side="left")

        lbl_title = tk.Label(title_box, text="Pig Monitor", font=("Segoe UI", 15, "bold"), fg="#0f172a")
        lbl_title.pack(anchor="w")

        lbl_subtitle = tk.Label(title_box, text="Livestock Performance & Reporting Suite v2.0", font=("Segoe UI", 8),
                                fg="#64748b")
        lbl_subtitle.pack(anchor="w")

        user_frame = ttk.Frame(header_card)
        user_frame.pack(side="right")

        btn_help = ttk.Button(user_frame, text="❓ Help", width=8, command=self._show_help_dialog)
        btn_help.pack(side="left")

    def _show_help_dialog(self):
        """Displays user assistance."""
        messagebox.showinfo(
            "Pig Monitor Assistance",
            "Pig Monitor Dashboard v2.0\n\n"
            "To create a new report - \n"
            "• Create the next input sheet\n"
            "• Fill out input sheet with that months figures\n"
            "• ▶ Run Report\n\n"
            "To fix previous inputs - \n"
            "• Correct the data on the input sheet\n"
            "• ▶ Run Report updating from that month\n")

    # ---------------------------------------------------------
    # 2. Sidebar Navigation Card
    # ---------------------------------------------------------
    def _build_sidebar(self):
        sidebar_frame = ttk.LabelFrame(self, text=" Navigation ", style="Card.TLabelframe", padding=8)
        sidebar_frame.grid(row=1, column=0, rowspan=3, sticky="nsew", padx=(10, 5), pady=(4, 10))
        sidebar_frame.rowconfigure(0, weight=1)
        sidebar_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(sidebar_frame, selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(sidebar_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.grid(row=0, column=1, sticky="ns")

        self.tree.heading("#0", text="Data Directory", anchor="w")

        # Define treeview data
        tree_data = []
        years = sorted([item for item in os.listdir() if 'figures 20' in item and os.path.isdir(item)], reverse=True)

        for i, v in enumerate(years):
            # append the years
            tree_data.append(["", len(tree_data) + 1, v[8:]])

            # append the figures
            tree_data.append([len(tree_data), len(tree_data) + 1, "Figures"])
            figs = sorted([item for item in os.listdir(v) if 'figures' in item], reverse=True)
            for j, w in enumerate(figs):
                tree_data.append([len(tree_data) - j, len(tree_data) + 1, w[:5]])

            # append the input sheets
            tree_data.append([len(tree_data) - len(figs) - 1, len(tree_data) + 1, "Inputs"])
            inputs = sorted([item for item in os.listdir(v) if 'inputs' in item], reverse=True)
            for j, w in enumerate(inputs):
                tree_data.append([len(tree_data) - j, len(tree_data) + 1, w[:5]])

        self.tree.tag_configure("height", font=("Segoe UI", 12))

        # Insert treeview data
        for item in tree_data:
            self.tree.insert(parent=item[0], index='end', iid=item[1], text=item[2])
            if not item[0]:
                self.tree.item(item[1], tags="height")

        # keep top year expanded
        if tree_data:
            self.tree.item(tree_data[0][1], open=True)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_selected)

    def on_tree_selected(self, event=None):
        selection = self.tree.selection()
        if not selection:
            return
        ele = selection[0]

        if self.tree.parent(ele) != "":
            if self.tree.parent(self.tree.parent(ele)) != "":
                self.date = self.tree.item(ele)["text"]
                if self.tree.item([self.tree.parent(ele)])["text"] == "Figures":

                    # obtain dates to collect data from
                    date = datetime.strptime(self.date, "%m-%y")
                    months_range = (int(date.strftime("%m-%y")[3:]) - 23) * 12 + 3 * (int(self.date[:2]) // 3) + 3
                    dates = [date - relativedelta(months=i) for i in range(0, months_range, 3)]

                    # collect data from Excel sheets
                    self.data = []
                    for i in [date.strftime("%m-%y") for date in dates]:
                        if i[:2] == self.date[:2]:
                            data_block = []
                            for j in ['Input1', 'Input2', 'KPF']:
                                excel_data = pd.read_excel(os.path.join(f"figures 20{i[3:]}", f"{i} figures.xlsx"), j)
                                data_block += excel_data.values.tolist()
                            if not self.data:
                                for k in range(len(data_block)):
                                    self.data.append([])
                            for k in range(len(data_block)):
                                self.data[k] += (data_block[k][1:])
                        else:
                            excel_data = pd.read_excel(os.path.join(f"figures 20{i[3:]}", f"{i} figures.xlsx"), 'KPF')
                            excel_data = excel_data.values.tolist()
                            for k in range(len(excel_data) - 1, -1, -1):
                                self.data[len(self.data) - k - 1] += excel_data[::-1][k][1:]

                    data_block = []
                    for j in ['Input1', 'Input2']:
                        excel_data = pd.read_excel(os.path.join("figures 2022", "12-22 figures.xlsx"), j)
                        data_block += excel_data.values.tolist()
                    for k in range(len(data_block)):
                        self.data[k] += data_block[k][13 - int(self.date[:2]):]

                    self.data = self.data[:46] + self.data[47:69] + self.data[72:84] + self.data[86:93] + self.data[94:]
                    self.data = [row[::-1] for row in self.data]

                    for i in self.data:
                        print(i)

                    self._update_left_plot()
                    self._update_right_plot()

                    self._build_kpi_cards()

                elif self.tree.item([self.tree.parent(ele)])["text"] == "Inputs":
                    os.startfile(os.path.join(f"figures 20{self.date[3:]}", f"{self.date} inputs.xlsx"))

    def _refresh_treeview(self):
        """Clears and rebuilds the navigation treeview with updated files."""
        # Clear existing treeview items
        for item in self.tree.get_children():
            self.tree.delete(item)

        tree_data = []
        years = sorted([item for item in os.listdir() if 'figures 20' in item and os.path.isdir(item)], reverse=True)

        for i, v in enumerate(years):
            # Append the years
            tree_data.append(["", len(tree_data) + 1, v[8:]])

            # Append the figures
            tree_data.append([len(tree_data), len(tree_data) + 1, "Figures"])
            figs = sorted([item for item in os.listdir(v) if 'figures' in item], reverse=True)
            for j, w in enumerate(figs):
                tree_data.append([len(tree_data) - j, len(tree_data) + 1, w[:5]])

            # Append the inputs
            tree_data.append([len(tree_data) - len(figs) - 1, len(tree_data) + 1, "Inputs"])
            inputs = sorted([item for item in os.listdir(v) if 'inputs' in item], reverse=True)
            for j, w in enumerate(inputs):
                tree_data.append([len(tree_data) - j, len(tree_data) + 1, w[:5]])

        # Re-insert items into treeview
        for item in tree_data:
            self.tree.insert(parent=item[0], index='end', iid=item[1], text=item[2])
            if not item[0]:
                self.tree.item(item[1], tags="height")

        # Keep top year expanded
        if tree_data:
            self.tree.item(tree_data[0][1], open=True)

    # ---------------------------------------------------------
    # 3. Top Line Commands: Input Data & Report Ribbon
    # ---------------------------------------------------------
    def _build_top_control_bar(self):
        control_frame = ttk.LabelFrame(self, text=" Input Data & Report Generation ", style="Card.TLabelframe",
                                       padding=8)
        control_frame.grid(row=1, column=1, sticky="ew", padx=(5, 10), pady=(4, 2))

        # Dynamic label & method callback attached
        self.btn_active_sheet = ttk.Button(
            control_frame,
            text=self._get_next_input_label(),
            width=14,
            command=self.on_active_sheet_click
        )
        self.btn_active_sheet.pack(side="left", padx=(0, 4))

        self.btn_edit_sheet = ttk.Button(
            control_frame,
            text="📝 Edit",
            width=7,
            command=self.on_edit_sheet_click
        )
        self.btn_edit_sheet.pack(side="left", padx=(0, 4))

        self.btn_delete_sheet = ttk.Button(
            control_frame,
            text="🗑 Bin",
            width=6,
            command=self.on_delete_sheet_click
        )
        self.btn_delete_sheet.pack(side="left", padx=(0, 15))

        ttk.Separator(control_frame, orient="vertical").pack(side="left", fill="y", padx=8)

        ttk.Label(control_frame, text="Update From:").pack(side="left", padx=(8, 4))

        # Spinbox containing monthly dates from 01/23 up to max valid input sheet date
        update_dates = self._get_update_from_dates()
        self.date_sb = ttk.Spinbox(control_frame, values=update_dates, width=7, state="readonly")
        if update_dates:
            self.date_sb.set(update_dates[-1])  # Default to most recent month
        self.date_sb.pack(side="left", padx=(0, 10))

        self.btn_run = ttk.Button(
            control_frame,
            text="▶ Run Report",
            style="Accent.TButton",
            command=self.on_run_report_click
        )
        self.btn_run.pack(side="left", padx=(0, 15))

        # Progress bar inserted between Run Report button and the divider
        self.progress_bar = ttk.Progressbar(
            control_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            value=100
        )
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 15))

        self.btn_excel = ttk.Button(
            control_frame,
            text="📊 Open Report in Excel",
            style="SoftGreen.TButton",
            command=self.on_open_excel_click
        )
        self.btn_excel.pack(side="right", padx=(8, 0))

        ttk.Separator(control_frame, orient="vertical").pack(side="right", fill="y", padx=8)

    # ---------------------------------------------------------
    # 4. KPI Summary Cards
    # ---------------------------------------------------------
    def _build_kpi_cards(self):
        kpi_container = ttk.Frame(self, padding=0)
        kpi_container.grid(row=2, column=1, sticky="ew", padx=(5, 10), pady=4)
        for col in range(4):
            kpi_container.columnconfigure(col, weight=1)

        metrics = [
            ["Piglets Weaned per Litter", self.data[-11][-3], (self.data[-11][-3] / self.data[-11][-6] - 1) * 100, '',
             ''],
            ["Liveweight FCR", self.data[-3][-3], (self.data[-3][-3] / self.data[-3][-6] - 1) * 100, '', ''],
            ["Daily Liveweight Gain", f"{self.data[-4][-3]} g", (self.data[-4][-3] / self.data[-4][-6] - 1) * 100, '',
             ''],
            ["Average P2", f"{self.data[-2][-3]} mm", (self.data[-2][-3] / self.data[-2][-6] - 1) * 100, '', '']]

        for i, item in enumerate(metrics):
            if item[2] > 0:
                metrics[i][2] = f"▲ {metrics[i][2]:+.1f}%"
                metrics[i][3], metrics[i][4] = "#dcfce7", "#166534"
            elif item[2] == 0:
                metrics[i][2] = f"■ {metrics[i][2]:+.1f}%"
                metrics[i][3], metrics[i][4] = "#f1f5f9", "#475569"
            else:
                metrics[i][2] = f"▼ {metrics[i][2]:+.1f}%"
                metrics[i][3], metrics[i][4] = "#e0f2fe", "#0369a1"

        for idx, (title, val, badge_text, bg_color, fg_color) in enumerate(metrics):
            card = ttk.LabelFrame(kpi_container, text=f" {title} ", style="Card.TLabelframe", padding=6)
            card.grid(row=0, column=idx, sticky="ew", padx=3)

            lbl_val = tk.Label(card, text=val, font=("Segoe UI", 12, "bold"), fg="#0f172a", bg="#ffffff")
            lbl_val.pack(anchor="w")

            pill_canvas = tk.Canvas(card, height=22, bg="#ffffff", highlightthickness=0)
            pill_canvas.pack(fill="x", pady=(2, 0))
            self._draw_pill_badge(pill_canvas, badge_text, bg_color=bg_color, fg_color=fg_color)

    def _draw_pill_badge(self, canvas, text, bg_color, fg_color):
        """Draws a rounded status badge for summary cards."""
        canvas.delete("all")
        canvas.create_oval(2, 2, 18, 18, fill=bg_color, outline="")
        canvas.create_oval(112, 2, 128, 18, fill=bg_color, outline="")
        canvas.create_rectangle(10, 2, 120, 18, fill=bg_color, outline="")

        canvas.create_text(65, 10, text=text, fill=fg_color, font=("Segoe UI", 8, "bold"))

    # ---------------------------------------------------------
    # 5. Main Canvas Workspace with Interactive Plot Controls
    # ---------------------------------------------------------
    def _build_main_canvas(self):
        canvas_container = ttk.Frame(self, padding=0)
        canvas_container.grid(row=3, column=1, sticky="nsew", padx=(5, 10), pady=(4, 10))
        canvas_container.columnconfigure(0, weight=1)
        canvas_container.columnconfigure(1, weight=1)
        canvas_container.rowconfigure(0, weight=1)

        # Left Plot Card and Controls
        plot1_frame = ttk.LabelFrame(canvas_container, text=" Monthly Line Plot ", style="Card.TLabelframe", padding=6)
        plot1_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        p1_bar = ttk.Frame(plot1_frame, padding=(0, 0, 0, 6))
        p1_bar.pack(fill="x", side="top")

        ttk.Label(p1_bar, text="KPI:").pack(side="left", padx=(0, 4))
        self.p1_category_cb = ttk.Combobox(p1_bar,
                                           values=["Sows & Served Gilts",
                                                   "Piglets Born Alive/Litter",
                                                   "Piglets Weaned/Litter",
                                                   "Pigs on unit",
                                                   "Liveweight/Pig Weaned",
                                                   "Liveweight/Pig out",
                                                   "P2"],
                                           width=18, state="readonly")
        self.p1_category_cb.set("Piglets Weaned/Litter")
        self.p1_category_cb.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.p1_category_cb.bind("<<ComboboxSelected>>", lambda e: self._update_left_plot())

        ttk.Label(p1_bar, text="Format:").pack(side="left", padx=(0, 4))
        self.p1_format_cb = ttk.Combobox(p1_bar, values=["Standard", "2-Yr Comparison", "Historical"], width=14,
                                         state="readonly")
        self.p1_format_cb.set("2-Yr Comparison")
        self.p1_format_cb.pack(side="left", fill="x", expand=True, padx=(0, 0))
        self.p1_format_cb.bind("<<ComboboxSelected>>", lambda e: self._update_left_plot())

        self.canvas_1 = tk.Canvas(plot1_frame, bg="#f8fafc", highlightthickness=0)
        self.canvas_1.pack(fill="both", expand=True)
        self.canvas_1.bind("<Configure>", lambda e: self._update_left_plot())

        # Right Plot Card and Controls
        plot2_frame = ttk.LabelFrame(canvas_container, text=" Averaged Monthly Bar Chart ", style="Card.TLabelframe",
                                     padding=6)
        plot2_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        p2_bar = ttk.Frame(plot2_frame, padding=(0, 0, 0, 6))
        p2_bar.pack(fill="x", side="top")

        ttk.Label(p2_bar, text="KPI:").pack(side="left", padx=(0, 4))
        self.p2_kpi_cb = ttk.Combobox(p2_bar,
                                      values=["Sows & Served Gilts",
                                              "Farrowing Rate",
                                              "Litters/Sow/Year",
                                              "Piglets Born Alive/Litter",
                                              "Piglets Weaned/Litter",
                                              "Piglet Mortality to Weaning",
                                              "Piglets/Sow/Year",
                                              "Pigs on unit",
                                              "Liveweight/Pig Weaned",
                                              "Liveweight/Pig out",
                                              "Daily Feed Intake/Pig",
                                              "Daily Liveweight Gain",
                                              "Liveweight FCR",
                                              "P2",
                                              "Daily Lean Gain"],
                                      width=19, state="readonly")
        self.p2_kpi_cb.set("Daily Liveweight Gain")
        self.p2_kpi_cb.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.p2_kpi_cb.bind("<<ComboboxSelected>>", lambda e: self._update_right_plot())

        ttk.Label(p2_bar, text="Format:").pack(side="left", padx=(0, 4))
        self.p2_format_cb = ttk.Combobox(p2_bar, values=["Standard", "Historical"], width=10, state="readonly")
        self.p2_format_cb.set("Standard")
        self.p2_format_cb.pack(side="left", fill="x", expand=True, padx=(0, 0))
        self.p2_format_cb.bind("<<ComboboxSelected>>", lambda e: self._update_right_plot())

        self.canvas_2 = tk.Canvas(plot2_frame, bg="#f8fafc", highlightthickness=0)
        self.canvas_2.pack(fill="both", expand=True)
        self.canvas_2.bind("<Configure>", lambda e: self._update_right_plot())

    # ---------------------------------------------------------
    # Left Plot Generating
    # ---------------------------------------------------------
    def _update_left_plot(self):
        plts = {"Sows & Served Gilts": {"row": 7, "label": ""},
                "Pigs on unit": {"row": 46, "label": ""},
                "P2": {"row": 67, "label": "mm"},
                "Liveweight/Pig Weaned": {"row": 50, "label": "kg"},
                "Liveweight/Pig out": {"row": 66, "label": "kg"}}

        div_plts = {"Piglets Born Alive/Litter": {"row": [26, 25], "label": ""},
                    "Piglets Weaned/Litter": {"row": [31, 30], "label": ""}}

        category = self.p1_category_cb.get()
        fmt = self.p1_format_cb.get()

        if category in ["Piglets Born Alive/Litter", "Piglets Weaned/Litter"]:
            data = [a / b for a, b in
                    zip(self.data[div_plts[category]["row"][0]], self.data[div_plts[category]["row"][1]])]
            y_label = div_plts[category]["label"]
        else:
            data = self.data[plts[category]["row"]]
            y_label = plts[category]["label"]

        if fmt == "Standard":
            data = [(i, data[i + len(self.data[0]) - 12]) for i in range(0, 12)]
            ticks = [self.data[0][i + len(self.data[0]) - 12] for i in range(0, 12)]

            plot_range = (max([a[1] for a in data]) - min([a[1] for a in data])) * 1.5
            avg_val = (max([a[1] for a in data]) + min([a[1] for a in data])) / 2
            min_y = avg_val - plot_range / 2
            max_y = avg_val + plot_range / 2

            self._draw_single_line_plot(self.canvas_1, title=f"{category}", data_points=data, x_ticks=ticks,
                                        line_color="#15803d", fill_color="#dcfce7", y_label=y_label, min_y_val=min_y,
                                        max_y_val=max_y)
        elif fmt == "2-Yr Comparison":
            data_curr = [(i, data[i + len(self.data[0]) - 12]) for i in range(0, 12)]
            data_prev = [(i, data[i + len(self.data[0]) - 24]) for i in range(0, 12)]
            ticks = [int(self.data[0][i + len(self.data[0]) - 12][:2]) for i in range(0, 12)]

            months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            ticks = [months[i - 1] for i in ticks]

            plot_range = (max([a[1] for a in data_curr + data_prev]) - min([a[1] for a in data_curr + data_prev])) * 1.5
            avg_val = (max([a[1] for a in data_curr + data_prev]) + min([a[1] for a in data_curr + data_prev])) / 2
            min_y = avg_val - plot_range / 2
            max_y = avg_val + plot_range / 2

            self._draw_comparison_line_plot(self.canvas_1, title=f"{category}", data_curr=data_curr,
                                            data_prev=data_prev, x_ticks=ticks, y_label=y_label, min_y_val=min_y,
                                            max_y_val=max_y)
        elif fmt == "Historical":
            data_hist = [(i, data[i]) for i in range(len(data))]
            ticks_hist = [self.data[0][i] for i in range(len(data))]
            for i in range(len(ticks_hist)):
                if ticks_hist[i][:2] != self.date[:2]:
                    ticks_hist[i] = ''

            plot_range = (max([b for a, b in data_hist]) - min([b for a, b in data_hist])) * 1.5
            avg_val = (max([a[1] for a in data_hist]) + min([a[1] for a in data_hist])) / 2
            min_y = avg_val - plot_range / 2
            max_y = avg_val + plot_range / 2

            self._draw_single_line_plot(self.canvas_1, title=f"{category}", data_points=data_hist, x_ticks=ticks_hist,
                                        line_color="#0369a1", fill_color="#e0f2fe", y_label=y_label, min_y_val=min_y,
                                        max_y_val=max_y)

    # ---------------------------------------------------------
    # Right Plot Generating
    # ---------------------------------------------------------
    def _update_right_plot(self):
        plts = {"Sows & Served Gilts": {"row": -15, "label": ""},
                "Farrowing Rate": {"row": -14, "label": "%"},
                "Litters/Sow/Year": {"row": -13, "label": ""},
                "Piglets Born Alive/Litter": {"row": -12, "label": ""},
                "Piglets Weaned/Litter": {"row": -11, "label": ""},
                "Piglet Mortality to Weaning": {"row": -10, "label": "%"},
                "Piglets/Sow/Year": {"row": -9, "label": ""},
                "Pigs on unit": {"row": -8, "label": ""},
                "Liveweight/Pig Weaned": {"row": -7, "label": "kg"},
                "Liveweight/Pig out": {"row": -6, "label": "kg"},
                "Daily Feed Intake/Pig": {"row": -5, "label": "kg"},
                "Daily Liveweight Gain": {"row": -4, "label": "g"},
                "Liveweight FCR": {"row": -3, "label": ""},
                "P2": {"row": -2, "label": "mm"},
                "Daily Lean Gain": {"row": -1, "label": "g"}}

        kpi = self.p2_kpi_cb.get()
        fmt = self.p2_format_cb.get()

        data = self.data[plts[kpi]["row"]]
        y_label = plts[kpi]["label"]

        if fmt == "Standard":
            groups = [("3 Month", data[-3], data[-15]), ("6 Month", data[-2], data[-14]),
                      ("12 Month", data[-1], data[-13])]
            vals = [data[-1], data[-2], data[-3], data[-13], data[-14], data[-15]]

            plot_range = (max(vals) - min(vals)) * 1.5
            avg_val = (max(vals) + min(vals)) / 2
            min_y = avg_val - plot_range
            max_y = avg_val + plot_range / 2

            self._draw_grouped_bar_chart(self.canvas_2, title=f"{kpi}", groups_data=groups, y_label=y_label,
                                         min_y_val=min_y, max_y_val=max_y)
        elif fmt == "Historical":
            hist_periods = [data[i] for i in range(0, len(data), 3)]

            plot_range = (max(hist_periods) - min(hist_periods)) * 1.5
            avg_val = (max(hist_periods) + min(hist_periods)) / 2
            min_y = avg_val - plot_range
            max_y = avg_val + plot_range / 2

            # obtain dates to collect data from
            date = datetime.strptime(self.date, "%m-%y")
            dates = [date - relativedelta(months=i) for i in range(0, len(data), 3)]

            ticks_hist = [f"{i[:2]}/{i[3:]}" for i in [date.strftime("%m-%y") for date in dates][::-1]]
            for i in range(len(ticks_hist)):
                if ticks_hist[i][:2] != self.date[:2]:
                    ticks_hist[i] = ''

            hist_periods = [(a, b) for a, b in zip(ticks_hist, hist_periods)]

            self._draw_multi_period_bar_chart(self.canvas_2, title=f"{kpi}", period_data=hist_periods,
                                              y_label="Metric Value", min_y_val=min_y, max_y_val=max_y)

    # ---------------------------------------------------------
    # Line Plots Formatting
    # ---------------------------------------------------------
    def _draw_single_line_plot(self, canvas, title, data_points, x_ticks, line_color, fill_color, y_label, min_y_val=0,
                               max_y_val=200):
        canvas.delete("all")
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w < 10 or h < 10:
            return

        left_m, right_m, top_m, bottom_m = 50, 25, 35, 35
        plot_w = max(w - left_m - right_m, 1)
        plot_h = max(h - top_m - bottom_m, 1)

        canvas.create_text(w / 2, 14, text=title, fill="#0f172a", font=("Segoe UI", 9, "bold"))
        canvas.create_text(left_m - 5, top_m - 15, text=f"↑ {y_label}", fill="#64748b", font=("Segoe UI", 7, "bold"),
                           anchor="w")

        # Calculate dynamic Y range and generated tick marks
        y_range = max([max_y_val - min_y_val, 0.1])  # Guard against division by zero
        num_y_ticks = 5
        y_step = y_range / (num_y_ticks - 1)
        y_ticks = [min_y_val + i * y_step for i in range(num_y_ticks)]

        for val in y_ticks:
            y_pos = (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h
            canvas.create_line(left_m, y_pos, w - right_m, y_pos, fill="#e2e8f0", dash=(2, 2))

            # Display as integer if clean division, otherwise formatted decimal
            tick_label = f"{int(val)}" if val == int(val) else f"{val:.1f}"
            canvas.create_text(left_m - 8, y_pos, text=tick_label, fill="#64748b", font=("Segoe UI", 7), anchor="e")

        x_denom = max(len(x_ticks) - 1, 1)
        for idx, label in enumerate(x_ticks):
            x_pos = left_m + (idx / x_denom) * plot_w
            canvas.create_text(x_pos, h - bottom_m + 12, text=label, fill="#64748b", font=("Segoe UI", 7))

        canvas.create_line(left_m, top_m, left_m, h - bottom_m, fill="#cbd5e1", width=1.5)
        canvas.create_line(left_m, h - bottom_m, w - right_m, h - bottom_m, fill="#cbd5e1", width=1.5)

        if not data_points:
            return

        scaled_pts = []
        for x_idx, val in data_points:
            xs = left_m + (x_idx / x_denom) * plot_w
            ys = (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h
            scaled_pts.append((xs, ys))

        poly_points = [scaled_pts[0][0], h - bottom_m]
        for pt in scaled_pts:
            poly_points.extend(pt)
        poly_points.extend([scaled_pts[-1][0], h - bottom_m])

        canvas.create_polygon(poly_points, fill=fill_color, outline="", smooth=False)

        for i in range(len(scaled_pts) - 1):
            canvas.create_line(scaled_pts[i], scaled_pts[i + 1], fill=line_color, width=2.5, smooth=True)

        for xs, ys in scaled_pts:
            canvas.create_oval(xs - 3.5, ys - 3.5, xs + 3.5, ys + 3.5, fill="#ffffff", outline=line_color, width=2)

    def _draw_comparison_line_plot(self, canvas, title, data_curr, data_prev, x_ticks, y_label, min_y_val=0,
                                   max_y_val=200):
        """Draws 2-year comparison where the fill of the lower line is always rendered on top."""
        canvas.delete("all")
        w, h = canvas.winfo_width(), canvas.winfo_height()
        if w < 10 or h < 10:
            return

        left_m, right_m, top_m, bottom_m = 50, 25, 40, 35
        plot_w = max(w - left_m - right_m, 1)
        plot_h = max(h - top_m - bottom_m, 1)

        # Centered Title
        canvas.create_text(w / 2, 14, text=title, fill="#0f172a", font=("Segoe UI", 9, "bold"))

        # Style Definitions
        color_curr = "#15803d"  # Emerald Green
        fill_curr = "#dcfce7"

        color_prev = "#64748b"  # Slate Grey
        fill_prev = "#e2e8f0"

        # 2-yr Comparison Legend
        canvas.create_line(w - 93, 14, w - 85, 14, fill=color_curr, width=2.5)
        canvas.create_text(w - 82, 14, text="Current", fill="#0f172a", font=("Segoe UI", 6, "bold"), anchor="w")

        canvas.create_line(w - 49, 14, w - 41, 14, fill=color_prev, width=2.5)
        canvas.create_text(w - 38, 14, text="Prev Year", fill="#0f172a", font=("Segoe UI", 6, "bold"), anchor="w")

        canvas.create_text(left_m - 5, top_m - 15, text=f"↑ {y_label}", fill="#64748b", font=("Segoe UI", 7, "bold"),
                           anchor="w")

        # Dynamic Y range and tick calculation
        y_range = max([max_y_val - min_y_val, 0.1])
        num_y_ticks = 5
        y_step = y_range / (num_y_ticks - 1)
        y_ticks = [min_y_val + i * y_step for i in range(num_y_ticks)]

        for val in y_ticks:
            y_pos = (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h
            canvas.create_line(left_m, y_pos, w - right_m, y_pos, fill="#e2e8f0", dash=(2, 2))

            tick_label = f"{int(val)}" if val == int(val) else f"{val:.1f}"
            canvas.create_text(left_m - 8, y_pos, text=tick_label, fill="#64748b", font=("Segoe UI", 7), anchor="e")

        x_denom = max(len(x_ticks) - 1, 1)
        for idx, label in enumerate(x_ticks):
            x_pos = left_m + (idx / x_denom) * plot_w
            canvas.create_text(x_pos, h - bottom_m + 12, text=label, fill="#64748b", font=("Segoe UI", 7))

        canvas.create_line(left_m, top_m, left_m, h - bottom_m, fill="#cbd5e1", width=1.5)
        canvas.create_line(left_m, h - bottom_m, w - right_m, h - bottom_m, fill="#cbd5e1", width=1.5)

        # Scaled points computation
        dict_curr = {
            x_idx: (left_m + (x_idx / x_denom) * plot_w, (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h, val)
            for x_idx, val in data_curr}
        dict_prev = {
            x_idx: (left_m + (x_idx / x_denom) * plot_w, (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h, val)
            for x_idx, val in data_prev}

        pts_curr = [(xs, ys) for xs, ys, _ in dict_curr.values()]
        pts_prev = [(xs, ys) for xs, ys, _ in dict_prev.values()]

        # Render dynamic overlapping fill areas segment-by-segment
        common_indices = sorted(list(set(dict_curr.keys()) & set(dict_prev.keys())))
        y_base = h - bottom_m

        for k in range(len(common_indices) - 1):
            idx1, idx2 = common_indices[k], common_indices[k + 1]
            xs1, ys_c1, v_c1 = dict_curr[idx1]
            _, ys_p1, v_p1 = dict_prev[idx1]
            xs2, ys_c2, v_c2 = dict_curr[idx2]
            _, ys_p2, v_p2 = dict_prev[idx2]

            d1 = v_c1 - v_p1
            d2 = v_c2 - v_p2

            # Check if lines cross between idx1 and idx2
            if d1 * d2 < 0:
                t = abs(d1) / (abs(d1) + abs(d2))
                xs_int = xs1 + t * (xs2 - xs1)
                ys_c_int = ys_c1 + t * (ys_c2 - ys_c1)
                ys_p_int = ys_p1 + t * (ys_p2 - ys_p1)

                subsegments = [
                    (xs1, ys_c1, ys_p1, d1, xs_int, ys_c_int, ys_p_int),
                    (xs_int, ys_c_int, ys_p_int, d2, xs2, ys_c2, ys_p2)
                ]
            else:
                diff_sign = d1 if d1 != 0 else d2
                subsegments = [(xs1, ys_c1, ys_p1, diff_sign, xs2, ys_c2, ys_p2)]

            for xa, ya_c, ya_p, diff, xb, yb_c, yb_p in subsegments:
                poly_c = [xa, y_base, xa, ya_c, xb, yb_c, xb, y_base]
                poly_p = [xa, y_base, xa, ya_p, xb, yb_p, xb, y_base]

                # Draw higher line fill first, lower line fill second (on top)
                if diff >= 0:
                    canvas.create_polygon(poly_c, fill=fill_curr, outline="", smooth=False)
                    canvas.create_polygon(poly_p, fill=fill_prev, outline="", smooth=False)
                else:
                    canvas.create_polygon(poly_p, fill=fill_prev, outline="", smooth=False)
                    canvas.create_polygon(poly_c, fill=fill_curr, outline="", smooth=False)

        # Draw Lines & Markers over top of fills
        if pts_prev:
            for i in range(len(pts_prev) - 1):
                canvas.create_line(pts_prev[i], pts_prev[i + 1], fill=color_prev, width=2.5, smooth=True)
            for xs, ys in pts_prev:
                canvas.create_oval(xs - 3.5, ys - 3.5, xs + 3.5, ys + 3.5, fill="#ffffff", outline=color_prev, width=2)

        if pts_curr:
            for i in range(len(pts_curr) - 1):
                canvas.create_line(pts_curr[i], pts_curr[i + 1], fill=color_curr, width=2.5, smooth=True)
            for xs, ys in pts_curr:
                canvas.create_oval(xs - 3.5, ys - 3.5, xs + 3.5, ys + 3.5, fill="#ffffff", outline=color_curr, width=2)

    # ---------------------------------------------------------
    # Bar Charts Formatting
    # ---------------------------------------------------------
    def _draw_grouped_bar_chart(self, canvas, title, groups_data, y_label, min_y_val=0, max_y_val=200):
        """Standard format: 6 bars organized in 3 pairs with tightly compacted legend."""
        canvas.delete("all")
        w, h = canvas.winfo_width(), canvas.winfo_height()

        left_m, right_m, top_m, bottom_m = 50, 25, 40, 35
        plot_w = max(w - left_m - right_m, 1)
        plot_h = max(h - top_m - bottom_m, 1)

        # Centered Title
        canvas.create_text(w / 2, 14, text=title, fill="#0f172a", font=("Segoe UI", 9, "bold"))

        color_curr = "#166534"
        color_prev = "#64748b"

        # Tightly Compacted Legend pushed to far right edge
        canvas.create_rectangle(w - 92, 10, w - 85, 17, fill=color_curr, outline="")
        canvas.create_text(w - 82, 14, text="Current", fill="#0f172a", font=("Segoe UI", 6, "bold"), anchor="w")

        canvas.create_rectangle(w - 48, 10, w - 41, 17, fill=color_prev, outline="")
        canvas.create_text(w - 38, 14, text="Prev Year", fill="#0f172a", font=("Segoe UI", 6, "bold"), anchor="w")

        canvas.create_text(left_m - 5, top_m - 15, text=f"↑ {y_label}", fill="#64748b", font=("Segoe UI", 7, "bold"),
                           anchor="w")

        # Dynamic Y range and tick generation
        y_range = max([max_y_val - min_y_val, 0.1])
        num_y_ticks = 5
        y_step = y_range / (num_y_ticks - 1)
        y_ticks = [min_y_val + i * y_step for i in range(num_y_ticks)]

        for val in y_ticks:
            y_pos = (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h
            canvas.create_line(left_m, y_pos, w - right_m, y_pos, fill="#e2e8f0", dash=(2, 2))

            tick_label = f"{int(val)}" if val == int(val) else f"{val:.1f}"
            canvas.create_text(left_m - 8, y_pos, text=tick_label, fill="#64748b", font=("Segoe UI", 7), anchor="e")

        canvas.create_line(left_m, top_m, left_m, h - bottom_m, fill="#cbd5e1", width=1.5)
        canvas.create_line(left_m, h - bottom_m, w - right_m, h - bottom_m, fill="#cbd5e1", width=1.5)

        num_groups = max(len(groups_data), 1)
        group_w = plot_w / num_groups

        for g_idx, (group_label, val_curr, val_prev) in enumerate(groups_data):
            center_x = left_m + (g_idx + 0.5) * group_w

            bar_w = min(group_w * 0.28, 32)
            gap = 4

            # Current Bar
            x1_c = center_x - bar_w - (gap / 2)
            x2_c = center_x - (gap / 2)
            y1_c = (h - bottom_m) - ((val_curr - min_y_val) / y_range) * plot_h
            y2_c = h - bottom_m

            canvas.create_rectangle(x1_c, y1_c, x2_c, y2_c, fill=color_curr, outline="")
            canvas.create_text((x1_c + x2_c) / 2, y1_c - 8, text=str(val_curr), fill=color_curr,
                               font=("Segoe UI", 7, "bold"))

            # Previous Bar
            x1_p = center_x + (gap / 2)
            x2_p = center_x + bar_w + (gap / 2)
            y1_p = (h - bottom_m) - ((val_prev - min_y_val) / y_range) * plot_h
            y2_p = h - bottom_m

            canvas.create_rectangle(x1_p, y1_p, x2_p, y2_p, fill=color_prev, outline="")
            canvas.create_text((x1_p + x2_p) / 2, y1_p - 8, text=str(val_prev), fill="#475569",
                               font=("Segoe UI", 7, "bold"))

            canvas.create_text(center_x, h - bottom_m + 14, text=f"{group_label} Avg", fill="#0f172a",
                               font=("Segoe UI", 8, "bold"))

    def _draw_multi_period_bar_chart(self, canvas, title, period_data, y_label, min_y_val=0, max_y_val=200):
        """Historical format: Multiple side-by-side bars representing successive 3M periods."""
        canvas.delete("all")
        w, h = canvas.winfo_width(), canvas.winfo_height()

        left_m, right_m, top_m, bottom_m = 50, 25, 40, 35
        plot_w = max(w - left_m - right_m, 1)
        plot_h = max(h - top_m - bottom_m, 1)

        canvas.create_text(w / 2, 14, text=title, fill="#0f172a", font=("Segoe UI", 9, "bold"))
        canvas.create_text(left_m - 5, top_m - 15, text=f"↑ {y_label}", fill="#64748b", font=("Segoe UI", 7, "bold"),
                           anchor="w")

        # Dynamic Y range and tick generation
        y_range = max([max_y_val - min_y_val, 0.1])
        num_y_ticks = 5
        y_step = y_range / (num_y_ticks - 1)
        y_ticks = [min_y_val + i * y_step for i in range(num_y_ticks)]

        for val in y_ticks:
            y_pos = (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h
            canvas.create_line(left_m, y_pos, w - right_m, y_pos, fill="#e2e8f0", dash=(2, 2))

            tick_label = f"{int(val)}" if val == int(val) else f"{val:.1f}"
            canvas.create_text(left_m - 8, y_pos, text=tick_label, fill="#64748b", font=("Segoe UI", 7), anchor="e")

        canvas.create_line(left_m, top_m, left_m, h - bottom_m, fill="#cbd5e1", width=1.5)
        canvas.create_line(left_m, h - bottom_m, w - right_m, h - bottom_m, fill="#cbd5e1", width=1.5)

        num_bars = max(len(period_data), 1)
        col_w = plot_w / num_bars
        bar_w = min(col_w * 0.65, 28)

        bar_color = "#166534"

        for idx, (label, val) in enumerate(period_data):
            center_x = left_m + (idx + 0.5) * col_w

            x1 = center_x - (bar_w / 2)
            x2 = center_x + (bar_w / 2)
            y1 = (h - bottom_m) - ((val - min_y_val) / y_range) * plot_h
            y2 = h - bottom_m

            canvas.create_rectangle(x1, y1, x2, y2, fill=bar_color, outline="")

            # Value Label
            canvas.create_text(center_x, y1 - 7, text=str(val), fill="#166534", font=("Segoe UI", 7, "bold"))

            # Category Label
            canvas.create_text(center_x, h - bottom_m + 12, text=label, fill="#64748b", font=("Segoe UI", 7))

    # ---------------------------------------------------------
    # Top Control Bar Button Methods
    # ---------------------------------------------------------
    def on_active_sheet_click(self):
        """Creates and opens the next months input sheet"""
        try:
            # 1. Locate all year directories sorted chronologically
            year_dirs = sorted([d for d in os.listdir() if d.startswith('figures 20') and os.path.isdir(d)])
            if not year_dirs:
                messagebox.showerror("Error", "No existing 'figures 20XX' folders found.")
                return

            # Scan for the absolute latest input sheet across all year folders
            input_sheets = []
            for year_dir in year_dirs:
                for file in sorted(os.listdir(year_dir)):
                    if 'inputs' in file and file[:5].replace('_', '-').replace('/', '-'):
                        try:
                            dt = datetime.strptime(file[:5], "%m-%y")
                            input_sheets.append((dt, os.path.join(year_dir, file)))
                        except ValueError:
                            continue

            if not input_sheets:
                messagebox.showerror("Error", "No previous input sheets found to base the next sheet on.")
                return

            # Get latest sheet date and determine target date
            latest_dt, latest_filepath = max(input_sheets, key=lambda x: x[0])
            next_dt = latest_dt + relativedelta(months=1)

            month_str = next_dt.strftime("%m-%y")

            # 2. Ensure target year directory exists
            target_year_dir = f"figures 20{month_str[3:]}"
            if not os.path.exists(target_year_dir):
                os.makedirs(target_year_dir)

            # 3. Read input datasets
            blank_path = 'blank inputs.xlsx'
            if not os.path.exists(blank_path):
                messagebox.showerror("Error", f"Template file '{blank_path}' not found.")
                return

            df_prev = pd.read_excel(latest_filepath)
            data2 = [df_prev.columns.tolist()] + df_prev.values.tolist()

            df_blank = pd.read_excel(blank_path)
            data = [df_blank.columns.tolist()] + df_blank.values.tolist()

            # Update header date (MM/YY format)
            data[0][15] = f"{month_str[:2]}/{month_str[3:]}"

            # Carry forward previous static values
            data[3][1] = data2[4][1]
            data[3][2] = data2[4][2]
            data[3][3] = data2[4][3]
            data[3][4] = data2[4][4]
            data[3][7] = data2[4][7]

            # Carry forward indexed columns
            for i in [4, 5, 6, 7, 10, 14, 15, 16, 17, 18, 19, 20]:
                if len(data2[i]) > 15:
                    data[i][12], data[i][13] = data2[i][14], data2[i][15]

            # Clean NaN and Unnamed pandas artifacts
            for i in range(min(26, len(data))):
                for j in range(min(16, len(data[i]))):
                    val = data[i][j]
                    if isinstance(val, float) and np.isnan(val):
                        data[i][j] = None
                    elif str(val).startswith("Unnamed:"):
                        data[i][j] = None

            # 4. Write workbook using XlsxWriter
            target_file_path = os.path.join(target_year_dir, f"{month_str} inputs.xlsx")
            workbook = xlsxwriter.Workbook(target_file_path)
            worksheet = workbook.add_worksheet()

            # Define formats
            input_fmt = workbook.add_format({'bg_color': 'white', 'border': 1})
            heads = workbook.add_format({'bg_color': 'white', 'bold': True, 'bottom': 1})
            white = workbook.add_format({'bg_color': 'white', 'bold': True})
            greyL = workbook.add_format({'bg_color': '#F2F2F2', 'bold': True, 'border': 1})
            greyM = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True})
            greyD = workbook.add_format({'bg_color': '#BFBFBF', 'bold': True, 'border': 1})
            green = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1})

            greyR = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'right': 1})
            greyB = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'bottom': 1})
            greyT = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'top': 1})
            gr_bl = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'bottom': 1, 'left': 1})
            gr_tl = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'top': 1, 'left': 1})
            gr_br = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'bottom': 1, 'right': 1})
            gr_tr = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'top': 1, 'right': 1})

            formats = [
                [heads, white, white, white, white, white, white, white, white, white, heads, white, white, white, greyL, input_fmt],
                [white, white, white, white, white, white, white, white, white, white, white, white, white, white, white, white],
                [greyD, greyL, greyL, greyL, greyL, white, greyD, greyL, white, greyL, greyD, greyD, greyD, greyD, greyD, greyD],
                [greyL, input_fmt, input_fmt, input_fmt, input_fmt, white, greyL, input_fmt, white, greyD, greyL, greyL, greyL, greyL, greyL, greyL],
                [greyL, input_fmt, input_fmt, input_fmt, input_fmt, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, input_fmt, greyM, input_fmt, input_fmt, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, input_fmt, greyM, input_fmt, input_fmt, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, input_fmt, greyM, input_fmt, input_fmt, white, greyL, input_fmt, white, gr_bl, green, green, green, green, green, green],
                [greyL, input_fmt, greyM, input_fmt, input_fmt, white, greyL, input_fmt, white, white, white, white, white, white, white, white],
                [greyL, input_fmt, input_fmt, input_fmt, input_fmt, white, greyL, input_fmt, white, greyL, white, white, white, white, white, white],
                [greyL, input_fmt, input_fmt, input_fmt, input_fmt, white, white, white, white, gr_bl, green, green, green, green, green, green],
                [greyL, input_fmt, input_fmt, input_fmt, input_fmt, white, white, white, white, white, white, white, white, white, white, white],
                [greyD, greyM, greyM, greyM, greyR, white, greyD, greyL, white, greyL, white, white, white, white, white, white],
                [greyL, greyM, input_fmt, input_fmt, greyR, white, greyL, input_fmt, white, greyL, white, white, white, white, white, white],
                [greyL, greyM, input_fmt, greyM, greyR, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, greyM, input_fmt, greyM, greyR, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, greyM, input_fmt, greyM, greyR, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyD, greyM, greyM, greyM, greyR, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, greyM, greyM, greyM, input_fmt, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, greyM, greyM, greyM, input_fmt, white, greyL, input_fmt, white, greyL, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt, input_fmt],
                [greyL, greyM, greyM, greyM, input_fmt, white, greyL, input_fmt, white, gr_bl, green, green, green, green, green, green],
                [greyL, greyM, greyM, greyM, input_fmt, white, white, white, white, white, white, white, white, white, white, white],
                [greyL, greyM, greyM, greyM, input_fmt, white, greyL, input_fmt, greyL, input_fmt, white, white, gr_tl, greyT, greyT, gr_tr],
                [greyL, greyM, greyM, greyM, input_fmt, white, greyL, input_fmt, greyL, input_fmt, white, white, greyL, input_fmt, greyM, greyR],
                [greyL, greyM, greyM, greyM, input_fmt, white, greyL, input_fmt, greyL, input_fmt, white, white, greyL, input_fmt, greyM, greyR],
                [greyL, greyB, greyB, greyB, input_fmt, white, greyL, input_fmt, greyL, input_fmt, white, white, greyL, input_fmt, greyB, gr_br]]

            # Direct 0-indexed row/col write
            for r in range(len(data)):
                for c in range(len(data[0])):
                    worksheet.write(r, c, data[r][c], formats[r][c])

            # Column widths setting
            col_widths = [35.36, 8.45, 19.91, 11.45, 11.45, 14.91, 24.82, 9.73, 17, 9.27, 18.73, 18.45, 16.27, 14.82, 15.45, 14.36]
            for col_idx, width in enumerate(col_widths):
                worksheet.set_column(col_idx, col_idx, width)

            worksheet.merge_range('K3:P3', '', greyD)

            # Write formulas
            for col in ['K', 'L', 'M', 'N', 'O', 'P']:
                worksheet.write_formula(f'{col}21', f'={col}15+{col}16+{col}17+{col}18+{col}19+{col}20', green)
                worksheet.write_formula(f'{col}8', f'={col}5+{col}6+{col}7', green)

            workbook.close()

            # 5. UI Updates & File Launch
            self._refresh_treeview()
            self.btn_active_sheet.config(text=self._get_next_input_label())

            update_dates = self._get_update_from_dates()
            self.date_sb.config(values=update_dates)
            if update_dates:
                self.date_sb.set(update_dates[-1])

            os.startfile(target_file_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate active input sheet: {str(e)}")

    def on_edit_sheet_click(self):
        """Opens a generic input sheet to edit standard figures"""
        os.startfile("blank inputs.xlsx")

    def on_delete_sheet_click(self):
        """
        Deletes the most recent input sheet only if no figures sheet exists for that month.
        Removes empty year folders and updates UI components (Treeview, Spinbox, Ribbon).
        """
        try:
            years = sorted([item for item in os.listdir() if 'figures 20' in item and os.path.isdir(item)])
            input_files = []
            fig_dates = set()

            # Scan directories for figure and input sheets
            for year in years:
                for file in sorted(os.listdir(year)):
                    filepath = os.path.join(year, file)
                    clean_name = file.replace('_', '-').replace('/', '-')
                    if len(clean_name) >= 5:
                        date_str = clean_name[:5]
                        try:
                            dt = datetime.strptime(date_str, "%m-%y")
                            if 'inputs' in file:
                                input_files.append((dt, filepath, year))
                            elif 'figures' in file:
                                fig_dates.add(dt)
                        except ValueError:
                            continue

            if not input_files:
                messagebox.showinfo("Delete Sheet", "No input sheets found to delete.")
                return

            # Identify the most recent input sheet
            latest_dt, latest_filepath, year_dir = max(input_files, key=lambda x: x[0])
            month_label = latest_dt.strftime('%m/%y')

            # Prevent deletion if a figures sheet already exists for this month
            if latest_dt in fig_dates:
                messagebox.showwarning(
                    "Action Blocked",
                    f"Cannot delete input sheet for {month_label} because a Figures sheet has already been generated for this month."
                )
                return

            # Confirm deletion with user
            confirm = messagebox.askyesno(
                "Confirm Deletion",
                f"Are you sure you want to delete the most recent input sheet ({month_label})?"
            )
            if not confirm:
                return

            # Delete the input sheet file
            if os.path.exists(latest_filepath):
                os.remove(latest_filepath)

            # Remove parent year directory if it is now empty
            if os.path.exists(year_dir) and not os.listdir(year_dir):
                os.rmdir(year_dir)

            # ---------------------------------------------------------
            # Refresh UI Elements
            # ---------------------------------------------------------
            # 1. Update Navigation Treeview
            self._refresh_treeview()

            # 2. Update Next Input Button Label
            self.btn_active_sheet.config(text=self._get_next_input_label())

            # 3. Resync Update From Spinbox Range
            update_dates = self._get_update_from_dates()
            self.date_sb.config(values=update_dates)
            if update_dates:
                self.date_sb.set(update_dates[-1])

            messagebox.showinfo("Success", f"Input sheet for {month_label} deleted successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete input sheet: {str(e)}")

    def on_run_report_click(self):
        """Processes monthly sheets from the selected spinbox date up to the latest input sheet date."""
        try:
            # 1. Locate year directories sorted chronologically
            year_dirs = sorted([d for d in os.listdir() if d.startswith('figures 20') and os.path.isdir(d)])
            if not year_dirs:
                messagebox.showerror("Error", "No 'figures 20XX' directories found.")
                return

            # Scan for all input sheets to determine the absolute latest date
            input_dates = []
            for y_dir in year_dirs:
                for file in sorted(os.listdir(y_dir)):
                    if 'inputs' in file and len(file) >= 5:
                        try:
                            dt = datetime.strptime(file[:5].replace('_', '-').replace('/', '-'), "%m-%y")
                            input_dates.append(dt)
                        except ValueError:
                            continue

            if not input_dates:
                messagebox.showerror("Error", "No valid input sheets found.")
                return

            last_dt = max(input_dates)

            # 2. Parse selected start date from the spinbox
            start_str = self.date_sb.get().replace('/', '-')
            current_dt = datetime.strptime(start_str, "%m-%y")

            # 3. Calculate total iterations required
            total_steps = (last_dt.year - current_dt.year) * 12 + (last_dt.month - current_dt.month) + 1

            if total_steps <= 0:
                return

            # Set initial progress bar bounds and value
            self.progress_bar.config(maximum=total_steps, value=0)
            self.update_idletasks()

            # 4. Execute processing loop
            for k in range(total_steps):
                month_str = current_dt.strftime("%m-%y")

                # Call external sheet generator
                Pg.pig_monitor(month_str)

                # Advance to next month
                current_dt += relativedelta(months=1)

                # Update progress bar value and refresh UI
                self.progress_bar.config(value=k + 1)
                self.update_idletasks()

            # 5. Resync Update From Spinbox Range
            update_dates = self._get_update_from_dates()
            self.date_sb.config(values=update_dates)
            if update_dates:
                self.date_sb.set(update_dates[-1])

            # 6. Rebuild navigation treeview
            self._refresh_treeview()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while running sheets: {str(e)}")
        finally:
            # Reset progress bar to full scale (100%) when execution finishes or encounters an error
            self.progress_bar.config(maximum=100, value=100)

    def on_open_excel_click(self):
        """Opens selected report in Excel"""
        os.startfile(os.path.join(f"figures 20{self.date[3:]}", f"{self.date} figures.xlsx"))


if __name__ == "__main__":
    app = PigMonitorDashboard()
    app.mainloop()
