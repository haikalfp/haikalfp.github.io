import os
import sys
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# --- Premium Theme Settings ---
ctk.set_appearance_mode("Dark")  # Force dark mode for a sleek look
ctk.set_default_color_theme("blue")

BG_COLOR = "#242424"
FG_COLOR = "#2b2b2b"
TEXT_COLOR = "#dce4ee"
BORDER_COLOR = "#444444"
ACCENT_COLOR = "#1f538d"

class HeaderLabel(ctk.CTkLabel):
    def __init__(self, master, text, **kwargs):
        super().__init__(master, text=text, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"), **kwargs)

class MetricCard(ctk.CTkFrame):
    def __init__(self, master, title, **kwargs):
        super().__init__(master, fg_color=FG_COLOR, border_width=1, border_color=BORDER_COLOR, corner_radius=10, **kwargs)
        self.title_lbl = ctk.CTkLabel(self, text=title, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"), text_color="#aaaaaa")
        self.title_lbl.pack(pady=(15, 5), padx=20)
        self.value_lbl = ctk.CTkLabel(self, text="-", font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"))
        self.value_lbl.pack(pady=(0, 5), padx=20)
        self.sub_lbl = ctk.CTkLabel(self, text="", font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#28a745")
        self.sub_lbl.pack(pady=(0, 15), padx=20)

class DrillholeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Drillhole Assay Viewer - Professional Edition")
        self.geometry("1400x900")
        
        # State
        self.df_collar = pd.DataFrame()
        self.df_assay = pd.DataFrame()
        self.sheets = {}
        self.active_data = pd.DataFrame()
        self.collar_z_val = 0.0
        
        self.chem_elements = ["Ni", "Fe", "MgO", "SiO2", "Co", "CaO", "Al2O3", "SiMa"]
        self.chem_colors = {
            "Ni": '#4dabf7', "Fe": '#ff6b6b', "MgO": '#fcc419', 
            "SiO2": '#8ce99a', "Co": '#faa2c1', "CaO": '#da77f2', 
            "Al2O3": '#c8a090', "SiMa": '#63e6be'
        }
        
        self.setup_vars()
        self.setup_ui()
        self.apply_treeview_style()
        
    def setup_vars(self):
        self.filename_var = ctk.StringVar(value="No file loaded")
        
        self.collar_sheet_var = ctk.StringVar()
        self.assay_sheet_var = ctk.StringVar()
        
        self.c_hole_id_var = ctk.StringVar(value="")
        self.c_z_var = ctk.StringVar(value="")
        
        self.a_hole_id_var = ctk.StringVar(value="")
        self.a_from_var = ctk.StringVar(value="")
        self.a_to_var = ctk.StringVar(value="")
        self.a_litho_var = ctk.StringVar(value="")
        self.a_topo_var = ctk.StringVar(value="")
        
        self.chem_vars = {el: ctk.StringVar(value="") for el in self.chem_elements}
        self.hole_id_select_var = ctk.StringVar(value="")
        
        # Plotly figure
        self.fig = plt.Figure(figsize=(6, 10), dpi=100, facecolor=BG_COLOR)
        self.canvas = None
        
    def apply_treeview_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        
        # Treeview Configuration
        style.configure("Treeview", 
                        background=FG_COLOR, 
                        foreground=TEXT_COLOR, 
                        fieldbackground=FG_COLOR, 
                        borderwidth=0, 
                        rowheight=35,
                        font=("Segoe UI", 11))
                        
        style.configure("Treeview.Heading", 
                        background="#343638", 
                        foreground=TEXT_COLOR, 
                        font=('Segoe UI', 12, 'bold'), 
                        borderwidth=0, 
                        relief="flat",
                        padding=5)
                        
        style.map('Treeview', background=[('selected', ACCENT_COLOR)])
        style.map('Treeview.Heading', background=[('active', '#3a3d3f')])
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})]) # Remove border

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # --- Sidebar ---
        self.sidebar = ctk.CTkScrollableFrame(self, width=340, corner_radius=0, fg_color=FG_COLOR)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        title_lbl = ctk.CTkLabel(self.sidebar, text="Drillhole Setup", font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"))
        title_lbl.pack(pady=(20, 20), padx=20, anchor="w")
        
        self.btn_load = ctk.CTkButton(self.sidebar, text="📁 Load Excel / CSV", command=self.load_file, 
                                      height=40, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"))
        self.btn_load.pack(fill="x", padx=20, pady=(0, 5))
        
        self.lbl_file = ctk.CTkLabel(self.sidebar, textvariable=self.filename_var, text_color="#aaaaaa", font=ctk.CTkFont(size=12))
        self.lbl_file.pack(padx=20, pady=(0, 15), anchor="w")
        
        self.create_separator(self.sidebar)
        
        # 1. Sheets
        HeaderLabel(self.sidebar, "1. Sheet Selection").pack(pady=(10, 5), padx=20, anchor="w")
        self.cb_c_sheet = ctk.CTkOptionMenu(self.sidebar, variable=self.collar_sheet_var, command=self.update_collar_cols, dynamic_resizing=False)
        self.cb_c_sheet.pack(fill="x", padx=20, pady=5)
        self.cb_a_sheet = ctk.CTkOptionMenu(self.sidebar, variable=self.assay_sheet_var, command=self.update_assay_cols, dynamic_resizing=False)
        self.cb_a_sheet.pack(fill="x", padx=20, pady=(5, 15))
        
        self.create_separator(self.sidebar)
        
        # 2. Collar Map
        HeaderLabel(self.sidebar, "2. Collar Mapping").pack(pady=(10, 5), padx=20, anchor="w")
        self.cb_c_hole = ctk.CTkOptionMenu(self.sidebar, variable=self.c_hole_id_var, dynamic_resizing=False)
        self.cb_c_hole.pack(fill="x", padx=20, pady=5)
        self.cb_c_z = ctk.CTkOptionMenu(self.sidebar, variable=self.c_z_var, dynamic_resizing=False)
        self.cb_c_z.pack(fill="x", padx=20, pady=(5, 15))
        
        self.create_separator(self.sidebar)
        
        # 3. Assay Map
        HeaderLabel(self.sidebar, "3. Assay Layout").pack(pady=(10, 5), padx=20, anchor="w")
        cb_configs = [
            (self.a_hole_id_var, "Hole ID"),
            (self.a_from_var, "Depth From"),
            (self.a_to_var, "Depth To"),
            (self.a_litho_var, "Lithology"),
            (self.a_topo_var, "Topo Position")
        ]
        
        self.assay_cbs = []
        for var, title in cb_configs:
            cb = ctk.CTkOptionMenu(self.sidebar, variable=var, dynamic_resizing=False)
            cb.pack(fill="x", padx=20, pady=5)
            self.assay_cbs.append(cb)
            
        # Unpack refs for update func
        self.cb_a_hole, self.cb_a_from, self.cb_a_to, self.cb_a_litho, self.cb_a_topo = self.assay_cbs
        
        self.create_separator(self.sidebar)
        
        # 4. Chem Map
        HeaderLabel(self.sidebar, "4. Chemistry mapping").pack(pady=(10, 5), padx=20, anchor="w")
        self.chem_widgets = {}
        for el in self.chem_elements:
            frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
            frame.pack(fill="x", padx=20, pady=2)
            ctk.CTkLabel(frame, text=el, width=50, anchor="w").pack(side="left")
            cb = ctk.CTkOptionMenu(frame, variable=self.chem_vars[el], dynamic_resizing=False)
            cb.pack(side="right", fill="x", expand=True)
            self.chem_widgets[el] = cb
            
        self.btn_apply = ctk.CTkButton(self.sidebar, text="Map Data & Update Dashboard", command=self.apply_mapping, 
                                       fg_color="#28a745", hover_color="#218838", height=45, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"))
        self.btn_apply.pack(fill="x", padx=20, pady=(30, 30))
        
        # --- Main Area ---
        self.main_panel = ctk.CTkFrame(self, corner_radius=0, fg_color=BG_COLOR)
        self.main_panel.grid(row=0, column=1, sticky="nsew")
        self.main_panel.grid_rowconfigure(1, weight=1)
        self.main_panel.grid_columnconfigure(0, weight=1)
        
        # Top Bar
        self.top_bar = ctk.CTkFrame(self.main_panel, height=70, fg_color=FG_COLOR, corner_radius=10)
        self.top_bar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(self.top_bar, text="Select Hole ID:", font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")).pack(side="left", padx=20, pady=20)
        self.cb_hole_select = ctk.CTkOptionMenu(self.top_bar, variable=self.hole_id_select_var, command=self.render_dashboard, 
                                                width=200, height=35, font=ctk.CTkFont(size=14))
        self.cb_hole_select.pack(side="left", padx=5, pady=20)
        
        # Tabview
        self.tabview = ctk.CTkTabview(self.main_panel, fg_color=BG_COLOR)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        
        self.tab_data = self.tabview.add("📊 Data Viewer")
        self.tab_sum = self.tabview.add("📉 Summary Statistics")
        self.tab_ore = self.tabview.add("⛏️ Ore Calc")
        self.tab_diag = self.tabview.add("📈 Assay Diagram")
        
        for tab in [self.tab_data, self.tab_sum, self.tab_ore, self.tab_diag]:
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            
        self.setup_data_viewer_ui()
        self.setup_summary_ui()
        self.setup_ore_calc_ui()
        self.setup_diagram_ui()

    def create_separator(self, parent):
        sep = ctk.CTkFrame(parent, height=2, fg_color=BORDER_COLOR)
        sep.pack(fill="x", padx=20, pady=10)
        
    def setup_data_viewer_ui(self):
        container = ctk.CTkFrame(self.tab_data, fg_color=FG_COLOR, corner_radius=10)
        container.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        self.tree_scroll_y = ttk.Scrollbar(container)
        self.tree_scroll_y.grid(row=0, column=1, sticky='ns', pady=1)
        self.tree_scroll_x = ttk.Scrollbar(container, orient="horizontal")
        self.tree_scroll_x.grid(row=1, column=0, sticky='ew', padx=1)
        
        self.tree = ttk.Treeview(container, show="headings", yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.grid(row=0, column=0, sticky='nsew', padx=1, pady=1)
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)
        
        # Conditional formatting tags (Pastel/Dark mode friendly)
        self.tree.tag_configure("ni_low", background="#3e2723", foreground="#ffcdd2")     # Dark red
        self.tree.tag_configure("ni_med", background="#e65100", foreground="#fff3e0")     # Dark orange
        self.tree.tag_configure("ni_high", background="#1b5e20", foreground="#c8e6c9")    # Dark green
        self.tree.tag_configure("striped", background="#2b2d30")

    def setup_summary_ui(self):
        self.tab_sum.grid_rowconfigure(1, weight=1)
        
        top_frame = ctk.CTkFrame(self.tab_sum, fg_color="transparent")
        top_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.sum_avail_cb = ctk.CTkCheckBox(top_frame, text="Available Materials Only (Below Topo)", command=self.render_summary)
        self.sum_avail_cb.pack(side="left", padx=15, pady=10)
        
        self.sum_container = ctk.CTkScrollableFrame(self.tab_sum, fg_color=FG_COLOR, corner_radius=10)
        self.sum_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

    def setup_ore_calc_ui(self):
        self.tab_ore.grid_rowconfigure(1, weight=1)
        
        top_frame = ctk.CTkFrame(self.tab_ore, fg_color="transparent")
        top_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.ore_avail_cb = ctk.CTkCheckBox(top_frame, text="Available Materials Only (Below Topo)", command=self.render_ore_calc)
        self.ore_avail_cb.pack(side="left", padx=15, pady=10)
        
        self.ore_results_frame = ctk.CTkFrame(self.tab_ore, fg_color="transparent")
        self.ore_results_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        self.ore_results_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        self.ore_cards = {}
        fields = [
            ("Total Thickness", 0, 0),
            ("Total Ore (Ni >= 1.0)", 0, 1),
            ("Non-Ore (Ni < 1.0)", 0, 2),
            ("Overburden (OB)", 1, 0),
            ("Stripping Ratio (SR)", 1, 1)
        ]
        
        for name, r, c in fields:
            card = MetricCard(self.ore_results_frame, name)
            card.grid(row=r, column=c, padx=10, pady=10, sticky="nsew")
            self.ore_cards[name] = card

    def setup_diagram_ui(self):
        self.tab_diag.grid_rowconfigure(1, weight=1)
        
        ctrl_frame = ctk.CTkFrame(self.tab_diag, fg_color=FG_COLOR, corner_radius=10)
        ctrl_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        inner_ctrl = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        inner_ctrl.pack(padx=15, pady=15, fill="x")
        
        ctk.CTkLabel(inner_ctrl, text="Ni Max:").pack(side="left", padx=(0, 5))
        self.entry_ni_max = ctk.CTkEntry(inner_ctrl, width=60)
        self.entry_ni_max.pack(side="left", padx=5)
        self.entry_ni_max.insert(0, "3.5")
        
        ctk.CTkLabel(inner_ctrl, text="Others Max:").pack(side="left", padx=(20, 5))
        self.entry_oth_max = ctk.CTkEntry(inner_ctrl, width=60)
        self.entry_oth_max.pack(side="left", padx=5)
        self.entry_oth_max.insert(0, "70.0")
        
        ctk.CTkButton(inner_ctrl, text="Update Chart", width=120, command=self.render_diagram).pack(side="left", padx=20)
        
        # El toggles
        toggles_frame = ctk.CTkFrame(inner_ctrl, fg_color="transparent")
        toggles_frame.pack(side="right")
        self.el_vars = {el: ctk.BooleanVar(value=el in ["Ni", "Fe", "MgO", "SiO2"]) for el in self.chem_elements}
        for el in self.chem_elements:
            cb = ctk.CTkCheckBox(toggles_frame, text=el, variable=self.el_vars[el], command=self.render_diagram, width=60)
            cb.pack(side="left", padx=5)
            
        self.chart_frame = ctk.CTkFrame(self.tab_diag, fg_color=FG_COLOR, corner_radius=10)
        self.chart_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

    def guess_index(self, options, target, exact=False):
        tgt = target.lower()
        if not options: return 0
        for i, opt in enumerate(options):
            if opt.lower() == tgt: return i + 1
        if not exact:
            for i, opt in enumerate(options):
                if tgt in opt.lower() or opt.lower() in tgt:
                    return i + 1
        return 0

    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm"), ("CSV files", "*.csv")])
        if not filepath: return
        
        self.filename_var.set(os.path.basename(filepath))
        
        try:
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath)
                self.sheets = {"Data": df}
            else:
                xls = pd.ExcelFile(filepath)
                self.sheets = {sht: xls.parse(sht) for sht in xls.sheet_names}
                
            sheet_names = list(self.sheets.keys())
            self.cb_c_sheet.configure(values=sheet_names)
            self.cb_a_sheet.configure(values=sheet_names)
            
            c_idx = next((i for i, s in enumerate(sheet_names) if 'collar' in s.lower()), 0)
            a_idx = next((i for i, s in enumerate(sheet_names) if 'assay' in s.lower()), min(1, len(sheet_names)-1) if len(sheet_names)>1 else 0)
            
            self.collar_sheet_var.set(sheet_names[c_idx])
            self.assay_sheet_var.set(sheet_names[a_idx])
            
            self.update_collar_cols(sheet_names[c_idx])
            self.update_assay_cols(sheet_names[a_idx])
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file:\n{e}")

    def update_collar_cols(self, sheet_name):
        if not sheet_name or sheet_name not in self.sheets: return
        cols = [""] + list(self.sheets[sheet_name].columns)
        str_cols = [str(c) for c in cols]
        
        self.cb_c_hole.configure(values=str_cols)
        self.cb_c_z.configure(values=str_cols)
        
        self.c_hole_id_var.set(str_cols[self.guess_index(cols[1:], "hole")])
        self.c_z_var.set(str_cols[self.guess_index(cols[1:], "z", True) or self.guess_index(cols[1:], "elev")])

    def update_assay_cols(self, sheet_name):
        if not sheet_name or sheet_name not in self.sheets: return
        cols = [""] + list(self.sheets[sheet_name].columns)
        str_cols = [str(c) for c in cols]
        
        for cb in self.assay_cbs: cb.configure(values=str_cols)
        
        self.a_hole_id_var.set(str_cols[self.guess_index(cols[1:], "hole")])
        self.a_from_var.set(str_cols[self.guess_index(cols[1:], "from")])
        self.a_to_var.set(str_cols[self.guess_index(cols[1:], "to", True) or self.guess_index(cols[1:], "depth")])
        self.a_litho_var.set(str_cols[self.guess_index(cols[1:], "zonasi") or self.guess_index(cols[1:], "litho")])
        self.a_topo_var.set(str_cols[self.guess_index(cols[1:], "topo")])
        
        for el in self.chem_elements:
            self.chem_widgets[el].configure(values=str_cols)
            self.chem_vars[el].set(str_cols[self.guess_index(cols[1:], el, True)])

    def apply_mapping(self):
        c_sht = self.collar_sheet_var.get()
        a_sht = self.assay_sheet_var.get()
        if not c_sht or not a_sht: return
        
        self.df_collar = self.sheets[c_sht]
        self.df_assay = self.sheets[a_sht]
        
        c_hole = self.c_hole_id_var.get()
        if not c_hole:
            messagebox.showwarning("Warning", "Collar Hole ID is not mapped!")
            return
            
        holes = sorted([str(x) for x in self.df_collar[c_hole].dropna().unique() if str(x).strip()])
        self.cb_hole_select.configure(values=holes)
        if holes:
            self.hole_id_select_var.set(holes[0])
            self.process_hole_data(holes[0])

    def process_hole_data(self, hole_id):
        if not hole_id: return
        
        c_hole = self.c_hole_id_var.get()
        c_z = self.c_z_var.get()
        
        c_row = self.df_collar[self.df_collar[c_hole].astype(str) == hole_id]
        self.collar_z_val = 0.0
        if c_z and not c_row.empty:
            try:
                self.collar_z_val = pd.to_numeric(c_row[c_z]).values[0]
            except: pass
            
        a_hole = self.a_hole_id_var.get()
        if a_hole in self.df_assay.columns:
            self.active_data = self.df_assay[self.df_assay[a_hole].astype(str) == hole_id].copy()
            a_from = self.a_from_var.get()
            if a_from in self.active_data.columns:
                self.active_data[a_from] = pd.to_numeric(self.active_data[a_from], errors='coerce')
                self.active_data = self.active_data.sort_values(by=a_from)
                self.active_data['Elevation'] = self.collar_z_val - self.active_data[a_from]
                
        self.render_dashboard()
        
    def render_dashboard(self, *args):
        if self.active_data.empty: return
        self.render_data_viewer()
        self.render_summary()
        self.render_ore_calc()
        self.render_diagram()

    def render_data_viewer(self):
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        a_from = self.a_from_var.get()
        a_to = self.a_to_var.get()
        a_litho = self.a_litho_var.get()
        a_topo = self.a_topo_var.get()
        
        cols = []
        if a_from: cols.append((a_from, "From"))
        if a_to: cols.append((a_to, "To"))
        cols.append(('Elevation', "Elev"))
        if a_litho: cols.append((a_litho, "Lithology"))
        if a_topo: cols.append((a_topo, "Topo Position"))
        
        chem_mapped = []
        for el in self.chem_elements:
            col = self.chem_vars[el].get()
            if col and col in self.active_data.columns:
                cols.append((col, el))
                chem_mapped.append(col)
                
        self.tree["columns"] = [d[1] for d in cols]
        for c in self.tree["columns"]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100, anchor='center')
            
        ni_col = self.chem_vars["Ni"].get()
        
        for idx, row in self.active_data.iterrows():
            vals = []
            for og_col, disp_col in cols:
                v = row.get(og_col, "")
                if pd.api.types.is_number(v) and pd.notna(v):
                    vals.append(f"{v:.2f}")
                else:
                    vals.append(str(v) if pd.notna(v) else "")
                    
            tag = "striped" if idx % 2 == 0 else ""
            if ni_col and ni_col in row:
                try:
                    ni_val = float(row[ni_col])
                    if pd.notna(ni_val):
                        if ni_val < 1.0: tag = "ni_low"
                        elif ni_val <= 1.3: tag = "ni_med"
                        else: tag = "ni_high"
                except: pass
                
            self.tree.insert("", tk.END, values=vals, tags=(tag,))

    def render_summary(self):
        for widget in self.sum_container.winfo_children():
            widget.destroy()
            
        a_topo = self.a_topo_var.get()
        target = self.active_data.copy()
        
        if self.sum_avail_cb.get() and a_topo and a_topo in target.columns:
            target = target[target[a_topo].astype(str).str.lower().str.contains("below", na=False)]
            
        if target.empty:
            ctk.CTkLabel(self.sum_container, text="No data matches criteria.", text_color="#aaaaaa").pack(pady=20)
            return
            
        # Draw header
        header_frame = ctk.CTkFrame(self.sum_container, fg_color="transparent")
        header_frame.pack(fill="x", pady=10)
        
        columns = ["Element", "Min", "Max", "Average", "Variance", "Std Dev"]
        for i, col in enumerate(columns):
            ctk.CTkLabel(header_frame, text=col, font=ctk.CTkFont(weight="bold", size=14)).grid(row=0, column=i, sticky="w", padx=20)
            header_frame.grid_columnconfigure(i, weight=1)
            
        self.create_separator(self.sum_container)
            
        row_idx = 1
        for el in self.chem_elements:
            col = self.chem_vars[el].get()
            if col and col in target.columns:
                vals = pd.to_numeric(target[col], errors='coerce').dropna()
                if not vals.empty:
                    df_frame = ctk.CTkFrame(self.sum_container, fg_color="transparent")
                    df_frame.pack(fill="x", pady=5)
                    
                    data = [el, f"{vals.min():.2f}", f"{vals.max():.2f}", f"{vals.mean():.2f}", f"{vals.var():.2f}", f"{vals.std():.2f}"]
                    for i, d in enumerate(data):
                        ctk.CTkLabel(df_frame, text=d, font=ctk.CTkFont(size=14, weight="bold" if i==0 else "normal")).grid(row=0, column=i, sticky="w", padx=20)
                        df_frame.grid_columnconfigure(i, weight=1)
                    row_idx += 1

    def render_ore_calc(self):
        a_from = self.a_from_var.get()
        a_to = self.a_to_var.get()
        a_litho = self.a_litho_var.get()
        a_topo = self.a_topo_var.get()
        ni_col = self.chem_vars["Ni"].get()
        
        target = self.active_data.copy()
        
        if self.ore_avail_cb.get() and a_topo and a_topo in target.columns:
            target = target[target[a_topo].astype(str).str.lower().str.contains("below", na=False)]
            
        for k in self.ore_cards: 
            self.ore_cards[k].value_lbl.configure(text="-")
            self.ore_cards[k].sub_lbl.configure(text="")
        
        if target.empty or not ni_col or ni_col not in target.columns or not a_from or not a_to:
            return
            
        target['Thickness'] = pd.to_numeric(target[a_to], errors='coerce') - pd.to_numeric(target[a_from], errors='coerce')
        target['Thickness'] = target['Thickness'].clip(lower=0)
        target['Ni_val'] = pd.to_numeric(target[ni_col], errors='coerce')
        
        total = target['Thickness'].sum()
        
        ore_data = target[target['Ni_val'] >= 1.0].copy()
        ore_thick = ore_data['Thickness'].sum()
        ore_avg_ni = ore_data['Ni_val'].mean() if not ore_data.empty else 0.0
        
        non_ore_data = target[target['Ni_val'] < 1.0].copy()
        non_ore_thick = non_ore_data['Thickness'].sum()
        
        ob_thick = non_ore_thick
        if a_litho and a_litho in non_ore_data.columns:
            ob_data = non_ore_data[~non_ore_data[a_litho].astype(str).str.upper().str.contains('BRK', na=False)]
            ob_thick = ob_data['Thickness'].sum()
            
        sr = ob_thick / ore_thick if ore_thick > 0 else 0
        
        self.ore_cards["Total Thickness"].value_lbl.configure(text=f"{total:.2f} m")
        
        self.ore_cards["Total Ore (Ni >= 1.0)"].value_lbl.configure(text=f"{ore_thick:.2f} m", text_color="#8ce99a")
        if ore_thick > 0:
            self.ore_cards["Total Ore (Ni >= 1.0)"].sub_lbl.configure(text=f"Avg Ni: {ore_avg_ni:.2f}%")
            
        self.ore_cards["Non-Ore (Ni < 1.0)"].value_lbl.configure(text=f"{non_ore_thick:.2f} m", text_color="#ff8787")
        self.ore_cards["Overburden (OB)"].value_lbl.configure(text=f"{ob_thick:.2f} m")
        self.ore_cards["Stripping Ratio (SR)"].value_lbl.configure(text=f"{sr:.2f} : 1")
        
    def render_diagram(self):
        a_to = self.a_to_var.get()
        if not a_to or a_to not in self.active_data.columns: return
        
        if self.canvas:
            self.canvas.get_tk_widget().destroy()
            self.fig.clear()
            
        try:
            n_max = float(self.entry_ni_max.get())
            o_max = float(self.entry_oth_max.get())
        except:
            n_max, o_max = 3.5, 70.0
            
        self.fig.patch.set_facecolor(FG_COLOR)
        ax_ni = self.fig.add_subplot(111)
        ax_ni.set_facecolor(FG_COLOR)
        
        # Style spines
        for spine in ax_ni.spines.values():
            spine.set_edgecolor(BORDER_COLOR)
            spine.set_linewidth(1)
            
        ax_oth = ax_ni.twiny()
        ax_oth.set_facecolor(FG_COLOR)
        for spine in ax_oth.spines.values():
            spine.set_edgecolor(BORDER_COLOR)
        
        y_vals = pd.to_numeric(self.active_data[a_to], errors='coerce')
        
        for el in self.chem_elements:
            if not self.el_vars[el].get(): continue
            col = self.chem_vars[el].get()
            if col and col in self.active_data.columns:
                x_vals = pd.to_numeric(self.active_data[col], errors='coerce')
                color = self.chem_colors[el]
                
                if el == "Ni":
                    ax_ni.plot(x_vals, y_vals, marker='o', color=color, label=el, linewidth=2.5, markersize=5)
                else:
                    ax_oth.plot(x_vals, y_vals, marker='s' if el in ['Fe', 'Al2O3'] else '^', color=color, label=el, linewidth=1.5, markersize=4, alpha=0.9)

        ax_ni.invert_yaxis()
        ax_ni.set_ylabel("Depth (m)", color=TEXT_COLOR, weight="bold")
        ax_ni.tick_params(axis='y', colors=TEXT_COLOR)
        
        ax_ni.set_xlabel("Ni %", color=self.chem_colors["Ni"], weight="bold")
        ax_ni.tick_params(axis='x', colors=self.chem_colors["Ni"])
        ax_ni.set_xlim(0, n_max)
        
        ax_oth.set_xlabel("Fe, MgO, SiO2, etc %", color=self.chem_colors["Fe"], weight="bold")
        ax_oth.tick_params(axis='x', colors=self.chem_colors["Fe"])
        ax_oth.set_xlim(0, o_max)
        ax_oth.grid(True, linestyle='--', color=BORDER_COLOR, alpha=0.6)
        
        lines1, labels1 = ax_ni.get_legend_handles_labels()
        lines2, labels2 = ax_oth.get_legend_handles_labels()
        
        legend = ax_ni.legend(lines1 + lines2, labels1 + labels2, loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=4, frameon=False, labelcolor=TEXT_COLOR)
        
        self.fig.tight_layout()
        
        # Embed in Tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
        self.canvas.draw()
        w = self.canvas.get_tk_widget()
        w.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

if __name__ == "__main__":
    app = DrillholeApp()
    app.mainloop()
