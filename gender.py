import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import re
from gender_guesser.detector import Detector
import pandas as pd
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)

class EmployeeGenderClassifier:
    _translations = {
        "title": "Employee Gender Classifier",
        "control_frame_title": "Import / Export Actions",
        "filter_frame": "Filter By Gender",
        "search_frame": "Search",
        "import": "Import File",
        "detect": "Detect Gender",
        "export_males": "Export Males",
        "export_females": "Export Females",
        "export_unknown": "Export Unknown",
        "export_all": "Export All",
        "stats": "Show Stats",
        "clear": "Clear Data",
        "all": "All",
        "male": "Male",
        "female": "Female",
        "unknown": "Unknown",
        "search": "Search",
        "ready": "Ready",
        "file_types": [("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv"), ("All files", "*.*")],
        "import_title": "Select File",
        "import_error": "Import Error",
        "import_success": "Loaded {} records from {}",
        "import_no_name": "No suitable name column found (e.g., 'first name', 'name'). Please ensure one exists.",
        "detect_no_data": "No data loaded",
        "detect_no_name": "Name column not identified. Please import data first.",
        "detect_complete": "Gender detection complete. {} records processed.",
        "export_no_data": "No processed data to export.",
        "export_none_found": "No {} employees found.",
        "export_title": "Save {} Employees As",
        "export_success": "Saved {} {} employees to {}",
        "export_error": "Export Error",
        "export_all_title": "Save All Data with Gender As",
        "export_all_success": "Saved all {} records with gender to {}",
        "stats_no_data": "No processed data to analyze.",
        "stats_title": "Gender Statistics",
        "stats_message": "Gender Distribution:\n\n{}\n\nTotal Records: {}",
        "stats_line": "{}: {} ({:.1f}%)",
        "clear_confirm_title": "Confirm Clear",
        "clear_confirm_msg": "Are you sure you want to clear all loaded and processed data?"
    }

    def __init__(self, root):
        self.bg_color = "#f5f5f5"
        self.primary_color = "#4a6fa5"
        self.secondary_color = "#6c757d"
        self.text_color = "#333333"
        self.highlight_color = "#e9ecef"
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".",
                        background=self.bg_color,
                        foreground=self.text_color,
                        font=('Segoe UI', 9))
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabelframe",
                        background=self.bg_color,
                        borderwidth=1,
                        relief="solid",
                        labelmargins=5)
        style.configure("TLabelframe.Label",
                        background=self.bg_color,
                        foreground=self.primary_color,
                        font=('Segoe UI', 9, 'bold'))
        style.configure("TButton",
                        background=self.primary_color,
                        foreground="white",
                        borderwidth=0,
                        relief="flat",
                        padding=8,
                        font=('Segoe UI', 9, 'bold'),
                        borderradius=4)
        style.map("TButton",
                  background=[('active', self.secondary_color), ('disabled', '#cccccc')],
                  foreground=[('disabled', '#999999')])
        style.configure("TRadiobutton",
                        background=self.bg_color,
                        font=('Segoe UI', 9))
        style.map("TRadiobutton",
                  background=[('active', self.bg_color)],
                  foreground=[('active', self.text_color)])
        style.configure("TEntry",
                        fieldbackground="white",
                        borderwidth=1,
                        relief="solid",
                        padding=5,
                        borderradius=4)
        style.configure("Treeview",
                        background="white",
                        fieldbackground="white",
                        foreground=self.text_color,
                        borderwidth=0,
                        font=('Segoe UI', 9))
        style.configure("Treeview.Heading",
                        background=self.primary_color,
                        foreground="white",
                        padding=5,
                        font=('Segoe UI', 9, 'bold'))
        style.map("Treeview.Heading",
                  background=[('active', self.secondary_color)])
        style.configure("TLabel",
                        background=self.secondary_color,
                        foreground="white",
                        padding=5,
                        font=('Segoe UI', 8))
        style.configure("Vertical.TScrollbar",
                        background=self.secondary_color,
                        troughcolor=self.bg_color,
                        bordercolor=self.secondary_color,
                        arrowcolor="white")
        style.configure("Horizontal.TScrollbar",
                        background=self.secondary_color,
                        troughcolor=self.bg_color,
                        bordercolor=self.secondary_color,
                        arrowcolor="white")
        self.root = root
        self.root.configure(background=self.bg_color)
        self.root.title(self._translations["title"])
        self.root.geometry("1200x700")

        self.name_columns = ["first name", "name"]
        self.gender_detector = Detector()

        self.original_data = pd.DataFrame()
        self.processed_data = pd.DataFrame()
        self.name_column = None
        self.current_filter = "All"

        self._create_widgets()
        self._setup_ui_text()
        self._configure_grid()

    def _configure_grid(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        self.top_frame.columnconfigure(0, weight=1)
        self.top_frame.columnconfigure(1, weight=1)
        self.top_frame.columnconfigure(2, weight=1)

        self.control_frame.columnconfigure((0, 1, 2), weight=1)
        self.filter_frame.columnconfigure((0, 1, 2, 3), weight=1)
        self.search_frame.columnconfigure(1, weight=1)

        self.display_frame.columnconfigure(0, weight=1)
        self.display_frame.rowconfigure(0, weight=1)

    def _create_widgets(self):
        self.top_frame = ttk.Frame(self.root, padding="10")
        self.top_frame.grid(row=0, column=0, sticky="ew")

        self.control_frame = ttk.LabelFrame(self.top_frame, padding="10")
        self.control_frame.grid(row=0, column=0, padx=(0, 5), pady=5, sticky="nsew")

        self.filter_frame = ttk.LabelFrame(self.top_frame, padding="10")
        self.filter_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.search_frame = ttk.LabelFrame(self.top_frame, padding="10")
        self.search_frame.grid(row=0, column=2, padx=(5, 0), pady=5, sticky="nsew")

        self.display_frame = ttk.Frame(self.root, padding="10 0 10 10")
        self.display_frame.grid(row=1, column=0, sticky="nsew")

        self.btn_import = ttk.Button(self.control_frame, command=self.import_file)
        self.btn_detect = ttk.Button(self.control_frame, command=self.detect_gender_from_data, state=tk.DISABLED)
        self.btn_export_males = ttk.Button(self.control_frame, command=lambda: self.export_by_gender("Male"), state=tk.DISABLED)
        self.btn_export_females = ttk.Button(self.control_frame, command=lambda: self.export_by_gender("Female"), state=tk.DISABLED)
        self.btn_export_unknown = ttk.Button(self.control_frame, command=lambda: self.export_by_gender("Unknown"), state=tk.DISABLED)
        self.btn_export_all = ttk.Button(self.control_frame, command=self.export_all_with_gender, state=tk.DISABLED)
        self.btn_stats = ttk.Button(self.control_frame, command=self.show_stats, state=tk.DISABLED)
        self.btn_clear = ttk.Button(self.control_frame, command=self.clear_data, state=tk.DISABLED)

        self.btn_import.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.btn_detect.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.btn_export_males.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.btn_export_females.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.btn_export_unknown.grid(row=1, column=2, padx=5, pady=5, sticky="ew")
        self.btn_export_all.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
        self.btn_stats.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.btn_clear.grid(row=2, column=2, padx=5, pady=5, sticky="ew")

        self.filter_var = tk.StringVar(value="All")
        self.radio_all = ttk.Radiobutton(self.filter_frame, value="All", variable=self.filter_var, command=self.apply_filter, state=tk.DISABLED)
        self.radio_male = ttk.Radiobutton(self.filter_frame, value="Male", variable=self.filter_var, command=self.apply_filter, state=tk.DISABLED)
        self.radio_female = ttk.Radiobutton(self.filter_frame, value="Female", variable=self.filter_var, command=self.apply_filter, state=tk.DISABLED)
        self.radio_unknown = ttk.Radiobutton(self.filter_frame, value="Unknown", variable=self.filter_var, command=self.apply_filter, state=tk.DISABLED)

        self.radio_all.grid(row=0, column=0, padx=5, sticky='w')
        self.radio_male.grid(row=0, column=1, padx=5, sticky='w')
        self.radio_female.grid(row=0, column=2, padx=5, sticky='w')
        self.radio_unknown.grid(row=0, column=3, padx=5, sticky='w')

        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_var, width=25, state=tk.DISABLED)
        self.search_entry.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        self.btn_search = ttk.Button(self.search_frame, command=self.apply_search, state=tk.DISABLED)
        self.search_entry.bind("<Return>", lambda event: self.apply_search())
        self.btn_search.grid(row=0, column=2, padx=5, pady=5)

        self.tree = ttk.Treeview(self.display_frame, show="headings", selectmode="extended", height=20)

        vsb = ttk.Scrollbar(self.display_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.display_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", padding="5")
        self.status_bar.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 5))

    def _setup_ui_text(self):
        self.control_frame.config(text=self._translations["control_frame_title"])
        self.filter_frame.config(text=self._translations["filter_frame"])
        self.search_frame.config(text=self._translations["search_frame"])

        self.btn_import.config(text=self._translations["import"])
        self.btn_detect.config(text=self._translations["detect"])
        self.btn_export_males.config(text=self._translations["export_males"])
        self.btn_export_females.config(text=self._translations["export_females"])
        self.btn_export_unknown.config(text=self._translations["export_unknown"])
        self.btn_export_all.config(text=self._translations["export_all"])
        self.btn_stats.config(text=self._translations["stats"])
        self.btn_clear.config(text=self._translations["clear"])

        self.radio_all.config(text=self._translations["all"])
        self.radio_male.config(text=self._translations["male"])
        self.radio_female.config(text=self._translations["female"])
        self.radio_unknown.config(text=self._translations["unknown"])

        self.btn_search.config(text=self._translations["search"])
        self.status_var.set(self._translations["ready"])

    def _update_widget_states(self):
        has_original_data = not self.original_data.empty
        has_processed_data = not self.processed_data.empty

        self.btn_detect.config(state=tk.NORMAL if has_original_data and self.name_column else tk.DISABLED)
        self.btn_clear.config(state=tk.NORMAL if has_original_data or has_processed_data else tk.DISABLED)

        export_state = tk.NORMAL if has_processed_data else tk.DISABLED
        self.btn_export_males.config(state=export_state)
        self.btn_export_females.config(state=export_state)
        self.btn_export_unknown.config(state=export_state)
        self.btn_export_all.config(state=export_state)
        self.btn_stats.config(state=export_state)

        filter_state = tk.NORMAL if has_processed_data else tk.DISABLED
        self.radio_all.config(state=filter_state)
        self.radio_male.config(state=filter_state)
        self.radio_female.config(state=filter_state)
        self.radio_unknown.config(state=filter_state)

        search_state = tk.NORMAL if has_processed_data else tk.DISABLED
        self.search_entry.config(state=search_state)
        self.btn_search.config(state=search_state)

    def import_file(self):
        file_path = filedialog.askopenfilename(
            title=self._translations["import_title"],
            filetypes=self._translations["file_types"])

        if not file_path:
            return

        try:
            if file_path.lower().endswith('.csv'):
                self.original_data = pd.read_csv(file_path)
            else:
                self.original_data = pd.read_excel(file_path)

            self.name_column = self._detect_name_column()
            if not self.name_column:
                messagebox.showerror(self._translations["import_error"], self._translations["import_no_name"])
                self.original_data = pd.DataFrame()
                return

            self.processed_data = pd.DataFrame()
            self.current_filter = "All"
            self.filter_var.set("All")
            self.search_var.set("")
            self.update_display(self.original_data)
            self.status_var.set(self._translations["import_success"].format(len(self.original_data), os.path.basename(file_path)))
            self._update_widget_states()

        except Exception as e:
            messagebox.showerror(self._translations["import_error"], f"Failed to import file:\n{str(e)}")
            self.original_data = pd.DataFrame()
            self.processed_data = pd.DataFrame()
            self.update_display(self.original_data)
            self.status_var.set(self._translations["import_error"])
            self._update_widget_states()

    def _detect_name_column(self):
        if self.original_data.empty:
            return None

        for col in self.original_data.columns:
            col_lower = str(col).strip().lower()
            for name_var in self.name_columns:
                if name_var == col_lower:
                    return col
        for col in self.original_data.columns:
            col_lower = str(col).strip().lower()
            for name_var in self.name_columns:
                if name_var in col_lower:
                    return col
        return None

    def detect_gender_from_data(self):
        if self.original_data.empty:
            messagebox.showwarning(self._translations["detect_no_data"], self._translations["detect_no_data"])
            return

        if not self.name_column:
            messagebox.showerror(self._translations["detect_no_name"], self._translations["detect_no_name"])
            return

        self.processed_data = self.original_data.copy()

        self.processed_data["Gender"] = self.processed_data[self.name_column].apply(
            lambda x: self._detect_gender(str(x)) if pd.notna(x) else "Unknown")

        self.current_filter = "All"
        self.filter_var.set("All")
        self.search_var.set("")
        self.update_display(self.processed_data)
        self.status_var.set(self._translations["detect_complete"].format(len(self.processed_data)))
        self._update_widget_states()

    def _detect_gender(self, name):
        try:
            name = re.sub(r'^\w+\.\s*', '', name).strip()
            first_name = name.split()[0] if name else ""
            if not first_name:
                return "Unknown"

            gender = self.gender_detector.get_gender(first_name)

            gender_map = {
                "male": "Male",
                "mostly_male": "Male",
                "female": "Female",
                "mostly_female": "Female",
                "andy": "Unknown",
                "unknown": "Unknown"
            }

            return gender_map.get(gender, "Unknown")
        except Exception as e:
            print(f"Gender detection error for '{name}': {e}")
            return "Unknown"

    def _get_current_display_data(self):
        return self.processed_data if not self.processed_data.empty else self.original_data

    def apply_filter(self):
        if self.processed_data.empty:
            self.update_display(self.processed_data)
            return

        self.current_filter = self.filter_var.get()
        self.apply_search()

    def apply_search(self):
        base_data = self._get_current_display_data()
        if base_data.empty:
            self.update_display(base_data)
            self.update_status(0)
            return

        if not self.processed_data.empty and self.current_filter != "All":
            filtered_data = self.processed_data[
                self.processed_data["Gender"] == self.current_filter
            ]
        else:
            filtered_data = base_data.copy()

        search_term = self.search_var.get().strip().lower()

        if search_term:
            mask = pd.Series(False, index=filtered_data.index)
            for col in filtered_data.columns:
                mask = mask | filtered_data[col].astype(str).str.lower().str.contains(search_term, na=False)
            search_result_data = filtered_data[mask]
        else:
            search_result_data = filtered_data

        self.update_display(search_result_data)
        self.update_status(len(search_result_data))

    def export_by_gender(self, gender):
        if self.processed_data.empty:
            messagebox.showwarning(self._translations["export_no_data"], self._translations["export_no_data"])
            return

        gender_data = self.processed_data[self.processed_data["Gender"] == gender]

        if gender_data.empty:
            messagebox.showwarning(
                self._translations["export_no_data"],
                self._translations["export_none_found"].format(gender)
            )
            return

        search_term = self.search_var.get().strip().lower()
        if search_term:
            mask = pd.Series(False, index=gender_data.index)
            for col in gender_data.columns:
                mask = mask | gender_data[col].astype(str).str.lower().str.contains(search_term, na=False)
            gender_data = gender_data[mask]

        if gender_data.empty:
            messagebox.showwarning(
                self._translations["export_no_data"],
                f"No {gender} employees found matching the current search."
            )
            return

        self._export_data(
            data_to_export=gender_data,
            title=self._translations["export_title"].format(gender),
            initial_filename=f"employees_{gender.lower()}.xlsx",
            success_message_template=f"Saved {{}} {gender} employees to {{}}"
        )

    def _export_data(self, data_to_export, title, initial_filename, success_message_template):
        if data_to_export.empty:
            source_df = self._get_current_display_data()
            if source_df.empty:
                messagebox.showwarning(self._translations["export_no_data"], "No data loaded or processed to export.")
            else:
                messagebox.showwarning(self._translations["export_no_data"], "No data matching the current filter/search to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            title=title,
            initialfile=initial_filename)

        if not file_path:
            return

        try:
            if file_path.lower().endswith('.csv'):
                data_to_export.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                try:
                    data_to_export.to_excel(file_path, index=False, engine='openpyxl')
                except ImportError:
                    messagebox.showerror(self._translations["export_error"], "Exporting to Excel requires the 'openpyxl' library.\nPlease install it (pip install openpyxl) and try again.")
                    return

            if os.path.exists(file_path):
                success_msg = success_message_template.format(len(data_to_export), os.path.basename(file_path))
                success_title = " ".join(self._translations["export_success"].split(" ")[0:2])
                messagebox.showinfo(success_title, success_msg)
                self.status_var.set(success_msg)
            else:
                messagebox.showerror(self._translations["export_error"], f"Failed to create file at:\n{file_path}")

        except Exception as e:
            messagebox.showerror(self._translations["export_error"], f"Failed to export data:\n{str(e)}")

    def export_all_with_gender(self):
        base_data = self.processed_data
        if base_data.empty:
            messagebox.showwarning(self._translations["export_no_data"], "No processed data available to export.")
            return

        if self.current_filter != "All":
            filtered_data = base_data[base_data["Gender"] == self.current_filter]
        else:
            filtered_data = base_data.copy()

        search_term = self.search_var.get().strip().lower()
        if search_term:
            mask = pd.Series(False, index=filtered_data.index)
            for col in filtered_data.columns:
                mask = mask | filtered_data[col].astype(str).str.lower().str.contains(search_term, na=False)
            data_to_export = filtered_data[mask]
        else:
            data_to_export = filtered_data

        if data_to_export.empty:
            messagebox.showwarning(self._translations["export_no_data"], "No data matching the current filter/search to export.")
            return

        self._export_data(
            data_to_export=data_to_export,
            title=self._translations["export_all_title"],
            initial_filename="employees_filtered_gender.xlsx",
            success_message_template=self._translations["export_all_success"]
        )

    def update_display(self, df):
        self.tree.delete(*self.tree.get_children())

        if df is None or df.empty:
            self.tree["columns"] = []
            return

        cols = list(df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col, anchor="w")
            col_width = 0
            try:
                max_data_len = df[col].astype(str).str.len().max()
                header_len = len(col)
                if pd.isna(max_data_len): max_data_len = 0
                col_width = max(int(max_data_len), header_len) * 7 + 20
            except Exception:
                col_width = len(col) * 7 + 20

            self.tree.column(col, width=min(max(col_width, 100), 350), anchor="w", stretch=True)

        for _, row in df.iterrows():
            values = [str(v) if pd.notna(v) else "" for v in row]
            self.tree.insert("", "end", values=values)

    def update_status(self, current_count):
        if self.processed_data.empty and self.original_data.empty:
            self.status_var.set(self._translations["ready"])
            return

        total_source_count = len(self.processed_data) if not self.processed_data.empty else len(self.original_data)
        status_parts = []

        status_parts.append(f"Showing {current_count} of {total_source_count}")

        if not self.processed_data.empty:
            status_parts.append("processed records")
            if self.current_filter != "All":
                status_parts.append(f"(filtered: {self.current_filter})")
        elif not self.original_data.empty:
            status_parts.append("original records")

        if self.search_var.get().strip():
            if not status_parts[-1].endswith(')'):
                status_parts.append("(searched)")
            else:
                last_part = status_parts.pop()
                last_part = last_part[:-1] + ", searched)"
                status_parts.append(last_part)

        self.status_var.set(" ".join(status_parts))

    def show_stats(self):
        if self.processed_data.empty:
            messagebox.showwarning(self._translations["stats_no_data"], self._translations["stats_no_data"])
            return

        stats = self.processed_data["Gender"].value_counts()
        total = len(self.processed_data)

        message_lines = []
        for gender in ["Male", "Female", "Unknown"]:
            count = stats.get(gender, 0)
            percentage = (count / total) * 100 if total > 0 else 0
            message_lines.append(self._translations["stats_line"].format(gender, count, percentage))

        message = self._translations["stats_message"].format("\n".join(message_lines), total)
        messagebox.showinfo(self._translations["stats_title"], message)

    def clear_data(self):
        if self.original_data.empty and self.processed_data.empty:
            return

        if messagebox.askyesno(self._translations["clear_confirm_title"], self._translations["clear_confirm_msg"]):
            self.original_data = pd.DataFrame()
            self.processed_data = pd.DataFrame()
            self.name_column = None
            self.current_filter = "All"
            self.filter_var.set("All")
            self.search_var.set("")
            self.update_display(pd.DataFrame())
            self.status_var.set(self._translations["ready"])
            self._update_widget_states()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeGenderClassifier(root)
    root.mainloop()