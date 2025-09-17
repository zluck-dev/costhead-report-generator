import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pathlib import Path
from PIL import Image, ImageDraw
import io
import json
import sys

from app.services.excel_service import load_excel, detect_numeric_columns
from app.services.calc_service import add_sum_column, add_average_column
from app.services.export_service import export_csv, export_excel
from app.services.gin_service import list_sheets, load_sheet, auto_detect_columns, build_activity_lookup, merge_gin_with_lookup
from app.services.costhead_service import generate_costhead_report

class AppWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("CostHead Report Generator")
        self.geometry("800x640")
        self.minsize(800, 640)
        self.resizable(True, True)

        # Set appearance (persisted)
        self._theme_store = Path.home() / ".costhead_ui_prefs.json"
        self.current_theme = self._load_saved_theme(default="dark")
        ctk.set_appearance_mode(self.current_theme)
        ctk.set_default_color_theme("blue")

        # Ensure app icon
        self._ensure_app_icon()

        # Initialize variables
        self.activities_file = ""
        self.gin_file = ""
        self.activities_df = None
        self.gin_df = None
        self.costhead_file = ""

        # Create UI
        self._create_ui()

        # Bring window to front on launch
        self.after(100, self._bring_to_front)

    def _resource_base(self) -> Path:
        try:
            return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent.parent))
        except Exception:
            return Path.cwd()

    def _asset_path(self, *parts: str) -> Path:
        return self._resource_base().joinpath(*parts)

    def _ensure_app_icon(self):
        """Do not set a custom app icon (removed)."""
        return

    def _create_ui(self):
        """Create the main UI"""
        # Main frame
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header_frame = ctk.CTkFrame(main_frame)
        header_frame.pack(fill="x", pady=(0, 20))

        title_label = ctk.CTkLabel(
            header_frame,
            text="CostHead Report Generator",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=20)

        # Top-right theme toggle icon button
        try:
            icon_path = self._asset_path("assets", "day-and-night.png")
            if icon_path.exists():
                img = ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(22, 22))
            else:
                img = None
        except Exception:
            img = None
        def _place_theme_button():
            btn = ctk.CTkButton(
                header_frame,
                text="",
                image=img,
                width=36,
                height=36,
                corner_radius=18,
                fg_color="#e0e0e0",
                hover_color="#c8c8c8",
                command=self._toggle_theme
            )
            # Place at top-right of the header card
            btn.place(relx=1.0, rely=0.0, x=-10, y=10, anchor="ne")
            self.theme_toggle_btn = btn
        _place_theme_button()
        # Apply persisted theme styles to button
        if self.current_theme == "dark":
            self.theme_toggle_btn.configure(fg_color="#3a3a3a", hover_color="#2e2e2e")

        # Create content area
        self.content = ctk.CTkFrame(main_frame)
        self.content.pack(fill="both", expand=True)

        # Build single-page UI (formerly GIN Mapper tab)
        self._create_gin_mapper_ui()

    def _create_gin_mapper_ui(self):
        """Create GIN Mapper UI"""
        # Activities section
        activities_frame = ctk.CTkFrame(self.content)
        activities_frame.pack(fill="x", padx=20, pady=10)

        activities_title = ctk.CTkLabel(
            activities_frame,
            text="Activities (Master)",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        activities_title.pack(pady=(10, 5))

        # Activities file selection
        activities_file_frame = ctk.CTkFrame(activities_frame)
        activities_file_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            activities_file_frame,
            text="Browse Activities File",
            command=self._browse_activities,
            corner_radius=8
        ).pack(side="left", padx=10, pady=10)

        self.activities_file_label = ctk.CTkLabel(
            activities_file_frame,
            text="No file selected",
            text_color="gray"
        )
        self.activities_file_label.pack(side="left", padx=10, pady=10)

                # Activities sheet selection (hidden)
        # activities_sheet_frame = ctk.CTkFrame(activities_frame)
        # activities_sheet_frame.pack(fill="x", padx=10, pady=5)
        # ctk.CTkLabel(activities_sheet_frame, text="Sheet:").pack(side="left", padx=10, pady=10)
        # self.activities_sheet_var = ctk.StringVar()
        # self.activities_sheet_combo = ctk.CTkOptionMenu(
        #     activities_sheet_frame,
        #     variable=self.activities_sheet_var,
        #     values=["Select sheet first"],
        #     width=200
        # )
        # self.activities_sheet_combo.pack(side="left", padx=10, pady=10)
        # ctk.CTkButton(
        #     activities_sheet_frame,
        #     text="Load Sheet",
        #     command=self._load_activities_sheet,
        #     corner_radius=8
        # ).pack(side="left", padx=10, pady=10)

        # GIN section
        gin_frame = ctk.CTkFrame(self.content)
        gin_frame.pack(fill="x", padx=20, pady=10)

        gin_title = ctk.CTkLabel(
            gin_frame,
            text="GIN (Good Issue Note)",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        gin_title.pack(pady=(10, 5))

        # GIN file selection
        gin_file_frame = ctk.CTkFrame(gin_frame)
        gin_file_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            gin_file_frame,
            text="Browse GIN File",
            command=self._browse_gin,
            corner_radius=8
        ).pack(side="left", padx=10, pady=10)

        self.gin_file_label = ctk.CTkLabel(
            gin_file_frame,
            text="No file selected",
            text_color="gray"
        )
        self.gin_file_label.pack(side="left", padx=10, pady=10)

                # GIN sheet selection (hidden)
        # gin_sheet_frame = ctk.CTkFrame(gin_frame)
        # gin_sheet_frame.pack(fill="x", padx=10, pady=5)
        # ctk.CTkLabel(gin_sheet_frame, text="Sheet:").pack(side="left", padx=10, pady=10)
        # self.gin_sheet_var = ctk.StringVar()
        # self.gin_sheet_combo = ctk.CTkOptionMenu(
        #     gin_sheet_frame,
        #     variable=self.gin_sheet_var,
        #     values=["Select sheet first"],
        #     width=200
        # )
        # self.gin_sheet_combo.pack(side="left", padx=10, pady=10)
        # ctk.CTkButton(
        #     gin_sheet_frame,
        #     text="Load Sheet",
        #     command=self._load_gin_sheet,
        #     corner_radius=8
        # ).pack(side="left", padx=10, pady=10)

        # CostHead Mapping section
        costhead_frame = ctk.CTkFrame(self.content)
        costhead_frame.pack(fill="x", padx=20, pady=10)

        costhead_title = ctk.CTkLabel(
            costhead_frame,
            text="CostHead Mapping File",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        costhead_title.pack(pady=(10, 5))

        costhead_file_frame = ctk.CTkFrame(costhead_frame)
        costhead_file_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            costhead_file_frame,
            text="Browse CostHead File",
            command=self._browse_costhead,
            corner_radius=8
        ).pack(side="left", padx=10, pady=10)

        self.costhead_file_label = ctk.CTkLabel(
            costhead_file_frame,
            text="No file selected",
            text_color="gray"
        )
        self.costhead_file_label.pack(side="left", padx=10, pady=10)

        # Export section
        export_frame = ctk.CTkFrame(self.content)
        export_frame.pack(fill="x", padx=20, pady=10)

        export_options_frame = ctk.CTkFrame(export_frame)
        export_options_frame.pack(fill="x", padx=10, pady=5)

        # Center the Merge & Export button in the row
        export_options_frame.grid_columnconfigure(0, weight=1)
        center_container = ctk.CTkFrame(export_options_frame)
        center_container.grid(row=0, column=0, sticky="nsew")
        center_container.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            center_container,
            text="Generate CostHead Report",
            command=self._on_generate_all,
            corner_radius=8,
            fg_color="#1f538d",
            hover_color="#14375e"
        ).grid(row=0, column=0, padx=20, pady=10, sticky="n")

        # Clear button
        ctk.CTkButton(
            export_frame,
            text="Clear All",
            command=self._on_clear,
            corner_radius=8,
            fg_color="#dc3545",
            hover_color="#c82333"
        ).pack(pady=10)

    def _create_costhead_reporter_ui(self):
        """Create CostHead Reporter UI"""
        # GIN Mapped file section
        gin_mapped_frame = ctk.CTkFrame(self.costhead_tab)
        gin_mapped_frame.pack(fill="x", padx=20, pady=10)

        gin_mapped_title = ctk.CTkLabel(
            gin_mapped_frame,
            text="GIN Mapped File (from GIN Mapper)",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        gin_mapped_title.pack(pady=(10, 5))

        # GIN Mapped file selection
        gin_mapped_file_frame = ctk.CTkFrame(gin_mapped_frame)
        gin_mapped_file_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            gin_mapped_file_frame,
            text="Browse GIN Mapped File",
            command=self._browse_gin_mapped,
            corner_radius=8
        ).pack(side="left", padx=10, pady=10)

        self.gin_mapped_file_label = ctk.CTkLabel(
            gin_mapped_file_frame,
            text="No file selected",
            text_color="gray"
        )
        self.gin_mapped_file_label.pack(side="left", padx=10, pady=10)

        # CostHead file section
        costhead_file_frame = ctk.CTkFrame(self.costhead_tab)
        costhead_file_frame.pack(fill="x", padx=20, pady=10)

        costhead_title = ctk.CTkLabel(
            costhead_file_frame,
            text="CostHead Mapping File",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        costhead_title.pack(pady=(10, 5))

        # CostHead file selection
        costhead_file_selection_frame = ctk.CTkFrame(costhead_file_frame)
        costhead_file_selection_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            costhead_file_selection_frame,
            text="Browse CostHead File",
            command=self._browse_costhead,
            corner_radius=8
        ).pack(side="left", padx=10, pady=10)

        self.costhead_file_label = ctk.CTkLabel(
            costhead_file_selection_frame,
            text="No file selected",
            text_color="gray"
        )
        self.costhead_file_label.pack(side="left", padx=10, pady=10)

        # Matching options
        matching_frame = ctk.CTkFrame(self.costhead_tab)
        matching_frame.pack(fill="x", padx=20, pady=10)

        matching_title = ctk.CTkLabel(
            matching_frame,
            text="Matching Options",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        matching_title.pack(pady=(10, 5))

        matching_options_frame = ctk.CTkFrame(matching_frame)
        matching_options_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(matching_options_frame, text="Match Mode:").pack(side="left", padx=10, pady=10)

        self.match_mode_var = ctk.StringVar(value="contains")
        match_mode_menu = ctk.CTkOptionMenu(
            matching_options_frame,
            values=["contains", "exact"],
            variable=self.match_mode_var,
            width=100
        )
        match_mode_menu.pack(side="left", padx=10, pady=10)

        # Generate reports section
        generate_frame = ctk.CTkFrame(self.costhead_tab)
        generate_frame.pack(fill="x", padx=20, pady=10)

        generate_title = ctk.CTkLabel(
            generate_frame,
            text="Generate Reports",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        generate_title.pack(pady=(10, 5))

        ctk.CTkButton(
            generate_frame,
            text="Generate CostHead Reports",
            command=self._on_generate_costhead_reports,
            corner_radius=8,
            fg_color="#28a745",
            hover_color="#218838"
        ).pack(pady=10)

        # Clear button
        ctk.CTkButton(
            generate_frame,
            text="Clear All",
            command=self._on_clear_costhead,
            corner_radius=8,
            fg_color="#dc3545",
            hover_color="#c82333"
        ).pack(pady=10)

    def _change_appearance(self, mode):
        """Change appearance mode"""
        ctk.set_appearance_mode(mode)

    def _toggle_theme(self):
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        ctk.set_appearance_mode(self.current_theme)
        self._save_theme() # Save theme after toggling
        # Subtle button background tweak
        if hasattr(self, 'theme_toggle_btn'):
            if self.current_theme == "dark":
                self.theme_toggle_btn.configure(fg_color="#3a3a3a", hover_color="#2e2e2e")
            else:
                self.theme_toggle_btn.configure(fg_color="#e0e0e0", hover_color="#c8c8c8")

    # GIN Mapper methods
    def _browse_activities(self):
        """Browse and auto-load first sheet from activities file"""
        file_path = filedialog.askopenfilename(
            title="Select Activities File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.activities_file = file_path
            self.activities_file_label.configure(text=os.path.basename(file_path))
            try:
                sheets = list_sheets(file_path)
                first_sheet = sheets[0] if sheets else None
                if not first_sheet:
                    raise ValueError("No sheets found in Activities file")
                self.activities_df = load_sheet(self.activities_file, first_sheet)
                messagebox.showinfo("Success", f"Loaded activities sheet '{first_sheet}' with {len(self.activities_df)} rows")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load activities sheet: {str(e)}")

    def _browse_gin(self):
        """Browse and auto-load first sheet from GIN file"""
        file_path = filedialog.askopenfilename(
            title="Select GIN File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.gin_file = file_path
            self.gin_file_label.configure(text=os.path.basename(file_path))
            try:
                sheets = list_sheets(file_path)
                first_sheet = sheets[0] if sheets else None
                if not first_sheet:
                    raise ValueError("No sheets found in GIN file")
                self.gin_df = load_sheet(self.gin_file, first_sheet)
                messagebox.showinfo("Success", f"Loaded GIN sheet '{first_sheet}' with {len(self.gin_df)} rows")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load GIN sheet: {str(e)}")

    def _load_activities_sheet(self):
        """Deprecated (auto-loads on file select)"""
        messagebox.showinfo("Info", "Activities sheet auto-loads on file selection.")

    def _load_gin_sheet(self):
        """Deprecated (auto-loads on file select)"""
        messagebox.showinfo("Info", "GIN sheet auto-loads on file selection.")

    def _on_merge_export(self):
        """Merge GIN with activities and export"""
        if self.activities_df is None or self.gin_df is None:
            messagebox.showwarning("Warning", "Please select both Activities and GIN files first")
            return

        try:
            # Auto-detect columns
            activities_cols = auto_detect_columns(self.activities_df, "master")
            gin_cols = auto_detect_columns(self.gin_df, "gin")

            # Check if all required columns are detected
            if not all([activities_cols["code"], activities_cols["name"], activities_cols["wbs"]]):
                messagebox.showerror("Error", "Could not detect required columns in activities sheet")
                return

            if not gin_cols["code"]:
                messagebox.showerror("Error", "Could not detect activity code column in GIN sheet")
                return

            # Build lookup and merge
            activities_lookup = build_activity_lookup(
                self.activities_df,
                activities_cols["code"],
                activities_cols["wbs"],
                activities_cols["name"]
            )

            merged_df = merge_gin_with_lookup(
                self.gin_df,
                gin_cols["code"],
                activities_lookup
            )

            # Ask for output directory
            output_dir = filedialog.askdirectory(title="Select Output Directory")
            if not output_dir:
                return

            # Always export as Excel
            output_base = os.path.join(output_dir, "GIN_Mapped")
            output_path = export_excel(merged_df, output_base)

            messagebox.showinfo("Success", f"Exported merged data to:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge and export: {str(e)}")

    def _on_clear(self):
        """Clear all GIN mapper data"""
        self.activities_file = ""
        self.gin_file = ""
        self.activities_df = None
        self.gin_df = None

        self.activities_file_label.configure(text="No file selected")
        self.gin_file_label.configure(text="No file selected")
        # Hidden widgets: keep state consistent if they exist
        if hasattr(self, 'activities_sheet_combo'):
            self.activities_sheet_combo.configure(values=["Select sheet first"])
        if hasattr(self, 'gin_sheet_combo'):
            self.gin_sheet_combo.configure(values=["Select sheet first"])
        if hasattr(self, 'activities_sheet_var'):
            self.activities_sheet_var.set("")
        if hasattr(self, 'gin_sheet_var'):
            self.gin_sheet_var.set("")

    # CostHead Reporter methods
    def _browse_gin_mapped(self):
        """Deprecated: CostHead reporter tab removed"""
        messagebox.showinfo("Info", "Use the GIN Mapper tab to select files and generate the report.")

    def _browse_costhead(self):
        """Browse for CostHead file"""
        file_path = filedialog.askopenfilename(
            title="Select CostHead File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.costhead_file = file_path
            self.costhead_file_label.configure(text=os.path.basename(file_path))

    def _on_generate_costhead_reports(self):
        """Deprecated: Use Generate CostHead Report in GIN Mapper tab"""
        messagebox.showinfo("Info", "Use the GIN Mapper tab to generate the CostHead report.")

    def _on_clear_costhead(self):
        """Deprecated: CostHead reporter tab removed"""
        self.costhead_file = ""
        if hasattr(self, 'costhead_file_label'):
            self.costhead_file_label.configure(text="No file selected")

    def _on_generate_all(self):
        """Export GIN_Mapped then generate CostHead report in a new folder."""
        # Preconditions
        if self.activities_df is None or self.gin_df is None:
            messagebox.showwarning("Warning", "Please select both Activities and GIN files first")
            return
        if not self.costhead_file:
            messagebox.showwarning("Warning", "Please select CostHead mapping file")
            return
        try:
            # Auto-detect columns
            activities_cols = auto_detect_columns(self.activities_df, "master")
            gin_cols = auto_detect_columns(self.gin_df, "gin")
            if not all([activities_cols["code"], activities_cols["name"], activities_cols["wbs"]]):
                messagebox.showerror("Error", "Could not detect required columns in activities sheet")
                return
            if not gin_cols["code"]:
                messagebox.showerror("Error", "Could not detect activity code column in GIN sheet")
                return
            # Build lookup and merge
            activities_lookup = build_activity_lookup(
                self.activities_df,
                activities_cols["code"],
                activities_cols["wbs"],
                activities_cols["name"]
            )
            merged_df = merge_gin_with_lookup(self.gin_df, gin_cols["code"], activities_lookup)

            # Choose base output directory
            base_dir = filedialog.askdirectory(title="Select Output Directory for CostHead Report")
            if not base_dir:
                return

            # Find next CostHead Report folder name
            def next_folder(base: str) -> str:
                root = Path(base)
                name = "CostHead Report"
                candidate = root / name
                idx = 1
                while candidate.exists():
                    candidate = root / f"{name} {idx}"
                    idx += 1
                candidate.mkdir(parents=True, exist_ok=True)
                return str(candidate)

            out_dir = next_folder(base_dir)

            # 1) Write GIN_Mapped.xlsx
            gin_mapped_base = os.path.join(out_dir, "GIN_Mapped")
            gin_mapped_path = export_excel(merged_df, gin_mapped_base)

            # 2) Generate CostHead report using that file
            result = generate_costhead_report(
                gin_mapped_path,
                self.costhead_file,
                out_dir,
                self.match_mode_var.get() if hasattr(self, 'match_mode_var') else "contains"
            )

            # 3) Append the GIN_Mapped sheet into the CostHead_Reports.xlsx as an extra sheet
            try:
                report_path = result["output_file"]
                # Open and append
                import pandas as pd
                with pd.ExcelWriter(report_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
                    merged_df.to_excel(xw, sheet_name="GIN_Mapped", index=False)
            except Exception:
                pass

            messagebox.showinfo(
                "Success",
                f"Generated CostHead report.\nFolder: {out_dir}\nReport: {result['output_file']}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate CostHead report: {str(e)}")

    def _load_saved_theme(self, default="dark"):
        """Load saved theme from a JSON file."""
        if self._theme_store.exists():
            try:
                with open(self._theme_store, "r") as f:
                    data = json.load(f)
                    return data.get("theme", default)
            except json.JSONDecodeError:
                pass # Ignore if file is corrupted or empty
        return default

    def _save_theme(self):
        """Save current theme to a JSON file."""
        try:
            with open(self._theme_store, "w") as f:
                json.dump({"theme": self.current_theme}, f)
        except Exception as e:
            print(f"Error saving theme: {e}")

    def _bring_to_front(self):
        try:
            # Make sure the window is visible and raised
            self.deiconify()
            self.lift()
            self.update_idletasks()
            # Hold topmost briefly to ensure we come to foreground
            self.attributes('-topmost', True)
            self.focus_force()
            # After a short delay, drop topmost and refocus again
            def _drop_topmost():
                try:
                    self.attributes('-topmost', False)
                    self.lift()
                    self.focus_force()
                except Exception:
                    pass
            self.after(1200, _drop_topmost)
        except Exception:
            pass
