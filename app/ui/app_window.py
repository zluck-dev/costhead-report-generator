import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pathlib import Path
from PIL import Image, ImageDraw
import io

from app.services.excel_service import load_excel, detect_numeric_columns
from app.services.calc_service import add_sum_column, add_average_column
from app.services.export_service import export_csv, export_excel
from app.services.gin_service import list_sheets, load_sheet, auto_detect_columns, build_activity_lookup, merge_gin_with_lookup
from app.services.costhead_service import generate_costhead_report

class AppWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("GIN Mapper & CostHead Reporter")
        self.geometry("900x700")
        self.minsize(800, 600)
        self.resizable(True, True)
        
        # Set appearance
        ctk.set_appearance_mode("light")
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
        
    def _ensure_app_icon(self):
        """Generate and set a custom app icon"""
        try:
            # Create a simple icon
            img = Image.new('RGB', (64, 64), color='#1f538d')
            draw = ImageDraw.Draw(img)
            
            # Draw a simple document icon
            draw.rectangle([16, 12, 48, 52], outline='white', width=2)
            draw.rectangle([20, 16, 44, 20], fill='white')
            draw.rectangle([20, 24, 40, 28], fill='white')
            draw.rectangle([20, 32, 36, 36], fill='white')
            
            # Save to temporary file
            icon_path = "temp_icon.ico"
            img.save(icon_path, format='ICO')
            
            # Set icon
            self.iconbitmap(icon_path)
            
            # Clean up
            if os.path.exists(icon_path):
                os.remove(icon_path)
                
        except Exception as e:
            print(f"Could not set icon: {e}")
    
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
            text="GIN Mapper & CostHead Reporter", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=20)
        
        # Appearance toggle
        appearance_frame = ctk.CTkFrame(header_frame)
        appearance_frame.pack(pady=(0, 20))
        
        ctk.CTkLabel(appearance_frame, text="Appearance:").pack(side="left", padx=10)
        self.appearance_var = ctk.StringVar(value="light")
        appearance_menu = ctk.CTkOptionMenu(
            appearance_frame,
            values=["light", "dark"],
            variable=self.appearance_var,
            command=self._change_appearance,
            width=100
        )
        appearance_menu.pack(side="left", padx=10)
        
        # Create notebook for tabs
        self.notebook = ctk.CTkTabview(main_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # GIN Mapper Tab
        self.gin_tab = self.notebook.add("GIN Mapper")
        self._create_gin_mapper_ui()
        
        # CostHead Reporter Tab
        self.costhead_tab = self.notebook.add("CostHead Reporter")
        self._create_costhead_reporter_ui()
        
    def _create_gin_mapper_ui(self):
        """Create GIN Mapper UI"""
        # Activities section
        activities_frame = ctk.CTkFrame(self.gin_tab)
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
        
        # Activities sheet selection
        activities_sheet_frame = ctk.CTkFrame(activities_frame)
        activities_sheet_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(activities_sheet_frame, text="Sheet:").pack(side="left", padx=10, pady=10)
        
        self.activities_sheet_var = ctk.StringVar()
        self.activities_sheet_combo = ctk.CTkOptionMenu(
            activities_sheet_frame,
            variable=self.activities_sheet_var,
            values=["Select sheet first"],
            width=200
        )
        self.activities_sheet_combo.pack(side="left", padx=10, pady=10)
        
        ctk.CTkButton(
            activities_sheet_frame,
            text="Load Sheet",
            command=self._load_activities_sheet,
            corner_radius=8
        ).pack(side="left", padx=10, pady=10)
        
        # GIN section
        gin_frame = ctk.CTkFrame(self.gin_tab)
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
        
        # GIN sheet selection
        gin_sheet_frame = ctk.CTkFrame(gin_frame)
        gin_sheet_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(gin_sheet_frame, text="Sheet:").pack(side="left", padx=10, pady=10)
        
        self.gin_sheet_var = ctk.StringVar()
        self.gin_sheet_combo = ctk.CTkOptionMenu(
            gin_sheet_frame,
            variable=self.gin_sheet_var,
            values=["Select sheet first"],
            width=200
        )
        self.gin_sheet_combo.pack(side="left", padx=10, pady=10)
        
        ctk.CTkButton(
            gin_sheet_frame,
            text="Load Sheet",
            command=self._load_gin_sheet,
            corner_radius=8
        ).pack(side="left", padx=10, pady=10)
        
        # Export section
        export_frame = ctk.CTkFrame(self.gin_tab)
        export_frame.pack(fill="x", padx=20, pady=10)
        
        export_title = ctk.CTkLabel(
            export_frame, 
            text="Export Options", 
            font=ctk.CTkFont(size=16, weight="bold")
        )
        export_title.pack(pady=(10, 5))
        
        export_options_frame = ctk.CTkFrame(export_frame)
        export_options_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(export_options_frame, text="Export as:").pack(side="left", padx=10, pady=10)
        
        self.export_format_var = ctk.StringVar(value="Excel")
        export_format_menu = ctk.CTkOptionMenu(
            export_options_frame,
            values=["Excel", "CSV"],
            variable=self.export_format_var,
            width=100
        )
        export_format_menu.pack(side="left", padx=10, pady=10)
        
        ctk.CTkButton(
            export_options_frame,
            text="Merge & Export",
            command=self._on_merge_export,
            corner_radius=8,
            fg_color="#1f538d",
            hover_color="#14375e"
        ).pack(side="left", padx=20, pady=10)
        
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
    
    # GIN Mapper methods
    def _browse_activities(self):
        """Browse for activities file"""
        file_path = filedialog.askopenfilename(
            title="Select Activities File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.activities_file = file_path
            self.activities_file_label.configure(text=os.path.basename(file_path))
            
            # Load sheets
            try:
                sheets = list_sheets(file_path)
                self.activities_sheet_combo.configure(values=sheets)
                if sheets:
                    self.activities_sheet_var.set(sheets[0])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load sheets: {str(e)}")
    
    def _browse_gin(self):
        """Browse for GIN file"""
        file_path = filedialog.askopenfilename(
            title="Select GIN File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.gin_file = file_path
            self.gin_file_label.configure(text=os.path.basename(file_path))
            
            # Load sheets
            try:
                sheets = list_sheets(file_path)
                self.gin_sheet_combo.configure(values=sheets)
                if sheets:
                    self.gin_sheet_var.set(sheets[0])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load sheets: {str(e)}")
    
    def _load_activities_sheet(self):
        """Load activities sheet and auto-detect columns"""
        if not self.activities_file or not self.activities_sheet_var.get():
            messagebox.showwarning("Warning", "Please select activities file and sheet first")
            return
        
        try:
            self.activities_df = load_sheet(self.activities_file, self.activities_sheet_var.get())
            messagebox.showinfo("Success", f"Loaded activities sheet with {len(self.activities_df)} rows")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load activities sheet: {str(e)}")
    
    def _load_gin_sheet(self):
        """Load GIN sheet and auto-detect columns"""
        if not self.gin_file or not self.gin_sheet_var.get():
            messagebox.showwarning("Warning", "Please select GIN file and sheet first")
            return
        
        try:
            self.gin_df = load_sheet(self.gin_file, self.gin_sheet_var.get())
            messagebox.showinfo("Success", f"Loaded GIN sheet with {len(self.gin_df)} rows")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load GIN sheet: {str(e)}")
    
    def _on_merge_export(self):
        """Merge GIN with activities and export"""
        if not self.activities_df is not None or not self.gin_df is not None:
            messagebox.showwarning("Warning", "Please load both activities and GIN sheets first")
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
            
            # Export
            if self.export_format_var.get() == "Excel":
                output_path = os.path.join(output_dir, "GIN_Mapped.xlsx")
                export_excel(merged_df, output_path)
            else:
                output_path = os.path.join(output_dir, "GIN_Mapped.csv")
                export_csv(merged_df, output_path)
            
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
        self.activities_sheet_combo.configure(values=["Select sheet first"])
        self.gin_sheet_combo.configure(values=["Select sheet first"])
        self.activities_sheet_var.set("")
        self.gin_sheet_var.set("")
    
    # CostHead Reporter methods
    def _browse_gin_mapped(self):
        """Browse for GIN Mapped file"""
        file_path = filedialog.askopenfilename(
            title="Select GIN Mapped File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.gin_mapped_file = file_path
            self.gin_mapped_file_label.configure(text=os.path.basename(file_path))
    
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
        """Generate CostHead reports"""
        if not hasattr(self, 'gin_mapped_file') or not self.costhead_file:
            messagebox.showwarning("Warning", "Please select both GIN Mapped and CostHead files first")
            return
        
        try:
            # Ask for output directory
            output_dir = filedialog.askdirectory(title="Select Output Directory for CostHead Reports")
            if not output_dir:
                return
            
            # Generate reports
            result = generate_costhead_report(
                self.gin_mapped_file,
                self.costhead_file,
                output_dir,
                self.match_mode_var.get()
            )
            
            messagebox.showinfo(
                "Success", 
                f"CostHead reports generated successfully!\n\n"
                f"Output file: {result['output_file']}\n"
                f"Sheets created: {', '.join(result['sheets'])}"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate CostHead reports: {str(e)}")
    
    def _on_clear_costhead(self):
        """Clear all CostHead reporter data"""
        if hasattr(self, 'gin_mapped_file'):
            delattr(self, 'gin_mapped_file')
        self.costhead_file = ""
        
        self.gin_mapped_file_label.configure(text="No file selected")
        self.costhead_file_label.configure(text="No file selected")
