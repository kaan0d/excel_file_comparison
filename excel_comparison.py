import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import json
import os
import sys

class SettingsManager:
    """Manage application settings with JSON persistence."""
    
    def __init__(self, settings_file="excel_compare_settings.json"):
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.settings_file = os.path.join(app_dir, settings_file)
        
        self.default_settings = {
            'code_column_index': 1,
            'name_column_index': 5,
            'incoming_column_index': 6,
            'outgoing_column_index': 7,
            'remaining_column_index': 8,
            'custom_comparisons': []
        }
        self.settings = self.load_settings()
    
    def load_settings(self):
        """Load settings from file or create & save defaults."""
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass

        settings = self.default_settings.copy()
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2)
        except:
            pass
        return settings
    
    def save_settings(self):
        """Save current settings to file."""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2)
            return True
        except:
            return False
    
    def get(self, key):
        """Get a setting value."""
        return self.settings.get(key, self.default_settings.get(key))
    
    def set(self, key, value):
        """Set a setting value."""
        self.settings[key] = value
    
    def reset_to_defaults(self):
        """Reset all settings to default values."""
        self.settings = self.default_settings.copy()


class SettingsWindow:
    """Settings window for column configuration."""
    
    def __init__(self, parent, settings_manager):
        self.parent = parent
        self.settings_manager = settings_manager
        
        self.window = tk.Toplevel(parent)
        self.window.title("Column Settings")
        self.window.geometry("550x520")
        self.window.configure(bg="#1e1e1e")
        self.window.resizable(False, False)
        
        # Make modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Store mousewheel binding ID
        self.mousewheel_bind_id = None
        
        self.setup_styles()
        self.build_ui()
        
        # Cleanup on close
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Center window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def on_closing(self):
        """Clean up before closing."""
        if self.mousewheel_bind_id:
            self.window.unbind_all("<MouseWheel>")
        self.window.destroy()
    
    def setup_styles(self):
        """Configure styles for settings window."""
        style = ttk.Style()
        
        bg = "#1e1e1e"
        card_bg = "#2d2d30"
        text_color = "#F3F4F6"
        muted_text = "#9CA3AF"
        primary = "#3B82F6"
        primary_dark = "#2563EB"
        success = "#10B981"
        success_dark = "#059669"
        
        style.configure("Settings.TFrame", background=bg)
        style.configure("SettingsCard.TFrame", background=card_bg)
        
        style.configure(
            "SettingsTitle.TLabel",
            background=bg,
            foreground=text_color,
            font=("Segoe UI", 16, "bold")
        )
        
        style.configure(
            "SettingsLabel.TLabel",
            background=card_bg,
            foreground=text_color,
            font=("Segoe UI", 10)
        )
        
        style.configure(
            "SettingsInfo.TLabel",
            background=card_bg,
            foreground=muted_text,
            font=("Segoe UI", 9, "italic")
        )
        
        style.configure(
            "Success.TButton",
            font=("Segoe UI", 9, "bold"),
            padding=(10, 6),
            relief="flat",
            borderwidth=0,
            background=success,
            foreground="white"
        )
        style.map(
            "Success.TButton",
            background=[("active", success_dark), ("pressed", "#047857")]
        )
        
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            padding=(10, 6),
            relief="flat",
            borderwidth=0,
            background="#374151",
            foreground=text_color
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#4B5563"), ("pressed", "#1F2937")]
        )
    
    def build_ui(self):
        """Build the settings UI."""
        main_frame = ttk.Frame(self.window, style="Settings.TFrame", padding=20)
        main_frame.pack(fill="both", expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="Column Settings",
            style="SettingsTitle.TLabel"
        )
        title_label.pack(anchor="w", pady=(0, 5))
        
        subtitle = ttk.Label(
            main_frame,
            text="Configure column indices in Excel files (starts from 0)",
            style="SettingsInfo.TLabel"
        )
        subtitle.pack(anchor="w", pady=(0, 15))
        
        # Settings card
        card = ttk.Frame(main_frame, style="SettingsCard.TFrame", padding=20)
        card.pack(fill="both", expand=True, pady=(0, 15))
        
        # Create scrollable frame for settings
        canvas = tk.Canvas(card, bg="#2d2d30", highlightthickness=0)
        scrollable_frame = ttk.Frame(canvas, style="SettingsCard.TFrame")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Make canvas expand to fill width
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # Mouse wheel scrolling with error handling
        def on_mousewheel(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        self.mousewheel_bind_id = canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Create input fields
        self.entries = {}
        
        settings_config = [
            ('code_column_index', 'Code Column Index:', 'Column containing unique code'),
            ('name_column_index', 'Product Name Column Index:', 'Column containing product name'),
            ('incoming_column_index', 'Incoming Column Index:', 'Column containing "Incoming" value'),
            ('outgoing_column_index', 'Outgoing Column Index:', 'Column containing "Outgoing" value'),
            ('remaining_column_index', 'Remaining Column Index:', 'Column containing "Remaining" value')
        ]
        
        for i, (key, label_text, help_text) in enumerate(settings_config):
            field_frame = ttk.Frame(scrollable_frame, style="SettingsCard.TFrame")
            field_frame.pack(fill="x", pady=(0, 15))
            
            label = ttk.Label(
                field_frame,
                text=label_text,
                style="SettingsLabel.TLabel",
                width=25,
                anchor="w"
            )
            label.pack(side="left")
            
            entry = ttk.Entry(field_frame, width=10, font=("Segoe UI", 10))
            entry.pack(side="left", padx=(10, 0))
            entry.insert(0, str(self.settings_manager.get(key)))
            self.entries[key] = entry
            
            help_label = ttk.Label(
                field_frame,
                text=f"  ({help_text})",
                style="SettingsInfo.TLabel"
            )
            help_label.pack(side="left", padx=(10, 0))
        
        # Separator
        separator = ttk.Separator(scrollable_frame, orient="horizontal")
        separator.pack(fill="x", pady=15)
        
        # Custom comparisons section
        custom_header = ttk.Frame(scrollable_frame, style="SettingsCard.TFrame")
        custom_header.pack(fill="x", pady=(0, 10))
        
        ttk.Label(
            custom_header,
            text="Custom Comparisons",
            style="SettingsLabel.TLabel",
            font=("Segoe UI", 11, "bold")
        ).pack(side="left")
        
        ttk.Button(
            custom_header,
            text="+ Add",
            style="Secondary.TButton",
            command=self.add_custom_comparison
        ).pack(side="right")
        
        # Custom comparisons list
        self.custom_frame = ttk.Frame(scrollable_frame, style="SettingsCard.TFrame")
        self.custom_frame.pack(fill="x")
        
        # Load existing custom comparisons
        self.custom_entries = []
        custom_comparisons = self.settings_manager.get('custom_comparisons') or []
        for comp in custom_comparisons:
            self.add_custom_comparison_row(comp['name'], comp['index'])
        
        canvas.pack(side="left", fill="both", expand=True)
        
        # Buttons
        button_frame = ttk.Frame(main_frame, style="Settings.TFrame")
        button_frame.pack(fill="x")
        
        ttk.Button(
            button_frame,
            text="Reset to Defaults",
            style="Secondary.TButton",
            command=self.reset_to_defaults
        ).pack(side="left")
        
        ttk.Button(
            button_frame,
            text="Cancel",
            style="Secondary.TButton",
            command=self.window.destroy
        ).pack(side="right", padx=(5, 0))
        
        ttk.Button(
            button_frame,
            text="Save",
            style="Success.TButton",
            command=self.save_settings
        ).pack(side="right")
    
    def add_custom_comparison(self):
        """Add a new custom comparison row."""
        self.add_custom_comparison_row("", "")
    
    def add_custom_comparison_row(self, name="", index=""):
        """Add a custom comparison input row."""
        row_frame = ttk.Frame(self.custom_frame, style="SettingsCard.TFrame")
        row_frame.pack(fill="x", pady=(0, 8))
        
        ttk.Label(
            row_frame,
            text="Name:",
            style="SettingsLabel.TLabel",
            width=8,
            anchor="w"
        ).pack(side="left", padx=(0, 5))
        
        name_entry = ttk.Entry(row_frame, width=20, font=("Segoe UI", 9))
        name_entry.pack(side="left", padx=(0, 10))
        name_entry.insert(0, name)
        
        ttk.Label(
            row_frame,
            text="Index:",
            style="SettingsLabel.TLabel"
        ).pack(side="left", padx=(5, 5))
        
        index_entry = ttk.Entry(row_frame, width=8, font=("Segoe UI", 9))
        index_entry.pack(side="left", padx=(0, 10))
        index_entry.insert(0, str(index))
        
        def remove_row():
            row_frame.destroy()
            self.custom_entries.remove((name_entry, index_entry, row_frame))
        
        remove_btn = ttk.Button(
            row_frame,
            text="Remove",
            style="Secondary.TButton",
            command=remove_row
        )
        remove_btn.pack(side="left")
        
        self.custom_entries.append((name_entry, index_entry, row_frame))
    
    def reset_to_defaults(self):
        """Reset all settings to default values."""
        if messagebox.askyesno(
            "Reset to Defaults",
            "Are you sure you want to reset all settings to default values?",
            parent=self.window
        ):
            self.settings_manager.reset_to_defaults()
            for key, entry in self.entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, str(self.settings_manager.get(key)))
            
            # Clear custom comparisons
            for name_entry, index_entry, row_frame in self.custom_entries:
                row_frame.destroy()
            self.custom_entries.clear()
    
    def save_settings(self):
        """Validate and save settings."""
        try:
            # Validate all entries are valid integers
            new_settings = {}
            for key, entry in self.entries.items():
                value = entry.get().strip()
                if not value.isdigit():
                    messagebox.showerror(
                        "Invalid Value",
                        f"Please enter a valid number (0 or greater) for '{key}'.",
                        parent=self.window
                    )
                    return
                new_settings[key] = int(value)
            
            # Validate custom comparisons
            custom_comparisons = []
            for name_entry, index_entry, _ in self.custom_entries:
                name = name_entry.get().strip()
                index = index_entry.get().strip()
                
                if name and index:
                    if not index.isdigit():
                        messagebox.showerror(
                            "Invalid Value",
                            f"Please enter a valid index for custom comparison '{name}'.",
                            parent=self.window
                        )
                        return
                    custom_comparisons.append({
                        'name': name,
                        'index': int(index)
                    })
                elif name or index:
                    messagebox.showerror(
                        "Missing Information",
                        "Both name and index must be provided for custom comparisons.",
                        parent=self.window
                    )
                    return
            
            # Check for negative values
            all_indices = list(new_settings.values()) + [c['index'] for c in custom_comparisons]
            if any(v < 0 for v in all_indices):
                messagebox.showerror(
                    "Invalid Value",
                    "Column indices must be 0 or greater.",
                    parent=self.window
                )
                return
            
            # Save settings
            for key, value in new_settings.items():
                self.settings_manager.set(key, value)
            
            self.settings_manager.set('custom_comparisons', custom_comparisons)
            
            if self.settings_manager.save_settings():
                messagebox.showinfo(
                    "Success",
                    "Settings saved successfully!",
                    parent=self.window
                )
                self.on_closing()
            else:
                messagebox.showerror(
                    "Error",
                    "An error occurred while saving settings.",
                    parent=self.window
                )
        
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Error occurred while saving settings:\n{str(e)}",
                parent=self.window
            )


class ExcelComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Comparison")
        self.root.geometry("640x420")
        self.root.resizable(False, False)
        self.root.configure(bg="#1e1e1e")

        self.file1_path = None
        self.file2_path = None
        self.settings_manager = SettingsManager()

        self.setup_styles()
        self.build_ui()
    
    def setup_styles(self):
        """Configure ttk styles for a modern dark UI."""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass

        primary = "#3B82F6"
        primary_dark = "#2563EB"
        danger = "#EF4444"
        danger_dark = "#DC2626"
        bg = "#1e1e1e"
        card_bg = "#2d2d30"
        text_color = "#F3F4F6"
        muted_text = "#9CA3AF"

        self.root.configure(bg=bg)

        style.configure("Root.TFrame", background=bg)
        style.configure("Card.TFrame", background=card_bg)

        style.configure(
            "Title.TLabel",
            background=bg,
            foreground=text_color,
            font=("Segoe UI", 18, "bold")
        )

        style.configure(
            "Subtitle.TLabel",
            background=bg,
            foreground=muted_text,
            font=("Segoe UI", 10)
        )

        style.configure(
            "TLabel",
            background=card_bg,
            foreground=text_color,
            font=("Segoe UI", 10)
        )

        style.configure(
            "Info.TLabel",
            background=card_bg,
            foreground=muted_text,
            font=("Segoe UI", 9)
        )

        style.configure(
            "File.TButton",
            font=("Segoe UI", 9),
            padding=(10, 6),
            relief="flat",
            borderwidth=0,
            background=card_bg,
            foreground=text_color
        )
        style.map(
            "File.TButton",
            background=[("active", "#374151"), ("pressed", "#111827")],
            foreground=[("disabled", "#6B7280")]
        )

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(12, 8),
            relief="flat",
            borderwidth=0,
            background=primary,
            foreground="white"
        )
        style.map(
            "Primary.TButton",
            background=[("active", primary_dark), ("pressed", "#1D4ED8")],
            foreground=[("disabled", "#9CA3AF")]
        )

        style.configure(
            "Danger.TButton",
            font=("Segoe UI", 9, "bold"),
            padding=(12, 6),
            relief="flat",
            borderwidth=0,
            background=danger,
            foreground="white"
        )
        style.map(
            "Danger.TButton",
            background=[("active", danger_dark), ("pressed", "#B91C1C")],
            foreground=[("disabled", "#9CA3AF")]
        )

        style.configure(
            "Settings.TButton",
            font=("Segoe UI", 9),
            padding=(8, 6),
            relief="flat",
            borderwidth=0,
            background="#6B7280",
            foreground="white"
        )
        style.map(
            "Settings.TButton",
            background=[("active", "#4B5563"), ("pressed", "#374151")]
        )

        style.configure(
            "TCheckbutton",
            background=bg,
            foreground=text_color,
            font=("Segoe UI", 10)
        )
        style.map(
            "TCheckbutton",
            background=[("active", bg), ("!disabled", bg)],
            foreground=[("active", text_color), ("!disabled", text_color)]
        )

        style.configure(
            "FileCard.TLabelframe",
            background=card_bg,
            foreground=text_color,
            padding=10,
            borderwidth=0
        )
        style.configure(
            "FileCard.TLabelframe.Label",
            background=card_bg,
            foreground=text_color,
            font=("Segoe UI", 10, "bold")
        )

    def build_ui(self):
        """Build the main application UI."""
        main_frame = ttk.Frame(self.root, style="Root.TFrame", padding=20)
        main_frame.pack(fill="both", expand=True)

        # Header with title and settings button
        header_frame = ttk.Frame(main_frame, style="Root.TFrame")
        header_frame.pack(fill="x", pady=(0, 15))

        title_frame = ttk.Frame(header_frame, style="Root.TFrame")
        title_frame.pack(side="left", fill="x", expand=True)

        title_label = ttk.Label(
            title_frame,
            text="Excel File Comparison",
            style="Title.TLabel"
        )
        title_label.pack(anchor="w")

        subtitle_label = ttk.Label(
            title_frame,
            text="Quickly compare two Excel files and view the differences.",
            style="Subtitle.TLabel"
        )
        subtitle_label.pack(anchor="w", pady=(2, 0))

        # Settings button
        ttk.Button(
            header_frame,
            text="Settings",
            style="Settings.TButton",
            command=self.open_settings
        ).pack(side="right")

        file_card = ttk.Frame(main_frame, style="Card.TFrame", padding=15)
        file_card.pack(fill="x", pady=(0, 12))

        # File 1
        file1_frame = ttk.LabelFrame(file_card, text="File 1", style="FileCard.TLabelframe")
        file1_frame.pack(fill="x", pady=(0, 10))

        top1 = ttk.Frame(file1_frame, style="Card.TFrame")
        top1.pack(fill="x")

        ttk.Label(top1, text="Source file:", width=15, anchor="w").pack(side="left")

        self.file1_label = ttk.Label(
            top1,
            text="No file selected",
            style="Info.TLabel"
        )
        self.file1_label.pack(side="left", padx=(5, 5), expand=True, fill="x")

        ttk.Button(
            top1,
            text="Browse",
            style="File.TButton",
            command=self.select_file1
        ).pack(side="right")

        # File 2
        file2_frame = ttk.LabelFrame(file_card, text="File 2", style="FileCard.TLabelframe")
        file2_frame.pack(fill="x")

        top2 = ttk.Frame(file2_frame, style="Card.TFrame")
        top2.pack(fill="x")

        ttk.Label(top2, text="Target file:", width=15, anchor="w").pack(side="left")

        self.file2_label = ttk.Label(
            top2,
            text="No file selected",
            style="Info.TLabel"
        )
        self.file2_label.pack(side="left", padx=(5, 5), expand=True, fill="x")

        ttk.Button(
            top2,
            text="Browse",
            style="File.TButton",
            command=self.select_file2
        ).pack(side="right")

        # Options
        checkbox_frame = ttk.Frame(main_frame, style="Root.TFrame")
        checkbox_frame.pack(fill="x", pady=(5, 10))

        self.gck_check = tk.BooleanVar()
        ttk.Checkbutton(
            checkbox_frame,
            text="Detailed comparison",
            variable=self.gck_check,
            style="TCheckbutton"
        ).pack(anchor="w")

        info_label = ttk.Label(
            main_frame,
            text="Note: The last row is automatically excluded.",
            style="Info.TLabel"
        )
        info_label.pack(anchor="w", pady=(0, 10))

        # Compare button
        btn_frame = ttk.Frame(main_frame, style="Root.TFrame")
        btn_frame.pack(fill="x", pady=(0, 5))

        ttk.Button(
            btn_frame,
            text="Start Comparison",
            style="Primary.TButton",
            command=self.compare_files
        ).pack(fill="x")

    def open_settings(self):
        """Open settings window."""
        SettingsWindow(self.root, self.settings_manager)

    def select_file1(self):
        """Handle File 1 selection."""
        file_path = filedialog.askopenfilename(
            title="Select File 1",
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.file1_path = file_path
            self.file1_label.config(text=file_path.split("/")[-1])
    
    def select_file2(self):
        """Handle File 2 selection."""
        file_path = filedialog.askopenfilename(
            title="Select File 2",
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.file2_path = file_path
            self.file2_label.config(text=file_path.split("/")[-1])
    
    def compare_files(self):
        """Read files, preprocess data and open result window."""
        if not self.file1_path or not self.file2_path:
            messagebox.showerror("Error", "Please select both files!")
            return
        
        try:
            df1 = pd.read_excel(self.file1_path, engine='xlrd', header=None)
            df2 = pd.read_excel(self.file2_path, engine='xlrd', header=None)
            
            # Remove last row
            df1 = df1.iloc[:-1]
            df2 = df2.iloc[:-1]
            
            # Set header row
            df1.columns = df1.iloc[0]
            df2.columns = df2.iloc[0]
            
            # Remove header row from data
            df1 = df1.iloc[1:]
            df2 = df2.iloc[1:]
            
            # Reset index
            df1 = df1.reset_index(drop=True)
            df2 = df2.reset_index(drop=True)
            
            result = self.calculate_result(df1, df2)
            self.open_result_window(result)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing files:\n{str(e)}")
    
    def calculate_result(self, df1, df2):
        """Calculate differences between two dataframes."""
        result = {
            'row_count_1': len(df1),
            'row_count_2': len(df2),
            'missing_codes': [],
            'extra_codes': [],
            'differences': []
        }
        
        # Get column indices from settings
        code_idx = self.settings_manager.get('code_column_index')
        desc_idx = self.settings_manager.get('name_column_index')
        
        code_column = df1.columns[code_idx]
        description_column = df1.columns[desc_idx]
        
        file1_codes = set(df1[code_column].values)
        file2_codes = set(df2[code_column].values)
        
        # Codes in file1 but not in file2
        missing_codes = file1_codes - file2_codes
        for code in missing_codes:
            description = df1[df1[code_column] == code][description_column].iloc[0]
            result['missing_codes'].append((code, description))
        
        # Codes in file2 but not in file1
        extra_codes = file2_codes - file1_codes
        for code in extra_codes:
            description = df2[df2[code_column] == code][description_column].iloc[0]
            result['extra_codes'].append((code, description))
        
        # Check if we need to compare any fields
        should_compare_gck = self.gck_check.get()
        custom_comparisons = self.settings_manager.get('custom_comparisons') or []
        has_custom = len(custom_comparisons) > 0
        
        # Comparison for common codes
        if should_compare_gck or has_custom:
            common_codes = file1_codes & file2_codes
            
            # Prepare columns to compare
            columns_to_compare = []
            
            if should_compare_gck:
                # Get GCK column indices from settings
                incoming_idx = self.settings_manager.get('incoming_column_index')
                outgoing_idx = self.settings_manager.get('outgoing_column_index')
                remaining_idx = self.settings_manager.get('remaining_column_index')
                
                columns_to_compare.extend([
                    ('Incoming', df1.columns[incoming_idx]),
                    ('Outgoing', df1.columns[outgoing_idx]),
                    ('Remaining', df1.columns[remaining_idx])
                ])
            
            # Add custom comparisons
            for comp in custom_comparisons:
                try:
                    col = df1.columns[comp['index']]
                    columns_to_compare.append((comp['name'], col))
                except:
                    pass
            
            # Compare all specified columns
            for code in common_codes:
                row1 = df1[df1[code_column] == code].iloc[0]
                row2 = df2[df2[code_column] == code].iloc[0]
                
                diff_fields = {}
                
                for field_name, col in columns_to_compare:
                    if row1[col] != row2[col]:
                        diff_fields[field_name] = (row1[col], row2[col])
                
                if diff_fields:
                    result['differences'].append({
                        'code': code,
                        'description': row1[description_column],
                        'fields': diff_fields
                    })
        
        return result
    
    def open_result_window(self, result):
        """Open a new window to display comparison results."""
        result_window = tk.Toplevel(self.root)
        result_window.title("Comparison Results")
        result_window.geometry("780x620")
        result_window.configure(bg="#1e1e1e")
        
        # Store mousewheel binding
        mousewheel_bind_id = None
        
        def on_result_closing():
            """Clean up before closing result window."""
            nonlocal mousewheel_bind_id
            if mousewheel_bind_id:
                result_window.unbind("<MouseWheel>")
            result_window.destroy()
        
        result_window.protocol("WM_DELETE_WINDOW", on_result_closing)

        style = ttk.Style()
        bg = "#1e1e1e"
        card_bg = "#2d2d30"
        text_color = "#F3F4F6"

        style.configure("ResultRoot.TFrame", background=bg)
        style.configure("ResultCard.TFrame", background=card_bg)
        style.configure(
            "ResultTitle.TLabel",
            background=bg,
            foreground=text_color,
            font=("Segoe UI", 14, "bold")
        )

        main_frame = ttk.Frame(result_window, style="ResultRoot.TFrame", padding=15)
        main_frame.pack(fill="both", expand=True)

        title_label = ttk.Label(
            main_frame,
            text="Comparison Results",
            style="ResultTitle.TLabel"
        )
        title_label.pack(anchor="w", pady=(0, 8))

        card = ttk.Frame(main_frame, style="ResultCard.TFrame", padding=10)
        card.pack(fill="both", expand=True)

        frame = ttk.Frame(card, style="ResultCard.TFrame")
        frame.pack(fill="both", expand=True)
        
        text = tk.Text(
            frame,
            wrap="word",
            font=("Consolas", 10),
            bg=card_bg,
            fg=text_color,
            insertbackground=text_color,
            borderwidth=0,
            relief="flat",
            padx=10,
            pady=10
        )
        text.pack(side="left", fill="both", expand=True)
        
        # Mouse wheel scrolling with error handling
        def on_mousewheel(event):
            try:
                if text.winfo_exists():
                    text.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        mousewheel_bind_id = text.bind("<MouseWheel>", on_mousewheel)
        
        text.insert("end", "=" * 70 + "\n")
        text.insert("end", f"File 1 Row Count: {result['row_count_1']}\n")
        text.insert("end", f"File 2 Row Count: {result['row_count_2']}\n")
        text.insert(
            "end",
            f"Difference: {abs(result['row_count_1'] - result['row_count_2'])} rows\n"
        )
        text.insert("end", "=" * 70 + "\n\n")
        
        # Missing codes (in file1 but not in file2)
        if result['missing_codes']:
            text.insert(
                "end",
                f"{len(result['missing_codes'])} Products in File 1 but NOT in File 2:\n",
                "header"
            )
            text.insert("end", "-" * 70 + "\n")
            for code, description in result['missing_codes']:
                text.insert("end", f"  • Code: {code} - {description}\n")
            text.insert("end", "\n")
        
        # Extra codes (in file2 but not in file1)
        if result['extra_codes']:
            text.insert(
                "end",
                f"{len(result['extra_codes'])} Products in File 2 but NOT in File 1:\n",
                "header"
            )
            text.insert("end", "-" * 70 + "\n")
            for code, description in result['extra_codes']:
                text.insert("end", f"  • Code: {code} - {description}\n")
            text.insert("end", "\n")
        
        # Differences in detailed settings
        if result['differences']:
            text.insert(
                "end",
                f"{len(result['differences'])} Products with Detailed Differences:\n",
                "header"
            )
            text.insert("end", "-" * 70 + "\n")
            for item in result['differences']:
                text.insert("end", f"  • Code: {item['code']} - {item['description']}\n")
                for field_name, (old, new) in item['fields'].items():
                    text.insert("end", f"    {field_name}: {old} → {new}\n")
                text.insert("end", "\n")
        
        if (
            not result['missing_codes']
            and not result['extra_codes']
            and not result['differences']
        ):
            text.insert("end", "No differences found. Files are identical!\n", "header")
        
        text.tag_config("header", font=("Segoe UI", 11, "bold"), foreground="#60A5FA")
        text.config(state="disabled")

        bottom_frame = ttk.Frame(main_frame, style="ResultRoot.TFrame")
        bottom_frame.pack(fill="x", pady=(8, 0))

        ttk.Button(
            bottom_frame,
            text="Close",
            style="Danger.TButton",
            command=on_result_closing
        ).pack(side="right")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparisonApp(root)
    root.mainloop()
