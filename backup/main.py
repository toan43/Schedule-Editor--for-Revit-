"""
XLS File Editor with Filtering - Modular Version
Main application file that coordinates all modules

This file contains the main XLSEditor class that initializes and coordinates
all the separate modules for different functionalities.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
from typing import Optional

# Import our modules
from translation_manager import TranslationManager
from file_operations import FileOperations
from data_management import DataManagement
from filter_operations import FilterOperations
from formula_operations import FormulaOperations
from schedule_properties import ScheduleProperties


class XLSEditor:
    def __init__(self, root):
        self.root = root
        
        # Initialize translation manager first
        self.translation_manager = TranslationManager()
        
        # Set up window
        self.root.title(self.tr("XLS File Editor with Filtering"))
        self.root.geometry("1200x800")
        
        # Data-related attributes
        self.current_file: Optional[str] = None
        self.df: Optional[pd.DataFrame] = None
        self.original_df: Optional[pd.DataFrame] = None  # Keep original data intact
        self.filtered_df: Optional[pd.DataFrame] = None
        self.active_filters = {}
        self.modified = False
        self.header_row = 0  # Default to first row as header
        self.visible_columns = []  # Track which columns are visible in schedule
        self.sort_settings = []  # List of sort criteria
        self.group_settings = []  # List of grouping criteria
        self.appearance_settings = {
            'show_headers': True,
            'show_title': True,
            'grid_lines': True,
            'outline': True
        }
        
        # Formula-related attributes
        self.formula_fields = {}  # Dictionary to store formula fields and their expressions
        self.formula_templates = {}  # Dictionary to store saved formula templates
        
        # Initialize operation modules
        self.file_ops = FileOperations(self)
        self.data_ops = DataManagement(self)
        self.filter_ops = FilterOperations(self)
        self.formula_ops = FormulaOperations(self)
        self.schedule_props = ScheduleProperties(self)
        
        # Create GUI
        self.create_menu()
        self.create_widgets()
    
    # Translation methods (delegated to translation manager)
    def tr(self, text):
        """Translate text based on current language"""
        return self.translation_manager.tr(text)
    
    def change_language(self, language_code):
        """Change the application language"""
        message = self.translation_manager.change_language(language_code)
        self.refresh_interface()
        messagebox.showinfo(self.tr("Success"), message)
    
    def refresh_interface(self):
        """Refresh the entire interface with new language"""
        # Update window title
        self.root.title(self.tr("XLS File Editor with Filtering"))
        
        # Clear current menu
        self.root.config(menu=tk.Menu(self.root))
        
        # Clear current widgets
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # Recreate interface
        self.create_menu()
        self.create_widgets()
        
        # Refresh data if available
        if self.df is not None:
            self.data_ops.populate_treeview()
            self.file_ops.update_file_info()
            self.filter_ops.update_filter_display()
            self.update_header_display()
    
    def create_menu(self):
        """Create the application menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("File"), menu=file_menu)
        file_menu.add_command(label=self.tr("Import XLS File"), command=self.file_ops.import_file)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("Save"), command=self.file_ops.save_file)
        file_menu.add_command(label=self.tr("Save As"), command=self.file_ops.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("Exit"), command=self.file_ops.on_closing)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("Edit"), menu=edit_menu)
        edit_menu.add_command(label=self.tr("Add Row"), command=self.data_ops.add_row)
        edit_menu.add_command(label=self.tr("Delete Row"), command=self.data_ops.delete_row)
        edit_menu.add_command(label=self.tr("Add Column"), command=self.data_ops.add_column)
        edit_menu.add_command(label=self.tr("Delete Column"), command=self.data_ops.delete_column)
        
        # Filter menu
        filter_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("Filter"), menu=filter_menu)
        filter_menu.add_command(label=self.tr("Clear All Filters"), command=self.filter_ops.clear_all_filters)
        filter_menu.add_command(label=self.tr("Manage Filters"), command=self.filter_ops.manage_filters)
        filter_menu.add_separator()
        filter_menu.add_command(label=self.tr("Set Header Row"), command=self.set_header_row)
        
        # Schedule menu (Revit-like features)
        schedule_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("Schedule"), menu=schedule_menu)
        schedule_menu.add_command(label=self.tr("Schedule Properties"), command=self.schedule_props.open_schedule_properties)
        schedule_menu.add_separator()
        schedule_menu.add_command(label=self.tr("Add Parameter"), command=self.add_parameter)
        schedule_menu.add_command(label=self.tr("Remove Parameter"), command=self.remove_parameter)
        
        # Language menu
        language_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("Language"), menu=language_menu)
        language_menu.add_command(label=self.tr("English"), command=lambda: self.change_language("en"))
        language_menu.add_command(label=self.tr("Vietnamese"), command=lambda: self.change_language("vi"))
    
    def create_widgets(self):
        """Create the main application widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # File info frame
        info_frame = ttk.LabelFrame(main_frame, text=self.tr("File Information"), padding="5")
        info_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        info_frame.columnconfigure(1, weight=1)
        
        ttk.Label(info_frame, text=self.tr("Current File:")).grid(row=0, column=0, sticky=tk.W)
        self.file_label = ttk.Label(info_frame, text=self.tr("No file loaded"), foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        # Control buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(button_frame, text=self.tr("Import XLS File"), command=self.file_ops.import_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text=self.tr("Save"), command=self.file_ops.save_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text=self.tr("Save As"), command=self.file_ops.save_as_file).pack(side=tk.LEFT, padx=(0, 5))
        
        # Filter control frame
        filter_frame = ttk.LabelFrame(main_frame, text="Filters", padding="5")
        filter_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        filter_frame.columnconfigure(1, weight=1)
        
        ttk.Button(filter_frame, text="Clear All", command=self.filter_ops.clear_all_filters).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(filter_frame, text="Manage", command=self.filter_ops.manage_filters).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(filter_frame, text="Header Row", command=self.set_header_row).grid(row=0, column=2, padx=(0, 5))
        ttk.Button(filter_frame, text="Schedule Properties", command=self.schedule_props.open_schedule_properties).grid(row=0, column=3, padx=(0, 5))
        
        # Active filters display
        self.filter_display = ttk.Label(filter_frame, text="No filters active", foreground="gray")
        self.filter_display.grid(row=0, column=4, sticky=(tk.W, tk.E), padx=(10, 0))
        
        # Header row display
        self.header_display = ttk.Label(filter_frame, text="Header: Row 1", foreground="blue", font=("Arial", 8))
        self.header_display.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(5, 0))
        
        # Data display frame
        data_frame = ttk.LabelFrame(main_frame, text="Data Editor", padding="5")
        data_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for data display
        self.tree_frame = ttk.Frame(data_frame)
        self.tree_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.tree_frame.columnconfigure(0, weight=1)
        self.tree_frame.rowconfigure(0, weight=1)
        
        # Treeview with scrollbars
        self.tree = ttk.Treeview(self.tree_frame)
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Bind double-click for editing
        self.tree.bind('<Double-1>', self.data_ops.on_cell_double_click)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
    
    # Header management
    def update_header_display(self):
        """Update the header row display"""
        self.header_display.config(text=f"Header: Row {self.header_row + 1}")
    
    def set_header_row(self):
        """Set which row to use as headers"""
        if self.current_file is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        # Create header row selection dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Set Header Row")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (300 // 2)
        y = (dialog.winfo_screenheight() // 2) - (200 // 2)
        dialog.geometry(f"300x200+{x}+{y}")
        
        # Current header info
        ttk.Label(dialog, text=f"Current header row: {self.header_row + 1}", 
                 font=("Arial", 10, "bold")).pack(pady=10)
        
        # Row selection
        ttk.Label(dialog, text="Select which row contains the column headers:").pack(pady=5)
        
        row_var = tk.IntVar(value=self.header_row + 1)
        
        # Radio buttons for row selection
        row_frame = ttk.Frame(dialog)
        row_frame.pack(pady=10)
        
        ttk.Radiobutton(row_frame, text="Row 1 (first row)", variable=row_var, value=1).pack(anchor=tk.W)
        ttk.Radiobutton(row_frame, text="Row 2 (second row)", variable=row_var, value=2).pack(anchor=tk.W)
        ttk.Radiobutton(row_frame, text="Row 3 (third row)", variable=row_var, value=3).pack(anchor=tk.W)
        
        # Custom row entry
        custom_frame = ttk.Frame(dialog)
        custom_frame.pack(pady=5)
        
        ttk.Radiobutton(custom_frame, text="Custom row:", variable=row_var, value=0).pack(side=tk.LEFT)
        custom_var = tk.StringVar()
        custom_entry = ttk.Entry(custom_frame, textvariable=custom_var, width=5)
        custom_entry.pack(side=tk.LEFT, padx=5)
        
        def apply_header_change():
            """Apply the header row change"""
            new_header_row = row_var.get() - 1  # Convert to 0-based index
            
            if row_var.get() == 0:  # Custom row
                try:
                    new_header_row = int(custom_var.get()) - 1
                    if new_header_row < 0:
                        messagebox.showerror("Error", "Row number must be 1 or greater.")
                        return
                except ValueError:
                    messagebox.showerror("Error", "Please enter a valid row number.")
                    return
            
            if new_header_row == self.header_row:
                dialog.destroy()
                return
                
            # Clear existing filters since column names might change
            if self.active_filters:
                result = messagebox.askyesno("Clear Filters", 
                    "Changing the header row will clear all active filters. Continue?")
                if not result:
                    return
            
            self.header_row = new_header_row
            self.active_filters = {}
            self.filtered_df = None
            
            # Reload the file with new header row
            try:
                if self.current_file.endswith('.xlsx'):
                    self.original_df = pd.read_excel(self.current_file, engine='openpyxl', header=self.header_row)
                else:
                    self.original_df = pd.read_excel(self.current_file, engine='xlrd', header=self.header_row)
                
                # Reset visible columns and working dataframe
                self.df = self.original_df.copy()
                self.visible_columns = list(self.df.columns)
                
                self.data_ops.populate_treeview()
                self.filter_ops.update_filter_display()
                self.update_header_display()
                self.status_var.set(f"Header row changed to row {self.header_row + 1}")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to reload file with new header row:\n{str(e)}")
                # Revert to previous header row
                self.header_row = 0
                self.update_header_display()
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Apply", command=apply_header_change).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    # Legacy methods for compatibility
    def add_parameter(self):
        """Add a new parameter/column"""
        self.data_ops.add_column()
        
    def remove_parameter(self):
        """Remove a parameter/column"""
        self.data_ops.delete_column()


def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = XLSEditor(root)
    
    # Handle window closing
    root.protocol("WM_DELETE_WINDOW", app.file_ops.on_closing)
    
    # Start the application
    root.mainloop()


if __name__ == "__main__":
    main()