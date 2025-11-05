"""
File Operations Module
Handles all file import, export, and management operations for the XLS Editor
"""

import pandas as pd
import os
from tkinter import filedialog, messagebox


class FileOperations:
    def __init__(self, editor_instance):
        self.editor = editor_instance
    
    def smart_import_file(self):
        """Smart import - automatically detects multi-sheet files and offers sheet selection"""
        file_path = filedialog.askopenfilename(
            title=self.editor.tr("Select XLS File"),
            filetypes=[
                (self.editor.tr("Excel files"), "*.xlsx *.xls"),
                (self.editor.tr("All files"), "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            # Check how many sheets the file has
            if file_path.endswith('.xlsx'):
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            else:
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
            
            sheet_names = excel_file.sheet_names
            
            # If only 1 sheet, use simple import
            if len(sheet_names) == 1:
                self._simple_import(file_path)
            else:
                # Multiple sheets - use sheet selection dialog
                if hasattr(self.editor, 'sheet_ops'):
                    self.editor.sheet_ops.import_file_with_sheet_selection(file_path)
                else:
                    # Fallback to simple import if sheet_ops not available
                    self._simple_import(file_path)
                    
        except Exception as e:
            messagebox.showerror(self.editor.tr("Error"), f"{self.editor.tr('Failed to open file')}:\n{str(e)}")
            self.editor.status_var.set(self.editor.tr("Import failed"))
    
    def _simple_import(self, file_path):
        """Internal method for simple single-sheet import"""
        try:
            # Read the Excel file with the specified header row
            if file_path.endswith('.xlsx'):
                self.editor.original_df = pd.read_excel(file_path, engine='openpyxl', header=self.editor.header_row)
            else:
                self.editor.original_df = pd.read_excel(file_path, engine='xlrd', header=self.editor.header_row)
            
            # Set working dataframe and visible columns
            self.editor.df = self.editor.original_df.copy()
            self.editor.visible_columns = list(self.editor.df.columns)
            
            # Clear formula fields since data structure might have changed
            self.editor.formula_fields = {}
            
            self.editor.current_file = file_path
            self.editor.modified = False
            self.editor.filtered_df = None
            self.editor.active_filters = {}
            self.update_file_info()
            self.editor.filter_ops.update_filter_display()
            self.editor.update_header_display()
            self.editor.data_ops.populate_treeview()
            self.editor.status_var.set(f"{self.editor.tr('File imported successfully')}: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror(self.editor.tr("Error"), f"{self.editor.tr('Failed to import file')}:\n{str(e)}")
            self.editor.status_var.set(self.editor.tr("Import failed"))
    
    def import_file(self):
        """Import an XLS file"""
        file_path = filedialog.askopenfilename(
            title=self.editor.tr("Select XLS File"),
            filetypes=[
                (self.editor.tr("Excel files"), "*.xlsx *.xls"),
                (self.editor.tr("All files"), "*.*")
            ]
        )
        
        if file_path:
            try:
                # Read the Excel file with the specified header row
                if file_path.endswith('.xlsx'):
                    self.editor.original_df = pd.read_excel(file_path, engine='openpyxl', header=self.editor.header_row)
                else:
                    self.editor.original_df = pd.read_excel(file_path, engine='xlrd', header=self.editor.header_row)
                
                # Set working dataframe and visible columns
                self.editor.df = self.editor.original_df.copy()
                self.editor.visible_columns = list(self.editor.df.columns)  # Initially all columns are visible
                
                # Clear formula fields since data structure might have changed
                self.editor.formula_fields = {}
                
                self.editor.current_file = file_path
                self.editor.modified = False
                self.editor.filtered_df = None
                self.editor.active_filters = {}
                self.update_file_info()
                self.editor.filter_ops.update_filter_display()
                self.editor.update_header_display()
                self.editor.data_ops.populate_treeview()
                self.editor.status_var.set(f"{self.editor.tr('File imported successfully')}: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror(self.editor.tr("Error"), f"{self.editor.tr('Failed to import file')}:\n{str(e)}")
                self.editor.status_var.set(self.editor.tr("Import failed"))
    
    def save_file(self):
        """Save the current file"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        if self.editor.current_file is None:
            self.save_as_file()
            return
            
        try:
            # Check if we have multiple sheets loaded
            if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                # Save all sheets
                self.editor.sheet_ops.save_all_sheets(self.editor.current_file)
            else:
                # Save single sheet
                self.editor.df.to_excel(self.editor.current_file, index=False, engine='openpyxl')
            
            self.editor.modified = False
            self.update_file_info()
            self.editor.status_var.set("File saved successfully")
        except Exception as e:
            messagebox.showerror(self.editor.tr("Error"), f"{self.editor.tr('Failed to save file')}:\n{str(e)}")
    
    def save_as_file(self):
        """Save the file with a new name"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        file_path = filedialog.asksaveasfilename(
            title=self.editor.tr("Save XLS File"),
            defaultextension=".xlsx",
            filetypes=[
                (self.editor.tr("Excel files"), "*.xlsx"),
                (self.editor.tr("All files"), "*.*")
            ]
        )
        
        if file_path:
            try:
                # Check if we have multiple sheets loaded
                if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                    # Save all sheets
                    self.editor.sheet_ops.save_all_sheets(file_path)
                else:
                    # Save single sheet
                    self.editor.df.to_excel(file_path, index=False, engine='openpyxl')
                
                self.editor.current_file = file_path
                self.editor.modified = False
                self.update_file_info()
                self.editor.status_var.set(f"File saved as: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror(self.editor.tr("Error"), f"{self.editor.tr('Failed to save file')}:\n{str(e)}")
    
    def update_file_info(self):
        """Update the file information display"""
        if self.editor.current_file:
            filename = os.path.basename(self.editor.current_file)
            if self.editor.modified:
                filename += " *"
            self.editor.file_label.config(text=filename, foreground="black")
        else:
            self.editor.file_label.config(text="No file loaded", foreground="gray")
    
    def on_closing(self):
        """Handle application closing"""
        # Save formula templates before closing
        if hasattr(self.editor, 'formula_ops'):
            self.editor.formula_ops.save_formula_templates_to_file()
        
        if self.editor.modified:
            result = messagebox.askyesnocancel(
                "Unsaved Changes", 
                "You have unsaved changes. Do you want to save before closing?"
            )
            if result is True:  # Yes, save
                self.save_file()
                if not self.editor.modified:  # Only close if save was successful
                    self.editor.root.destroy()
            elif result is False:  # No, don't save
                self.editor.root.destroy()
            # Cancel - do nothing
        else:
            self.editor.root.destroy()