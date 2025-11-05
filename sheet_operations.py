"""
Sheet Operations Module
Handles multi-sheet Excel file operations including reading sheets,
managing formula fields across sheets, and integrating with Schedule Properties
"""

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import os

class SheetOperations:
    def __init__(self, editor_instance):
        self.editor = editor_instance
        self.available_sheets = {}  # Dict of {sheet_name: DataFrame}
        self.current_sheet = None
        self.sheet_names = []
        
    def get_sheet_names(self, file_path):
        """Get all sheet names from an Excel file"""
        try:
            if file_path.endswith('.xlsx'):
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            else:
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
            
            self.sheet_names = excel_file.sheet_names
            excel_file.close()
            return self.sheet_names
        except Exception as e:
            messagebox.showerror(self.editor.tr("Error"), f"Failed to read sheet names:\n{str(e)}")
            return []
    
    def import_file_with_sheet_selection(self, file_path=None):
        """Import Excel file with sheet selection dialog"""
        from tkinter import filedialog
        
        # If file_path not provided, ask user to select file
        if not file_path:
            file_path = filedialog.askopenfilename(
                title=self.editor.tr("Select Excel File"),
                filetypes=[
                    (self.editor.tr("Excel files"), "*.xlsx *.xls"),
                    (self.editor.tr("All files"), "*.*")
                ]
            )
        
        if not file_path:
            return
            
        # Get available sheets
        sheet_names = self.get_sheet_names(file_path)
        if not sheet_names:
            return
            
        # If only one sheet, import it directly
        if len(sheet_names) == 1:
            self.load_sheet(file_path, sheet_names[0])
            return
            
        # Show sheet selection dialog
        self.show_sheet_selection_dialog(file_path, sheet_names)
    
    def show_sheet_selection_dialog(self, file_path, sheet_names):
        """Show dialog to select which sheets to load"""
        print(f"DEBUG: Creating dialog for sheets: {sheet_names}")
        
        # Simple message box for now to test
        import tkinter.messagebox as msgbox
        
        if len(sheet_names) == 1:
            # Single sheet - load directly
            self.load_sheet(file_path, sheet_names[0])
            return
        
        # Multiple sheets - show simple dialog
        sheet_list = "\n".join([f"- {name}" for name in sheet_names])
        result = msgbox.askyesno(
            "Multiple Sheets Found", 
            f"Found {len(sheet_names)} sheets:\n{sheet_list}\n\nLoad all sheets?\n\n(Yes = Load All, No = Load First Sheet Only)"
        )
        
        if result:  # Yes - load all sheets
            self.load_multiple_sheets(file_path, sheet_names, sheet_names[0])
        else:  # No - load first sheet only  
            self.load_sheet(file_path, sheet_names[0])
        
        return  # Skip the complex dialog for now
        
        # Original complex dialog code (commented out)
        dialog = tk.Toplevel(self.editor.root)
        dialog.title(self.editor.tr("Select Sheets"))
        dialog.geometry("450x400")
        dialog.resizable(True, True)
        dialog.transient(self.editor.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (
            self.editor.root.winfo_rootx() + 100,
            self.editor.root.winfo_rooty() + 100
        ))
        
        # Instructions
        label = ttk.Label(dialog, text=f"Found {len(sheet_names)} sheet(s). Select sheets to load:", font=("Arial", 11, "bold"))
        label.pack(pady=(10, 5))
        
        # Sheet selection frame - simplified approach
        sheet_frame = ttk.LabelFrame(dialog, text="Available Sheets", padding="10")
        sheet_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Sheet checkboxes - direct approach without canvas
        sheet_vars = {}
        print(f"DEBUG: Creating checkboxes for {len(sheet_names)} sheets")
        for i, sheet_name in enumerate(sheet_names):
            print(f"DEBUG: Creating checkbox for sheet: '{sheet_name}'")
            var = tk.BooleanVar(value=(i == 0))  # Select first sheet by default
            sheet_vars[sheet_name] = var
            
            cb = ttk.Checkbutton(
                sheet_frame, 
                text=f"Sheet: {sheet_name}", 
                variable=var,
                font=("Arial", 10)
            )
            cb.pack(anchor="w", pady=3, padx=5)
            print(f"DEBUG: Checkbox created and packed for '{sheet_name}'")
        
        # Primary sheet selection
        primary_frame = ttk.LabelFrame(dialog, text=self.editor.tr("Primary Sheet (for main view)"))
        primary_frame.pack(fill="x", padx=10, pady=5)
        
        primary_var = tk.StringVar(value=sheet_names[0])
        primary_combo = ttk.Combobox(primary_frame, textvariable=primary_var, values=sheet_names, state="readonly")
        primary_combo.pack(fill="x", padx=5, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        def load_selected_sheets():
            selected_sheets = [name for name, var in sheet_vars.items() if var.get()]
            if not selected_sheets:
                messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("Please select at least one sheet."))
                return
                
            primary_sheet = primary_var.get()
            if primary_sheet not in selected_sheets:
                messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("Primary sheet must be selected."))
                return
                
            # Load all selected sheets
            self.load_multiple_sheets(file_path, selected_sheets, primary_sheet)
            dialog.destroy()
        
        def select_all():
            for var in sheet_vars.values():
                var.set(True)
        
        def select_none():
            for var in sheet_vars.values():
                var.set(False)
        
        ttk.Button(button_frame, text=self.editor.tr("Select All"), command=select_all).pack(side="left", padx=5)
        ttk.Button(button_frame, text=self.editor.tr("Select None"), command=select_none).pack(side="left", padx=5)
        ttk.Button(button_frame, text=self.editor.tr("Cancel"), command=dialog.destroy).pack(side="right", padx=5)
        ttk.Button(button_frame, text=self.editor.tr("Load Sheets"), command=load_selected_sheets).pack(side="right", padx=5)
    
    def load_multiple_sheets(self, file_path, sheet_names, primary_sheet):
        """Load multiple sheets from Excel file"""
        try:
            self.available_sheets = {}
            
            for sheet_name in sheet_names:
                try:
                    if file_path.endswith('.xlsx'):
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=self.editor.header_row)
                    else:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd', header=self.editor.header_row)
                    
                    self.available_sheets[sheet_name] = df
                except Exception as sheet_error:
                    messagebox.showerror(self.editor.tr("Error"), f"Failed to load sheet '{sheet_name}':\n{str(sheet_error)}")
                    continue
            
            # Set primary sheet as current working data
            self.current_sheet = primary_sheet
            self.editor.original_df = self.available_sheets[primary_sheet].copy()
            self.editor.df = self.editor.original_df.copy()
            self.editor.visible_columns = list(self.editor.df.columns)
            
            # Clear existing data
            self.editor.filtered_df = None
            self.editor.active_filters = {}
            self.editor.formula_fields = {}
            
            # Update file info
            self.editor.current_file = file_path
            self.editor.modified = False
            self.editor.file_ops.update_file_info()
            self.editor.filter_ops.update_filter_display()
            self.editor.update_header_display()
            self.editor.data_ops.populate_treeview()
            
            # Update status with sheet info
            sheet_count = len(self.available_sheets)
            self.editor.status_var.set(f"Loaded {sheet_count} sheets. Current: {primary_sheet}")
            
            # Add sheet switcher to the interface
            self.add_sheet_switcher()
            
        except Exception as e:
            messagebox.showerror(self.editor.tr("Error"), f"Failed to load sheets:\n{str(e)}")
    
    def load_sheet(self, file_path, sheet_name):
        """Load a single sheet"""
        try:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=self.editor.header_row)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd', header=self.editor.header_row)
            
            self.available_sheets = {sheet_name: df}
            self.current_sheet = sheet_name
            
            # Set as current working data
            self.editor.original_df = df.copy()
            self.editor.df = self.editor.original_df.copy()
            self.editor.visible_columns = list(self.editor.df.columns)
            
            # Clear existing data
            self.editor.filtered_df = None
            self.editor.active_filters = {}
            self.editor.formula_fields = {}
            
            # Update file info
            self.editor.current_file = file_path
            self.editor.modified = False
            self.editor.file_ops.update_file_info()
            self.editor.filter_ops.update_filter_display()
            self.editor.update_header_display()
            self.editor.data_ops.populate_treeview()
            
            self.editor.status_var.set(f"Loaded sheet: {sheet_name}")
            
        except Exception as e:
            messagebox.showerror(self.editor.tr("Error"), f"Failed to load sheet {sheet_name}:\n{str(e)}")
    
    def add_sheet_switcher(self):
        """Add sheet switcher to the main interface"""
        if len(self.available_sheets) <= 1:
            return
            
        # Check if sheet switcher already exists and remove it
        if hasattr(self.editor, 'sheet_switcher_frame'):
            self.editor.sheet_switcher_frame.destroy()
        
        # Find the button frame by searching through the widgets
        # The button frame should be in main_frame
        main_frame = None
        for child in self.editor.root.winfo_children():
            if isinstance(child, ttk.Frame):
                main_frame = child
                break
        
        if main_frame:
            # Look for the button frame (should be the first frame with buttons)
            button_frame = None
            for child in main_frame.winfo_children():
                if isinstance(child, ttk.Frame) and len(child.winfo_children()) > 0:
                    # Check if it contains buttons
                    has_buttons = any(isinstance(grandchild, ttk.Button) for grandchild in child.winfo_children())
                    if has_buttons:
                        button_frame = child
                        break
            
            if button_frame:
                # Add sheet switcher to the right side of button frame
                ttk.Label(button_frame, text="Sheet:").pack(side="right", padx=(10, 5))
                
                self.editor.current_sheet_var = tk.StringVar(value=self.current_sheet)
                sheet_combo = ttk.Combobox(
                    button_frame, 
                    textvariable=self.editor.current_sheet_var,
                    values=list(self.available_sheets.keys()),
                    state="readonly",
                    width=15
                )
                sheet_combo.pack(side="right", padx=(0, 5))
                sheet_combo.bind("<<ComboboxSelected>>", self.switch_sheet)
                
                # Store reference for potential cleanup
                self.editor.sheet_switcher_combo = sheet_combo
    
    def switch_sheet(self, event=None):
        """Switch to a different sheet"""
        new_sheet = self.editor.current_sheet_var.get()
        if new_sheet == self.current_sheet:
            return
            
        if new_sheet in self.available_sheets:
            # Save current sheet's data back to available_sheets (including any formula fields)
            if self.current_sheet:
                self.available_sheets[self.current_sheet] = self.editor.df.copy()
            
            # Switch to new sheet
            self.current_sheet = new_sheet
            self.editor.original_df = self.available_sheets[new_sheet].copy()
            self.editor.df = self.editor.original_df.copy()
            self.editor.visible_columns = list(self.editor.df.columns)
            
            # Clear filters and refresh display
            self.editor.filtered_df = None
            self.editor.active_filters = {}
            
            self.editor.filter_ops.update_filter_display()
            self.editor.update_header_display()
            self.editor.data_ops.populate_treeview()
            
            self.editor.status_var.set(f"Switched to sheet: {new_sheet}")
    
    def get_available_sheets_for_formula(self):
        """Get list of available sheets for formula creation"""
        return list(self.available_sheets.keys())
    
    def get_sheet_columns(self, sheet_name):
        """Get column names from a specific sheet"""
        if sheet_name in self.available_sheets:
            return list(self.available_sheets[sheet_name].columns)
        return []
    
    def get_sheet_data(self, sheet_name):
        """Get DataFrame for a specific sheet"""
        return self.available_sheets.get(sheet_name, None)
    
    def create_cross_sheet_formula(self, target_sheet, formula_field_name, formula_expression):
        """Create a formula that can reference fields from other sheets using the main formula engine"""
        if target_sheet not in self.available_sheets:
            messagebox.showerror(self.editor.tr("Error"), f"Sheet '{target_sheet}' not found")
            return False
        
        try:
            # Save current sheet state
            original_sheet = self.current_sheet
            original_df = self.editor.original_df.copy() if self.editor.original_df is not None else None
            original_visible_columns = self.editor.visible_columns.copy()
            original_formula_fields = self.editor.formula_fields.copy()
            
            # Temporarily switch to target sheet
            self.current_sheet = target_sheet
            self.editor.original_df = self.available_sheets[target_sheet].copy()
            self.editor.df = self.editor.original_df.copy()
            self.editor.visible_columns = list(self.editor.df.columns)
            
            # Use the main formula engine to validate and calculate
            if not self.editor.formula_ops.validate_formula(formula_expression):
                # Restore original state
                self.current_sheet = original_sheet
                self.editor.original_df = original_df
                self.editor.visible_columns = original_visible_columns
                self.editor.formula_fields = original_formula_fields
                return False
            
            # Add formula field to target sheet
            self.editor.formula_fields[formula_field_name] = {
                'expression': formula_expression,
                'type': 'Number'  # Default to Number type
            }
            
            # Calculate the formula field using the main formula engine
            self.editor.formula_ops.calculate_formula_field(formula_field_name)
            
            # Save the result to available_sheets
            self.available_sheets[target_sheet] = self.editor.original_df.copy()
            
            # Restore original state
            if original_sheet != target_sheet:
                self.current_sheet = original_sheet
                self.editor.original_df = original_df
                self.editor.df = original_df.copy() if original_df is not None else None
                self.editor.visible_columns = original_visible_columns
                self.editor.formula_fields = original_formula_fields
                
                # Refresh display to show original sheet
                self.editor.data_ops.populate_treeview()
            else:
                # Target is current sheet, keep the updated data
                self.editor.df = self.editor.original_df.copy()
                if formula_field_name not in self.editor.visible_columns:
                    self.editor.visible_columns.append(formula_field_name)
                self.editor.data_ops.populate_treeview()
            
            messagebox.showinfo(self.editor.tr("Success"), f"Formula field '{formula_field_name}' created successfully in sheet '{target_sheet}'")
            self.editor.modified = True
            return True
            
        except Exception as e:
            # Restore original state on error
            if original_df is not None:
                self.current_sheet = original_sheet
                self.editor.original_df = original_df
                self.editor.visible_columns = original_visible_columns
                self.editor.formula_fields = original_formula_fields
            
            messagebox.showerror(self.editor.tr("Error"), f"Failed to create formula:\n{str(e)}")
            return False
    
    def get_cross_sheet_fields_for_schedule_properties(self):
        """Get all fields from all sheets for Schedule Properties"""
        all_fields = {}
        for sheet_name, df in self.available_sheets.items():
            all_fields[sheet_name] = list(df.columns)
        return all_fields
    
    def save_all_sheets(self, file_path):
        """Save all sheets to Excel file"""
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for sheet_name, df in self.available_sheets.items():
                    # Update current sheet data if it was modified
                    if sheet_name == self.current_sheet:
                        df = self.editor.df.copy()
                    
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            self.editor.modified = False
            self.editor.status_var.set(f"All sheets saved to: {os.path.basename(file_path)}")
            return True
            
        except Exception as e:
            messagebox.showerror(self.editor.tr("Error"), f"Failed to save sheets:\n{str(e)}")
            return False