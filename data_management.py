"""
Data Management Module
Handles data display, editing, and row/column operations for the XLS Editor
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd


class DataManagement:
    def __init__(self, editor_instance):
        self.editor = editor_instance
    
    def populate_treeview(self):
        """Populate the treeview with DataFrame data"""
        # Use filtered data if filters are active, otherwise use working data
        display_df = self.editor.filtered_df if self.editor.filtered_df is not None else self.editor.df
        
        if display_df is None:
            return
        
        # Only show visible columns
        if self.editor.visible_columns:
            visible_cols = [col for col in self.editor.visible_columns if col in display_df.columns]
            if visible_cols:
                display_df = display_df[visible_cols]
            
        # Clear existing data
        for item in self.editor.tree.get_children():
            self.editor.tree.delete(item)
            
        # Configure columns
        columns = list(display_df.columns)
        self.editor.tree['columns'] = columns
        self.editor.tree['show'] = 'tree headings'
        
        # Configure column widths and headings
        self.editor.tree.column('#0', width=50, minwidth=50)
        self.editor.tree.heading('#0', text='Row')
        
        for col in columns:
            self.editor.tree.column(col, width=100, minwidth=80)
            self.editor.tree.heading(col, text=str(col))
            
        # Insert data
        for index, row in display_df.iterrows():
            values = [str(val) if pd.notna(val) else '' for val in row]
            # Use original DataFrame index for editing purposes
            original_index = index if self.editor.filtered_df is None else display_df.index[display_df.index == index][0]
            self.editor.tree.insert('', 'end', text=str(original_index), values=values)
    
    def on_cell_double_click(self, event):
        """Handle double-click on a cell for editing"""
        if self.editor.df is None:
            return
            
        item = self.editor.tree.selection()[0]
        column = self.editor.tree.identify_column(event.x)
        
        if column == '#0':  # Row number column
            return
            
        # Get column index
        col_index = int(column.replace('#', '')) - 1
        if col_index >= len(self.editor.df.columns):
            return
            
        # Get row index
        row_index = int(self.editor.tree.item(item, 'text'))
        
        # Get current value
        current_value = self.editor.df.iloc[row_index, col_index]
        if pd.isna(current_value):
            current_value = ''
        else:
            current_value = str(current_value)
            
        # Create edit dialog
        self.edit_cell(row_index, col_index, current_value)
    
    def edit_cell(self, row_index, col_index, current_value):
        """Open dialog to edit cell value"""
        dialog = tk.Toplevel(self.editor.root)
        dialog.title("Edit Cell")
        dialog.geometry("300x150")
        dialog.transient(self.editor.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (300 // 2)
        y = (dialog.winfo_screenheight() // 2) - (150 // 2)
        dialog.geometry(f"300x150+{x}+{y}")
        
        # Create widgets
        ttk.Label(dialog, text=f"Edit cell [{row_index}, {self.editor.df.columns[col_index]}]:").pack(pady=10)
        
        entry_var = tk.StringVar(value=current_value)
        entry = ttk.Entry(dialog, textvariable=entry_var, width=30)
        entry.pack(pady=5)
        entry.focus()
        entry.select_range(0, tk.END)
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def save_edit():
            new_value = entry_var.get()
            if new_value == '':
                self.editor.df.iloc[row_index, col_index] = None
            else:
                # Try to convert to appropriate type
                try:
                    # Try numeric conversion
                    if '.' in new_value:
                        new_value = float(new_value)
                    else:
                        new_value = int(new_value)
                except ValueError:
                    # Keep as string
                    pass
                self.editor.df.iloc[row_index, col_index] = new_value

            self.editor.modified = True
            self.editor.file_ops.update_file_info()
            self.populate_treeview()
            self.editor.status_var.set("Cell updated")
            dialog.destroy()
            
        ttk.Button(button_frame, text="Save", command=save_edit).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Bind Enter key to save
        entry.bind('<Return>', lambda e: save_edit())
    
    def add_row(self):
        """Add a new row to the DataFrame"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        # Add empty row
        new_row = pd.Series([None] * len(self.editor.df.columns), index=self.editor.df.columns)
        self.editor.df = pd.concat([self.editor.df, new_row.to_frame().T], ignore_index=True)
        
        self.editor.modified = True
        self.editor.file_ops.update_file_info()
        self.populate_treeview()
        self.editor.status_var.set("Row added")
    
    def delete_row(self):
        """Delete selected row"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        selection = self.editor.tree.selection()
        if not selection:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("Please select a row to delete."))
            return
            
        row_index = int(self.editor.tree.item(selection[0], 'text'))
        
        if messagebox.askyesno("Confirm", f"Delete row {row_index}?"):
            self.editor.df = self.editor.df.drop(index=row_index).reset_index(drop=True)
            self.editor.modified = True
            self.editor.file_ops.update_file_info()
            self.populate_treeview()
            self.editor.status_var.set("Row deleted")
    
    def add_column(self):
        """Add a new column to the DataFrame"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        # Get column name
        column_name = simpledialog.askstring("Add Column", "Enter column name:")
        if column_name and column_name not in self.editor.df.columns:
            self.editor.df[column_name] = None
            self.editor.modified = True
            self.editor.file_ops.update_file_info()
            self.populate_treeview()
            self.editor.status_var.set(f"Column '{column_name}' added")
        elif column_name in self.editor.df.columns:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("Column name already exists."))
    
    def delete_column(self):
        """Delete a column from the DataFrame"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        # Get column to delete
        columns = list(self.editor.df.columns)
        if not columns:
            return
            
        # Create selection dialog
        dialog = tk.Toplevel(self.editor.root)
        dialog.title("Delete Column")
        dialog.geometry("250x150")
        dialog.transient(self.editor.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Select column to delete:").pack(pady=10)
        
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(dialog, textvariable=column_var, values=columns, state="readonly")
        column_combo.pack(pady=5)
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def delete_selected():
            if column_var.get():
                if messagebox.askyesno("Confirm", f"Delete column '{column_var.get()}'?"):
                    self.editor.df = self.editor.df.drop(columns=[column_var.get()])
                    self.editor.modified = True
                    self.editor.file_ops.update_file_info()
                    self.populate_treeview()
                    self.editor.status_var.set(f"Column '{column_var.get()}' deleted")
                    dialog.destroy()
                    
        ttk.Button(button_frame, text="Delete", command=delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)