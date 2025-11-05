"""
Filtering Module
Handles all filtering system functions for the XLS Editor
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd


class FilterOperations:
    def __init__(self, editor_instance):
        self.editor = editor_instance
    
    def add_filter(self):
        """Add a new filter to the data"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        # Create filter dialog
        dialog = tk.Toplevel(self.editor.root)
        dialog.title("Add Filter")
        dialog.geometry("400x300")
        dialog.transient(self.editor.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")
        
        # Column selection
        ttk.Label(dialog, text="Column:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(dialog, textvariable=column_var, values=list(self.editor.df.columns), state="readonly")
        column_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # Filter type selection
        ttk.Label(dialog, text="Filter Type:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        filter_type_var = tk.StringVar(value="equals")
        filter_types = ["equals", "not equals", "contains", "not contains", "starts with", "ends with", 
                       "greater than", "less than", "greater or equal", "less or equal", "is empty", "is not empty"]
        filter_combo = ttk.Combobox(dialog, textvariable=filter_type_var, values=filter_types, state="readonly")
        filter_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # Filter value
        ttk.Label(dialog, text="Value:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        value_var = tk.StringVar()
        value_entry = ttk.Entry(dialog, textvariable=value_var)
        value_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # Case sensitive checkbox
        case_sensitive_var = tk.BooleanVar()
        case_check = ttk.Checkbutton(dialog, text="Case sensitive", variable=case_sensitive_var)
        case_check.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Preview frame
        preview_frame = ttk.LabelFrame(dialog, text="Unique Values Preview", padding="5")
        preview_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Listbox for unique values
        preview_listbox = tk.Listbox(preview_frame)
        preview_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        preview_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=preview_listbox.yview)
        preview_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        preview_listbox.configure(yscrollcommand=preview_scroll.set)
        
        def update_preview(*args):
            """Update the preview of unique values"""
            if column_var.get():
                unique_values = self.editor.df[column_var.get()].dropna().unique()
                preview_listbox.delete(0, tk.END)
                for value in sorted(unique_values, key=str):
                    preview_listbox.insert(tk.END, str(value))
                    
        def on_preview_select(event):
            """Handle selection from preview"""
            selection = preview_listbox.curselection()
            if selection:
                value_var.set(preview_listbox.get(selection[0]))
        
        column_var.trace('w', update_preview)
        preview_listbox.bind('<Double-Button-1>', on_preview_select)
        
        # Configure grid weights
        dialog.columnconfigure(1, weight=1)
        dialog.rowconfigure(4, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        def apply_filter():
            """Apply the filter"""
            if not column_var.get():
                messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("Please select a column."))
                return
                
            filter_type = filter_type_var.get()
            if filter_type not in ["is empty", "is not empty"] and not value_var.get():
                messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("Please enter a filter value."))
                return
                
            # Store the filter
            filter_id = f"{column_var.get()}_{len(self.editor.active_filters)}"
            self.editor.active_filters[filter_id] = {
                'column': column_var.get(),
                'type': filter_type,
                'value': value_var.get(),
                'case_sensitive': case_sensitive_var.get()
            }
            
            self.apply_filters()
            self.update_filter_display()
            self.editor.status_var.set(f"Filter applied to column '{column_var.get()}'")
            dialog.destroy()
            
        ttk.Button(button_frame, text="Apply Filter", command=apply_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def apply_filters(self):
        """Apply all active filters to the DataFrame"""
        if self.editor.df is None or not self.editor.active_filters:
            self.editor.filtered_df = None
            self.editor.data_ops.populate_treeview()
            return
        
        # Debug logging
        print(f"\n[FILTER DEBUG] Starting filter application")
        print(f"  Current DataFrame: {len(self.editor.df)} rows, columns: {list(self.editor.df.columns)}")
        print(f"  Active filters: {len(self.editor.active_filters)}")
        for fname, finfo in self.editor.active_filters.items():
            print(f"    - {fname}: {finfo['column']} {finfo['type']} '{finfo['value']}'")
            
        filtered_df = self.editor.df.copy()
        
        for filter_info in self.editor.active_filters.values():
            column = filter_info['column']
            filter_type = filter_info['type']
            value = filter_info['value']
            case_sensitive = filter_info['case_sensitive']
            
            if column not in filtered_df.columns:
                print(f"  ⚠️  Column '{column}' not found in DataFrame!")
                continue
            
            print(f"  Applying filter on '{column}' ({filter_type} '{value}')...")
            print(f"    Before: {len(filtered_df)} rows")
                
            # Apply filter based on type
            if filter_type == "equals":
                if case_sensitive:
                    mask = filtered_df[column].astype(str) == value
                else:
                    mask = filtered_df[column].astype(str).str.lower() == value.lower()
            elif filter_type == "not equals":
                if case_sensitive:
                    mask = filtered_df[column].astype(str) != value
                else:
                    mask = filtered_df[column].astype(str).str.lower() != value.lower()
            elif filter_type == "contains":
                if case_sensitive:
                    mask = filtered_df[column].astype(str).str.contains(value, na=False)
                else:
                    mask = filtered_df[column].astype(str).str.lower().str.contains(value.lower(), na=False)
            elif filter_type == "not contains":
                if case_sensitive:
                    mask = ~filtered_df[column].astype(str).str.contains(value, na=False)
                else:
                    mask = ~filtered_df[column].astype(str).str.lower().str.contains(value.lower(), na=False)
            elif filter_type == "starts with":
                if case_sensitive:
                    mask = filtered_df[column].astype(str).str.startswith(value)
                else:
                    mask = filtered_df[column].astype(str).str.lower().str.startswith(value.lower())
            elif filter_type == "ends with":
                if case_sensitive:
                    mask = filtered_df[column].astype(str).str.endswith(value)
                else:
                    mask = filtered_df[column].astype(str).str.lower().str.endswith(value.lower())
            elif filter_type == "greater than":
                try:
                    mask = pd.to_numeric(filtered_df[column], errors='coerce') > float(value)
                except ValueError:
                    mask = filtered_df[column].astype(str) > value
            elif filter_type == "less than":
                try:
                    mask = pd.to_numeric(filtered_df[column], errors='coerce') < float(value)
                except ValueError:
                    mask = filtered_df[column].astype(str) < value
            elif filter_type == "greater or equal":
                try:
                    mask = pd.to_numeric(filtered_df[column], errors='coerce') >= float(value)
                except ValueError:
                    mask = filtered_df[column].astype(str) >= value
            elif filter_type == "less or equal":
                try:
                    mask = pd.to_numeric(filtered_df[column], errors='coerce') <= float(value)
                except ValueError:
                    mask = filtered_df[column].astype(str) <= value
            elif filter_type == "is empty":
                mask = filtered_df[column].isna() | (filtered_df[column].astype(str) == '')
            elif filter_type == "is not empty":
                mask = ~(filtered_df[column].isna() | (filtered_df[column].astype(str) == ''))
            else:
                print(f"  ⚠️  Unknown filter type: {filter_type}")
                continue
            
            print(f"    Mask: {mask.sum()} rows match")
            filtered_df = filtered_df[mask]
            print(f"    After: {len(filtered_df)} rows")
        
        print(f"\n  Final result: {len(filtered_df)} rows (from {len(self.editor.df)})")
        self.editor.filtered_df = filtered_df if len(filtered_df) < len(self.editor.df) else None
        print(f"  Setting filtered_df: {self.editor.filtered_df is not None}")
        self.editor.data_ops.populate_treeview()
    
    def clear_all_filters(self):
        """Clear all active filters"""
        if not self.editor.active_filters:
            messagebox.showinfo("Info", "No filters are currently active.")
            return
            
        if messagebox.askyesno("Confirm", "Clear all filters?"):
            self.editor.active_filters = {}
            self.editor.filtered_df = None
            self.apply_filters()
            self.update_filter_display()
            self.editor.status_var.set("All filters cleared")
    
    def manage_filters(self):
        """Open filter management dialog"""
        if not self.editor.active_filters:
            messagebox.showinfo("Info", "No filters are currently active.")
            return
            
        # Create filter management dialog
        dialog = tk.Toplevel(self.editor.root)
        dialog.title("Manage Filters")
        dialog.geometry("500x300")
        dialog.transient(self.editor.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"500x300+{x}+{y}")
        
        # Filter list
        ttk.Label(dialog, text="Active Filters:").pack(anchor=tk.W, padx=5, pady=5)
        
        filter_frame = ttk.Frame(dialog)
        filter_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Treeview for filters
        filter_tree = ttk.Treeview(filter_frame, columns=('Column', 'Type', 'Value'), show='headings')
        filter_tree.heading('Column', text='Column')
        filter_tree.heading('Type', text='Filter Type')
        filter_tree.heading('Value', text='Value')
        
        filter_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        filter_scroll = ttk.Scrollbar(filter_frame, orient=tk.VERTICAL, command=filter_tree.yview)
        filter_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        filter_tree.configure(yscrollcommand=filter_scroll.set)
        
        # Populate filter list
        for filter_id, filter_info in self.editor.active_filters.items():
            filter_tree.insert('', 'end', iid=filter_id, values=(
                filter_info['column'],
                filter_info['type'],
                filter_info['value'] if filter_info['value'] else '(empty)'
            ))
            
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def remove_selected():
            """Remove selected filter"""
            selection = filter_tree.selection()
            if not selection:
                messagebox.showwarning("Warning", "Please select a filter to remove.")
                return
                
            for item in selection:
                if item in self.editor.active_filters:
                    del self.editor.active_filters[item]
                filter_tree.delete(item)
                
            self.apply_filters()
            self.update_filter_display()
            self.editor.status_var.set("Filter(s) removed")
            
        def close_dialog():
            """Close the dialog"""
            dialog.destroy()
            
        ttk.Button(button_frame, text="Remove Selected", command=remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=close_dialog).pack(side=tk.LEFT, padx=5)
    
    def update_filter_display(self):
        """Update the filter display label"""
        if not self.editor.active_filters:
            self.editor.filter_display.config(text="No filters active", foreground="gray")
        else:
            filter_count = len(self.editor.active_filters)
            if self.editor.filtered_df is not None:
                row_count = len(self.editor.filtered_df)
                total_count = len(self.editor.df)
                self.editor.filter_display.config(
                    text=f"{filter_count} filter(s) active - Showing {row_count} of {total_count} rows", 
                    foreground="blue"
                )
            else:
                self.editor.filter_display.config(
                    text=f"{filter_count} filter(s) active", 
                    foreground="blue"
                )