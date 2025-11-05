"""
Schedule Properties Module
Handles Schedule Properties dialog and tab creation functions for the XLS Editor
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd


class ScheduleProperties:
    def __init__(self, editor_instance):
        self.editor = editor_instance
    
    def open_schedule_properties(self):
        """Open Schedule Properties dialog with tabs like Revit"""
        if self.editor.df is None:
            messagebox.showwarning(self.editor.tr("Warning"), self.editor.tr("No file is currently loaded."))
            return
            
        # Create main dialog
        dialog = tk.Toplevel(self.editor.root)
        dialog.title("Schedule Properties")
        dialog.geometry("800x600")
        dialog.minsize(600, 400)  # Set minimum size
        dialog.resizable(True, True)  # Allow resizing
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (800 // 2)
        y = (dialog.winfo_screenheight() // 2) - (600 // 2)
        dialog.geometry(f"800x600+{x}+{y}")
        
        # Create notebook for tabs
        notebook = ttk.Notebook(dialog)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Fields
        self.create_fields_tab(notebook)
        
        # Tab 2: Filter
        self.create_filter_tab(notebook)
        
        # Tab 3: Sorting/Grouping
        self.create_sorting_tab(notebook)
        
        # Tab 4: Formula
        self.create_formula_tab(notebook)
        
        # Tab 5: Appearance
        self.create_appearance_tab(notebook)
        
        # Bottom buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        def apply_changes():
            """Apply all changes and refresh the view"""
            # Apply field visibility changes from Fields tab
            if hasattr(self.editor, 'scheduled_fields_listbox'):
                # Get current scheduled fields order (these become the visible columns)
                new_visible_columns = []
                for i in range(self.editor.scheduled_fields_listbox.size()):
                    field_name = self.editor.scheduled_fields_listbox.get(i)
                    if field_name in self.editor.original_df.columns:
                        new_visible_columns.append(field_name)
                
                # Update visible columns and working dataframe
                if new_visible_columns != self.editor.visible_columns:
                    self.editor.visible_columns = new_visible_columns
                    # Update working dataframe to show only visible columns in correct order
                    if self.editor.visible_columns:
                        self.editor.df = self.editor.original_df[self.editor.visible_columns].copy()
                    else:
                        self.editor.df = self.editor.original_df.copy()
                    self.editor.modified = True
                    
            # Apply appearance settings
            if hasattr(self.editor, 'appearance_vars'):
                for key, var in self.editor.appearance_vars.items():
                    self.editor.appearance_settings[key] = var.get()
            
            # Apply filters and sorting
            self.editor.filter_ops.apply_filters()
            self.apply_sorting()
            self.editor.data_ops.populate_treeview()
            self.editor.file_ops.update_file_info()
            self.editor.filter_ops.update_filter_display()
            self.editor.status_var.set("Schedule properties applied")
            dialog.destroy()
            
        ttk.Button(button_frame, text="OK", command=apply_changes).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Apply", command=lambda: [apply_changes(), self.open_schedule_properties()]).pack(side=tk.RIGHT, padx=5)
    
    def create_fields_tab(self, notebook):
        """Create Fields tab like Revit"""
        fields_frame = ttk.Frame(notebook)
        notebook.add(fields_frame, text="Fields")
        
        # Main layout
        main_frame = ttk.Frame(fields_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left side - Available fields
        left_frame = ttk.LabelFrame(main_frame, text="Available fields:", padding="5")
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        # Available fields listbox
        self.editor.available_fields_listbox = tk.Listbox(left_frame, selectmode=tk.MULTIPLE)
        self.editor.available_fields_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Populate available fields (from all sheets if available)
        if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
            # Multi-sheet mode: show fields from all sheets
            for sheet_name, df in self.editor.sheet_ops.available_sheets.items():
                # Add sheet header
                self.editor.available_fields_listbox.insert(tk.END, f"--- {sheet_name} ---")
                for col in df.columns:
                    self.editor.available_fields_listbox.insert(tk.END, f"{sheet_name}.{col}")
        elif self.editor.original_df is not None:
            # Single sheet mode: show current sheet fields
            for col in self.editor.original_df.columns:
                self.editor.available_fields_listbox.insert(tk.END, col)
        
        # Middle - Buttons
        middle_frame = ttk.Frame(main_frame)
        middle_frame.grid(row=0, column=1, padx=5)
        
        ttk.Button(middle_frame, text="Add →", command=self.add_field_to_schedule).pack(pady=5)
        ttk.Button(middle_frame, text="← Remove", command=self.remove_field_from_schedule).pack(pady=5)
        ttk.Button(middle_frame, text="↑", command=self.move_field_up).pack(pady=2)
        ttk.Button(middle_frame, text="↓", command=self.move_field_down).pack(pady=2)
        
        # Right side - Scheduled fields
        right_frame = ttk.LabelFrame(main_frame, text="Scheduled fields (in order):", padding="5")
        right_frame.grid(row=0, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        # Scheduled fields listbox
        self.editor.scheduled_fields_listbox = tk.Listbox(right_frame)
        self.editor.scheduled_fields_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Populate scheduled fields (currently visible columns)
        if self.editor.visible_columns:
            for col in self.editor.visible_columns:
                self.editor.scheduled_fields_listbox.insert(tk.END, col)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(0, weight=1)
    
    def create_filter_tab(self, notebook):
        """Create Filter tab like Revit"""
        filter_frame = ttk.Frame(notebook)
        notebook.add(filter_frame, text="Filter")
        
        # Main frame with padding
        main_frame = ttk.Frame(filter_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Instructions
        ttk.Label(main_frame, text="Filter by:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Filter list frame with scrollbar
        filter_list_frame = ttk.LabelFrame(main_frame, text="Active Filters", padding="5")
        filter_list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create frame for filter list with scrollbar
        list_container = ttk.Frame(filter_list_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        # Treeview to show current filters
        self.editor.filter_tree = ttk.Treeview(list_container, columns=('Field', 'Type', 'Value'), show='headings', height=6)
        self.editor.filter_tree.heading('Field', text='Field')
        self.editor.filter_tree.heading('Type', text='Filter Type')
        self.editor.filter_tree.heading('Value', text='Value')
        
        self.editor.filter_tree.column('Field', width=150)
        self.editor.filter_tree.column('Type', width=120)
        self.editor.filter_tree.column('Value', width=150)
        
        self.editor.filter_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar for filter tree
        filter_scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.editor.filter_tree.yview)
        filter_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.editor.filter_tree.configure(yscrollcommand=filter_scrollbar.set)
        
        # Populate existing filters
        self.refresh_filter_tree()
        
        # Add new filter section
        add_filter_frame = ttk.LabelFrame(main_frame, text="Add New Filter", padding="5")
        add_filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Filter row
        row_frame = ttk.Frame(add_filter_frame)
        row_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(row_frame, text="Field:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.editor.new_filter_field = tk.StringVar()
        field_combo = ttk.Combobox(row_frame, textvariable=self.editor.new_filter_field, width=15)
        if self.editor.original_df is not None:    
            field_combo['values'] = list(self.editor.original_df.columns)
        field_combo.grid(row=0, column=1, padx=5)
        
        ttk.Label(row_frame, text="Type:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.editor.new_filter_type = tk.StringVar()
        filter_combo = ttk.Combobox(row_frame, textvariable=self.editor.new_filter_type, width=15)
        filter_combo['values'] = ["equals", "not equals", "contains", "not contains", 
                                 "starts with", "ends with", "greater than", "less than",
                                 "greater or equal", "less or equal", "is empty", "is not empty"]
        filter_combo.set("equals")
        filter_combo.grid(row=0, column=3, padx=5)
        
        ttk.Label(row_frame, text="Value:").grid(row=0, column=4, sticky=tk.W, padx=(10, 5))
        self.editor.new_filter_value = tk.StringVar()
        value_entry = ttk.Entry(row_frame, textvariable=self.editor.new_filter_value, width=20)
        value_entry.grid(row=0, column=5, padx=5)
        
        # Case sensitive checkbox
        self.editor.new_filter_case = tk.BooleanVar()
        ttk.Checkbutton(row_frame, text="Case sensitive", 
                       variable=self.editor.new_filter_case).grid(row=0, column=6, padx=10)
        
        # Buttons frame
        button_frame = ttk.Frame(add_filter_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="Add Filter", 
                  command=self.add_filter_from_schedule_properties).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Selected", 
                  command=self.remove_selected_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All Filters", 
                  command=self.clear_all_filters_and_refresh).pack(side=tk.LEFT, padx=5)
                  
        # Update filter type hint
        def update_value_state(*args):
            if self.editor.new_filter_type.get() in ["is empty", "is not empty"]:
                value_entry.config(state="disabled")
                self.editor.new_filter_value.set("")
            else:
                value_entry.config(state="normal")
                
        self.editor.new_filter_type.trace('w', update_value_state)
    
    def refresh_filter_tree(self):
        """Refresh the filter tree view with current active filters"""
        # Clear existing items
        for item in self.editor.filter_tree.get_children():
            self.editor.filter_tree.delete(item)
            
        # Add current filters
        for filter_id, filter_info in self.editor.active_filters.items():
            value_display = filter_info['value'] if filter_info['value'] else '(empty)'
            self.editor.filter_tree.insert('', 'end', iid=filter_id, values=(
                filter_info['column'],
                filter_info['type'],
                value_display
            ))
    
    def add_filter_from_schedule_properties(self):
        """Add filter from Schedule Properties Filter tab"""
        if not self.editor.new_filter_field.get():
            messagebox.showwarning("Warning", "Please select a field.")
            return
            
        filter_type = self.editor.new_filter_type.get()
        filter_value = self.editor.new_filter_value.get().strip()
        
        # Check if value is required
        if filter_type not in ["is empty", "is not empty"] and not filter_value:
            messagebox.showwarning("Warning", "Please enter a filter value.")
            return
            
        # Store the filter
        filter_id = f"{self.editor.new_filter_field.get()}_{len(self.editor.active_filters)}"
        self.editor.active_filters[filter_id] = {
            'column': self.editor.new_filter_field.get(),
            'type': filter_type,
            'value': filter_value,
            'case_sensitive': self.editor.new_filter_case.get()
        }
        
        # Refresh the filter tree
        self.refresh_filter_tree()
        
        # Clear the input fields
        self.editor.new_filter_field.set("")
        self.editor.new_filter_type.set("equals")
        self.editor.new_filter_value.set("")
        self.editor.new_filter_case.set(False)
        
        # Update filter display
        self.editor.filter_ops.update_filter_display()
    
    def remove_selected_filter(self):
        """Remove selected filter from tree"""
        selection = self.editor.filter_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a filter to remove.")
            return
            
        for item in selection:
            if item in self.editor.active_filters:
                del self.editor.active_filters[item]
            self.editor.filter_tree.delete(item)
            
        self.editor.filter_ops.update_filter_display()
    
    def clear_all_filters_and_refresh(self):
        """Clear all filters and refresh the tree"""
        self.editor.filter_ops.clear_all_filters()
        self.refresh_filter_tree()
    
    def create_sorting_tab(self, notebook):
        """Create Sorting/Grouping tab like Revit"""
        sorting_frame = ttk.Frame(notebook)
        notebook.add(sorting_frame, text="Sorting/Grouping")
        
        main_frame = ttk.Frame(sorting_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Instructions
        instructions = ttk.Label(main_frame, 
            text="Sort by: Primary sort (sorts main parameter first)\n" +
                 "Then by: Secondary sort (sorts within groups created by primary sort)",
            font=("Arial", 9), foreground="blue")
        instructions.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 10))
        
        # Create sort rows
        self.editor.sort_rows = []
        for i in range(4):  # Allow up to 4 sort levels like Revit
            self.create_sort_row(main_frame, i + 1)
            
        # Options section
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="5")
        options_frame.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=10)
        
        # Grand totals
        self.editor.grand_totals_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Grand totals", 
                       variable=self.editor.grand_totals_var).pack(anchor=tk.W)
        
        # Itemize every instance
        self.editor.itemize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Itemize every instance", 
                       variable=self.editor.itemize_var).pack(anchor=tk.W)
                       
        # Test sorting button
        test_frame = ttk.Frame(main_frame)
        test_frame.grid(row=7, column=0, columnspan=4, pady=10)
        
        ttk.Button(test_frame, text="Preview Sort", 
                  command=self.preview_sorting).pack(side=tk.LEFT, padx=5)
        ttk.Button(test_frame, text="Clear Sort", 
                  command=self.clear_sorting).pack(side=tk.LEFT, padx=5)
                  
        # Current sort status
        self.editor.sort_status_label = ttk.Label(main_frame, text="No sorting applied", 
                                          font=("Arial", 8), foreground="gray")
        self.editor.sort_status_label.grid(row=8, column=0, columnspan=4, sticky=tk.W, pady=5)
    
    def preview_sorting(self):
        """Preview the sorting without closing the dialog"""
        self.apply_sorting()
        self.editor.data_ops.populate_treeview()
        self.update_sort_status()
        messagebox.showinfo("Preview", "Sorting preview applied! Check the main data view.")
    
    def clear_sorting(self):
        """Clear all sorting settings"""
        if hasattr(self.editor, 'sort_rows'):
            for sort_info in self.editor.sort_rows:
                sort_info['column_var'].set('(none)')
                sort_info['direction_var'].set('Ascending')
                sort_info['header_var'].set(False)
                sort_info['footer_var'].set(False)
                sort_info['blank_var'].set(False)
        
        # Reset dataframes to original order
        if self.editor.original_df is not None:
            if self.editor.visible_columns:
                self.editor.df = self.editor.original_df[self.editor.visible_columns].copy()
            else:
                self.editor.df = self.editor.original_df.copy()
                
        self.update_sort_status()
        messagebox.showinfo("Clear Sort", "All sorting cleared!")
    
    def create_sort_row(self, parent, row_num):
        """Create a single sort row"""
        row_frame = ttk.Frame(parent)
        row_frame.grid(row=row_num, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=2)
        
        if row_num == 1:
            label_text = "Sort by:"
        else:
            label_text = "Then by:"
            
        ttk.Label(row_frame, text=label_text, width=10).grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        # Column selection
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(row_frame, textvariable=column_var, width=20, state="readonly")
        
        # Get available columns from original dataframe
        available_columns = ['(none)']
        if self.editor.original_df is not None:
            available_columns.extend(list(self.editor.original_df.columns))
        column_combo['values'] = available_columns
        column_combo.set('(none)')
        column_combo.grid(row=0, column=1, padx=5)
        
        # Sort direction frame
        direction_frame = ttk.Frame(row_frame)
        direction_frame.grid(row=0, column=2, padx=10)
        
        direction_var = tk.StringVar(value="Ascending")
        ttk.Radiobutton(direction_frame, text="Ascending", variable=direction_var, 
                       value="Ascending").pack(side=tk.LEFT)
        ttk.Radiobutton(direction_frame, text="Descending", variable=direction_var, 
                       value="Descending").pack(side=tk.LEFT, padx=(10, 0))
        
        # Options frame
        options_frame = ttk.Frame(row_frame)
        options_frame.grid(row=0, column=3, padx=10)
        
        # Checkboxes for headers and footers
        header_var = tk.BooleanVar()
        footer_var = tk.BooleanVar()
        blank_var = tk.BooleanVar()
        
        ttk.Checkbutton(options_frame, text="Header", variable=header_var).pack(side=tk.LEFT)
        ttk.Checkbutton(options_frame, text="Footer", variable=footer_var).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(options_frame, text="Blank line", variable=blank_var).pack(side=tk.LEFT, padx=5)
        
        # Store the variables for later use
        sort_info = {
            'column_var': column_var,
            'direction_var': direction_var,
            'header_var': header_var,
            'footer_var': footer_var,
            'blank_var': blank_var
        }
        self.editor.sort_rows.append(sort_info)
        
        # Add trace to update subsequent dropdowns when selection changes
        def on_column_change(*args):
            self.update_sort_status()
            
        column_var.trace('w', on_column_change)
    
    def update_sort_status(self):
        """Update sort status and provide feedback"""
        if hasattr(self.editor, 'sort_rows'):
            sort_count = 0
            for sort_info in self.editor.sort_rows:
                if sort_info['column_var'].get() != '(none)':
                    sort_count += 1
            
            if sort_count > 0:
                self.editor.status_var.set(f"Sorting by {sort_count} column(s) - click Apply to see changes")
            else:
                self.editor.status_var.set("No sorting applied")
    
    def create_formula_tab(self, notebook):
        """Create Formula tab for creating calculated fields"""
        formula_frame = ttk.Frame(notebook)
        notebook.add(formula_frame, text="Formula")
        
        main_frame = ttk.Frame(formula_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Instructions
        instructions = ttk.Label(main_frame, 
            text="Create calculated fields by combining existing fields with formulas.\n" +
                 "Use field names in brackets like [Field Name] and operators +, -, *, /, etc.\n" +
                 "COUNT([Field Name]) - counts how many times the current row's value appears in that field.",
            font=("Arial", 9), foreground="blue", wraplength=600)
        instructions.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Existing formula fields
        formula_fields_frame = ttk.LabelFrame(main_frame, text="Existing Formula Fields", padding="5")
        formula_fields_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        formula_fields_frame.columnconfigure(0, weight=1)
        formula_fields_frame.rowconfigure(0, weight=1)
        
        # Formula fields treeview
        columns = ('Field Name', 'Formula', 'Type')
        self.editor.formula_tree = ttk.Treeview(formula_fields_frame, columns=columns, show='headings', height=6)
        for col in columns:
            self.editor.formula_tree.heading(col, text=col)
            self.editor.formula_tree.column(col, width=150)
        
        self.editor.formula_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        formula_scroll = ttk.Scrollbar(formula_fields_frame, orient=tk.VERTICAL, command=self.editor.formula_tree.yview)
        formula_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.editor.formula_tree.configure(yscrollcommand=formula_scroll.set)
        
        # Populate existing formula fields
        self.editor.formula_ops.refresh_formula_tree()
        
        # New formula section
        new_formula_frame = ttk.LabelFrame(main_frame, text="Create New Formula Field", padding="5")
        new_formula_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        new_formula_frame.columnconfigure(1, weight=1)
        
        # Field name
        ttk.Label(new_formula_frame, text="Field Name:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.editor.new_formula_name = tk.StringVar()
        name_entry = ttk.Entry(new_formula_frame, textvariable=self.editor.new_formula_name, width=20)
        name_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        # Formula expression
        ttk.Label(new_formula_frame, text="Formula:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        self.editor.new_formula_expression = tk.StringVar()
        formula_entry = ttk.Entry(new_formula_frame, textvariable=self.editor.new_formula_expression, width=50)
        formula_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        # Pi button
        def insert_pi():
            """Insert π (pi) value into formula"""
            current_formula = self.editor.new_formula_expression.get()
            self.editor.new_formula_expression.set(current_formula + "3.141592")
        
        pi_button = ttk.Button(new_formula_frame, text="π", command=insert_pi, width=3)
        pi_button.grid(row=1, column=2, padx=(0, 5), pady=2)
        
        # Formula type
        ttk.Label(new_formula_frame, text="Type:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5))
        self.editor.new_formula_type = tk.StringVar()
        type_combo = ttk.Combobox(new_formula_frame, textvariable=self.editor.new_formula_type, width=15, state="readonly")
        type_combo['values'] = ["Number"]
        type_combo.set("Number")
        type_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Available fields
        fields_frame = ttk.LabelFrame(main_frame, text="Available Fields (double-click to insert)", padding="5")
        fields_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        self.editor.available_fields_formula = tk.Listbox(fields_frame, height=8)
        self.editor.available_fields_formula.pack(fill=tk.BOTH, expand=True)
        
        # Populate available fields - include fields from all sheets if available
        if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
            # Multi-sheet mode: show fields from all sheets
            for sheet_name, df in self.editor.sheet_ops.available_sheets.items():
                # Add sheet header
                self.editor.available_fields_formula.insert(tk.END, f"=== {sheet_name} ===")
                for col in df.columns:
                    # Format: SheetName.FieldName for cross-sheet reference
                    if sheet_name == self.editor.sheet_ops.current_sheet:
                        # Current sheet - show both formats
                        self.editor.available_fields_formula.insert(tk.END, f"[{col}]")
                    else:
                        # Other sheet - show cross-sheet format
                        self.editor.available_fields_formula.insert(tk.END, f"{sheet_name}.{col}")
        elif self.editor.original_df is not None:
            # Single sheet mode: show current sheet fields only
            for col in self.editor.original_df.columns:
                self.editor.available_fields_formula.insert(tk.END, col)
        
        # Bind double-click to insert field
        self.editor.available_fields_formula.bind('<Double-Button-1>', self.editor.formula_ops.insert_field_in_formula)
        
        # Formula templates
        templates_frame = ttk.LabelFrame(main_frame, text="Formula Templates", padding="5")
        templates_frame.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        self.editor.formula_templates_listbox = tk.Listbox(templates_frame, height=8)
        self.editor.formula_templates_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Populate templates
        self.editor.formula_ops.refresh_formula_templates()
        
        # Bind double-click to load template
        self.editor.formula_templates_listbox.bind('<Double-Button-1>', self.editor.formula_ops.load_formula_template)
        
        # Buttons
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(buttons_frame, text="Create Formula Field", 
                  command=self.editor.formula_ops.create_formula_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Update Selected", 
                  command=self.editor.formula_ops.update_formula_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Delete Selected", 
                  command=self.editor.formula_ops.delete_formula_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Save as Template", 
                  command=self.editor.formula_ops.save_formula_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Refresh All Formulas", 
                  command=self.editor.formula_ops.refresh_all_formulas).pack(side=tk.LEFT, padx=5)
        
        # Formula examples
        examples_frame = ttk.LabelFrame(main_frame, text="Formula Examples", padding="5")
        examples_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        examples_text = """Examples:
• [Length] * [Width]                    (Multiply two fields)
• [Price] * 1.1                        (Add 10% markup)
• COUNT([Type])                        (Count occurrences of this row's value)
• COUNT(Main Steel)                    (Count total "Main Steel" in all data)
• COUNT(Main Steel) * [Diameter]       (Multiply by total count)
• [Length] * COUNT([Type])             (Multiply by occurrence count)
• [Diameter] * 3.141592                (Use π for circle calculations)
• 2 * 3.141592 * [Radius]              (Calculate circumference: 2πr)
• IF([Status] = "Active", [Salary], 0)  (Conditional formula)
• MAX([Value1], [Value2], [Value3])     (Maximum of multiple values)
• ROUND([Price] * [Quantity], 2)        (Round to 2 decimal places)
Tip: Click π button to insert 3.141592"""
        
        ttk.Label(examples_frame, text=examples_text, font=("Courier", 8), 
                 foreground="darkgreen").pack(anchor=tk.W)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
    
    def create_appearance_tab(self, notebook):
        """Create Appearance tab like Revit"""
        appearance_frame = ttk.Frame(notebook)
        notebook.add(appearance_frame, text="Appearance")
        
        main_frame = ttk.Frame(appearance_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Graphics section
        graphics_frame = ttk.LabelFrame(main_frame, text="Graphics", padding="5")
        graphics_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Grid lines
        grid_var = tk.BooleanVar(value=self.editor.appearance_settings['grid_lines'])
        ttk.Checkbutton(graphics_frame, text="Grid lines", variable=grid_var).pack(anchor=tk.W)
        
        # Outline
        outline_var = tk.BooleanVar(value=self.editor.appearance_settings['outline'])
        ttk.Checkbutton(graphics_frame, text="Outline", variable=outline_var).pack(anchor=tk.W)
        
        # Text section
        text_frame = ttk.LabelFrame(main_frame, text="Text", padding="5")
        text_frame.pack(fill=tk.X)
        
        # Show title
        title_var = tk.BooleanVar(value=self.editor.appearance_settings['show_title'])
        ttk.Checkbutton(text_frame, text="Show Title", variable=title_var).pack(anchor=tk.W)
        
        # Show headers
        headers_var = tk.BooleanVar(value=self.editor.appearance_settings['show_headers'])
        ttk.Checkbutton(text_frame, text="Show Headers", variable=headers_var).pack(anchor=tk.W)
        
        # Store variables for later use
        self.editor.appearance_vars = {
            'grid_lines': grid_var,
            'outline': outline_var,
            'show_title': title_var,
            'show_headers': headers_var
        }
    
    # Field management methods
    def add_field_to_schedule(self):
        """Add selected field to schedule (make it visible)"""
        selection = self.editor.available_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to add.")
            return
            
        for index in selection:
            field_name = self.editor.available_fields_listbox.get(index)
            # Check if field is already in scheduled fields
            scheduled_items = [self.editor.scheduled_fields_listbox.get(i) for i in range(self.editor.scheduled_fields_listbox.size())]
            if field_name not in scheduled_items:
                self.editor.scheduled_fields_listbox.insert(tk.END, field_name)
    
    def remove_field_from_schedule(self):
        """Remove selected field from schedule (hide it, but keep in available fields)"""
        selection = self.editor.scheduled_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to remove.")
            return
            
        # Remove selected items (in reverse order to maintain indices)
        for index in reversed(selection):
            self.editor.scheduled_fields_listbox.delete(index)
    
    def move_field_up(self):
        """Move selected field up in schedule"""
        selection = self.editor.scheduled_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to move.")
            return
            
        if selection[0] > 0:
            index = selection[0]
            item = self.editor.scheduled_fields_listbox.get(index)
            self.editor.scheduled_fields_listbox.delete(index)
            self.editor.scheduled_fields_listbox.insert(index - 1, item)
            self.editor.scheduled_fields_listbox.selection_set(index - 1)
    
    def move_field_down(self):
        """Move selected field down in schedule"""
        selection = self.editor.scheduled_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to move.")
            return
            
        if selection[0] < self.editor.scheduled_fields_listbox.size() - 1:
            index = selection[0]
            item = self.editor.scheduled_fields_listbox.get(index)
            self.editor.scheduled_fields_listbox.delete(index)
            self.editor.scheduled_fields_listbox.insert(index + 1, item)
            self.editor.scheduled_fields_listbox.selection_set(index + 1)
    
    def apply_sorting(self):
        """Apply sorting based on sort settings"""
        if self.editor.df is None or not hasattr(self.editor, 'sort_rows'):
            return
            
        # Get active sort criteria
        sort_columns = []
        sort_ascending = []
        
        for sort_info in self.editor.sort_rows:
            column = sort_info['column_var'].get()
            if column and column != '(none)':
                # Check if column exists in visible columns or original dataframe
                if column in self.editor.visible_columns or (self.editor.visible_columns and column in self.editor.original_df.columns):
                    sort_columns.append(column)
                    sort_ascending.append(sort_info['direction_var'].get() == "Ascending")
        
        # Apply sorting to the appropriate dataframe
        if sort_columns:
            try:
                # Determine which dataframe to sort
                if self.editor.filtered_df is not None:
                    # Sort the filtered dataframe
                    sorted_df = self.editor.filtered_df.sort_values(by=sort_columns, ascending=sort_ascending)
                    self.editor.filtered_df = sorted_df
                else:
                    # Sort the main working dataframe
                    sorted_df = self.editor.df.sort_values(by=sort_columns, ascending=sort_ascending)
                    self.editor.df = sorted_df
                    
                # Also sort the original dataframe to maintain consistency
                if sort_columns:
                    available_sort_cols = [col for col in sort_columns if col in self.editor.original_df.columns]
                    if available_sort_cols:
                        available_ascending = [sort_ascending[i] for i, col in enumerate(sort_columns) if col in available_sort_cols]
                        self.editor.original_df = self.editor.original_df.sort_values(by=available_sort_cols, ascending=available_ascending)
                        
                        # Update the working dataframe with new order
                        if self.editor.visible_columns:
                            self.editor.df = self.editor.original_df[self.editor.visible_columns].copy()
                        else:
                            self.editor.df = self.editor.original_df.copy()
                            
                self.editor.modified = True
                
            except Exception as e:
                messagebox.showerror("Sorting Error", f"Failed to apply sorting:\n{str(e)}")
                
        # Store current sort settings for future use
        self.editor.current_sort_columns = sort_columns
        self.editor.current_sort_ascending = sort_ascending