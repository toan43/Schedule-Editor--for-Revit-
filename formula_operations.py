"""
Formula Operations Module
Handles all formula calculation and template functions for the XLS Editor
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import json
import os
import re


class FormulaOperations:
    def __init__(self, editor_instance):
        self.editor = editor_instance
        self.load_formula_templates_from_file()
    
    def save_formula_templates_to_file(self):
        """Save formula templates to a file"""
        if not self.editor.formula_templates:
            return
        
        try:
            templates_file = "formula_templates.json"
            with open(templates_file, 'w') as f:
                json.dump(self.editor.formula_templates, f, indent=2)
        except Exception as e:
            print(f"Error saving formula templates: {e}")
    
    def load_formula_templates_from_file(self):
        """Load formula templates from file"""
        try:
            templates_file = "formula_templates.json"
            if os.path.exists(templates_file):
                with open(templates_file, 'r') as f:
                    self.editor.formula_templates = json.load(f)
        except Exception as e:
            print(f"Error loading formula templates: {e}")
            self.editor.formula_templates = {}
    
    def refresh_formula_tree(self):
        """Refresh the formula fields tree"""
        # Clear existing items
        for item in self.editor.formula_tree.get_children():
            self.editor.formula_tree.delete(item)
            
        # Add current formula fields
        for field_name, formula_info in self.editor.formula_fields.items():
            self.editor.formula_tree.insert('', 'end', values=(
                field_name,
                formula_info['expression'],
                formula_info['type']
            ))
    
    def refresh_formula_templates(self):
        """Refresh the formula templates listbox"""
        self.editor.formula_templates_listbox.delete(0, tk.END)
        for template_name in self.editor.formula_templates.keys():
            self.editor.formula_templates_listbox.insert(tk.END, template_name)
    
    def insert_field_in_formula(self, event):
        """Insert selected field into formula expression"""
        selection = self.editor.available_fields_formula.curselection()
        if selection:
            field_name = self.editor.available_fields_formula.get(selection[0])
            
            # Skip if it's a sheet header (like "=== SheetName ===")
            if field_name.startswith("==="):
                return
            
            current_formula = self.editor.new_formula_expression.get()
            
            # Check if it's already in the correct format
            if field_name.startswith("[") and field_name.endswith("]"):
                # Already bracketed (current sheet field)
                new_formula = current_formula + field_name
            elif "." in field_name and not field_name.startswith("["):
                # Cross-sheet reference (SheetName.FieldName)
                new_formula = current_formula + field_name
            else:
                # Plain field name - add brackets
                new_formula = current_formula + f"[{field_name}]"
            
            self.editor.new_formula_expression.set(new_formula)
    
    def load_formula_template(self, event):
        """Load selected template into formula fields"""
        selection = self.editor.formula_templates_listbox.curselection()
        if selection:
            template_name = self.editor.formula_templates_listbox.get(selection[0])
            if template_name in self.editor.formula_templates:
                template = self.editor.formula_templates[template_name]
                self.editor.new_formula_name.set(template['name'])
                self.editor.new_formula_expression.set(template['expression'])
                self.editor.new_formula_type.set(template['type'])
    
    def create_formula_field(self):
        """Create a new formula field"""
        field_name = self.editor.new_formula_name.get().strip()
        expression = self.editor.new_formula_expression.get().strip()
        field_type = self.editor.new_formula_type.get()
        
        if not field_name:
            messagebox.showwarning("Warning", "Please enter a field name.")
            return
            
        if not expression:
            messagebox.showwarning("Warning", "Please enter a formula expression.")
            return
            
        # Check if field name already exists
        if field_name in self.editor.df.columns or field_name in self.editor.formula_fields:
            messagebox.showwarning("Warning", f"Field '{field_name}' already exists.")
            return
        
        # Validate formula
        if not self.validate_formula(expression):
            return
            
        # Store formula information
        self.editor.formula_fields[field_name] = {
            'expression': expression,
            'type': field_type
        }
        
        # Calculate and add the new field
        try:
            self.calculate_formula_field(field_name)
            self.refresh_formula_tree()
            self.editor.data_ops.populate_treeview()  # Refresh the main view
            self.editor.modified = True
            
            # Sync to available_sheets if multi-sheet mode is active
            self.editor.sync_current_sheet_data()
            
            # Clear input fields
            self.editor.new_formula_name.set("")
            self.editor.new_formula_expression.set("")
            self.editor.new_formula_type.set("Number")
            
            messagebox.showinfo("Success", f"Formula field '{field_name}' created successfully!")
            
        except Exception as e:
            del self.editor.formula_fields[field_name]  # Remove if calculation failed
            messagebox.showerror("Error", f"Failed to create formula field:\n{str(e)}")
    
    def update_formula_field(self):
        """Update selected formula field"""
        selection = self.editor.formula_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a formula field to update.")
            return
            
        # Get the selected field name
        selected_item = self.editor.formula_tree.item(selection[0])
        old_field_name = selected_item['values'][0]
        
        field_name = self.editor.new_formula_name.get().strip()
        expression = self.editor.new_formula_expression.get().strip()
        field_type = self.editor.new_formula_type.get()
        
        if not field_name or not expression:
            messagebox.showwarning("Warning", "Please enter field name and formula expression.")
            return
        
        # Validate formula
        if not self.validate_formula(expression):
            return
        
        try:
            # Remove old field if name changed
            if field_name != old_field_name:
                if old_field_name in self.editor.df.columns:
                    self.editor.df = self.editor.df.drop(columns=[old_field_name])
                if old_field_name in self.editor.original_df.columns:
                    self.editor.original_df = self.editor.original_df.drop(columns=[old_field_name])
                del self.editor.formula_fields[old_field_name]
            
            # Update formula information
            self.editor.formula_fields[field_name] = {
                'expression': expression,
                'type': field_type
            }
            
            # Recalculate the field
            self.calculate_formula_field(field_name)
            self.refresh_formula_tree()
            self.editor.data_ops.populate_treeview()
            self.editor.modified = True
            
            # Sync to available_sheets if multi-sheet mode is active
            self.editor.sync_current_sheet_data()
            
            messagebox.showinfo("Success", f"Formula field '{field_name}' updated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update formula field:\n{str(e)}")
    
    def delete_formula_field(self):
        """Delete selected formula field"""
        selection = self.editor.formula_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a formula field to delete.")
            return
            
        selected_item = self.editor.formula_tree.item(selection[0])
        field_name = selected_item['values'][0]
        
        if messagebox.askyesno("Confirm", f"Delete formula field '{field_name}'?"):
            try:
                # Remove from dataframes
                if field_name in self.editor.df.columns:
                    self.editor.df = self.editor.df.drop(columns=[field_name])
                if field_name in self.editor.original_df.columns:
                    self.editor.original_df = self.editor.original_df.drop(columns=[field_name])
                
                # Remove from formula fields
                del self.editor.formula_fields[field_name]
                
                # Remove from visible columns if present
                if field_name in self.editor.visible_columns:
                    self.editor.visible_columns.remove(field_name)
                
                # Sync to available_sheets if multi-sheet mode is active
                self.editor.sync_current_sheet_data()
                
                self.refresh_formula_tree()
                self.editor.data_ops.populate_treeview()
                self.editor.modified = True
                
                messagebox.showinfo("Success", f"Formula field '{field_name}' deleted successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete formula field:\n{str(e)}")
    
    def save_formula_template(self):
        """Save current formula as a template"""
        field_name = self.editor.new_formula_name.get().strip()
        expression = self.editor.new_formula_expression.get().strip()
        field_type = self.editor.new_formula_type.get()
        
        if not field_name or not expression:
            messagebox.showwarning("Warning", "Please enter field name and formula expression.")
            return
        
        template_name = simpledialog.askstring("Save Template", 
                                             "Enter template name:", 
                                             initialvalue=field_name)
        if template_name:
            self.editor.formula_templates[template_name] = {
                'name': field_name,
                'expression': expression,
                'type': field_type
            }
            self.refresh_formula_templates()
            messagebox.showinfo("Success", f"Template '{template_name}' saved successfully!")
    
    def refresh_all_formulas(self):
        """Refresh all formula fields with current data"""
        if not self.editor.formula_fields:
            messagebox.showinfo("Info", "No formula fields to refresh.")
            return
        
        try:
            # Recalculate all formula fields
            for field_name in self.editor.formula_fields.keys():
                self.calculate_formula_field(field_name)
            
            self.editor.data_ops.populate_treeview()
            self.editor.modified = True
            messagebox.showinfo("Success", "All formula fields refreshed successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh formula fields:\n{str(e)}")
    
    def validate_formula(self, expression):
        """Validate formula expression (supports cross-sheet references and COUNT function)"""
        if not expression:
            return False
        
        # Regex for COUNT variations
        count_field_pattern = r'COUNT\(\[([^\]]+)\]\)'      # COUNT([Field])
        count_value_pattern = r'COUNT\(([^)]+)\)'          # COUNT(Value) or COUNT(Sheet.Value)

        # Regex for fixed value references: Sheet.[Column(index)]
        fixed_value_pattern = r'(\w+)\.\[([\w\s-]+)\((\d+)\)\]'
        fixed_value_refs = re.findall(fixed_value_pattern, expression)

        # Temporarily remove fixed value references to avoid incorrect parsing
        temp_expression = expression
        fixed_value_placeholders = {}
        placeholder_count = 0

        def replace_fixed_val(match):
            nonlocal placeholder_count
            placeholder = f"__FIXED_VAL_{placeholder_count}__"
            fixed_value_placeholders[placeholder] = match.group(0) # Store the full match
            placeholder_count += 1
            return placeholder

        temp_expression = re.sub(fixed_value_pattern, replace_fixed_val, temp_expression)

        # Regex for standard field references on the modified expression
        field_pattern = r'\[([^\]]+)\]'
        referenced_fields = re.findall(field_pattern, temp_expression)

        # Temporarily remove COUNT(...) expressions to avoid false positives in cross-sheet pattern
        temp_for_cross_sheet = temp_expression
        count_placeholders = {}
        count_placeholder_id = 0
        
        def replace_count(match):
            nonlocal count_placeholder_id
            placeholder = f"__COUNT_{count_placeholder_id}__"
            count_placeholders[placeholder] = match.group(0)
            count_placeholder_id += 1
            return placeholder
        
        temp_for_cross_sheet = re.sub(r'COUNT\([^)]+\)', replace_count, temp_for_cross_sheet, flags=re.IGNORECASE)

        # Regex for cross-sheet row-by-row references: Sheet.Field (but not in COUNT or fixed value syntax)
        cross_sheet_pattern = r'(\w+)\.([\w\s-]+)(?!\[\w+\(\d+\)\])' # Exclude fixed value syntax
        cross_sheet_refs = re.findall(cross_sheet_pattern, temp_for_cross_sheet)
        
        # --- Validation Logic ---
        if self.editor.original_df is not None:
            available_fields = list(self.editor.original_df.columns) + list(self.editor.formula_fields.keys())
            
            # 1. Validate COUNT([Field])
            for field in re.findall(count_field_pattern, expression, re.IGNORECASE):
                if field not in self.editor.original_df.columns:
                    messagebox.showerror("Error", f"COUNT function: Field '{field}' not found in current sheet.")
                    return False

            # 2. Validate COUNT(Value) and COUNT(Sheet.Value)
            if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                for value_str in re.findall(count_value_pattern, expression, re.IGNORECASE):
                    # Check if it's a cross-sheet COUNT, e.g., "Sheet1.Main Steel"
                    if '.' in value_str:
                        parts = value_str.split('.', 1)
                        sheet_name, val = parts[0].strip(), parts[1].strip()
                        
                        # Skip if sheet_name looks like a number (e.g., from 3.14)
                        if sheet_name.isdigit():
                            continue
                        
                        if sheet_name not in self.editor.sheet_ops.available_sheets:
                            messagebox.showerror("Error", f"COUNT function: Sheet '{sheet_name}' not found.")
                            return False
            
            # 3. Validate [Field] references (from the temp_expression)
            for field in referenced_fields:
                if field not in available_fields:
                    messagebox.showerror("Error", f"Field '{field}' not found in current sheet.")
                    return False
            
            # 4. Validate Sheet.Field references (row-by-row)
            if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                for sheet_name, field_name in cross_sheet_refs:
                    if sheet_name.isdigit(): continue # Skip numbers
                    if sheet_name not in self.editor.sheet_ops.available_sheets:
                        messagebox.showerror("Error", f"Sheet '{sheet_name}' not found.")
                        return False
                    sheet_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    if field_name not in sheet_df.columns:
                        messagebox.showerror("Error", f"Field '{field_name}' not found in sheet '{sheet_name}'.")
                        return False
            elif cross_sheet_refs:
                messagebox.showerror("Error", "Cross-sheet references are not available. Please load multiple sheets first.")
                return False

            # 5. Validate Sheet.[Column(index)] references (fixed value)
            if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                for sheet_name, field_name, index_str in fixed_value_refs:
                    if sheet_name.isdigit(): continue
                    if sheet_name not in self.editor.sheet_ops.available_sheets:
                        messagebox.showerror("Error", f"Reference Error: Sheet '{sheet_name}' not found.")
                        return False
                    sheet_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    if field_name not in sheet_df.columns:
                        messagebox.showerror("Error", f"Reference Error: Field '{field_name}' not found in sheet '{sheet_name}'.")
                        return False
                    index = int(index_str)
                    if not (0 <= index < len(sheet_df)):
                        messagebox.showerror("Error", f"Reference Error: Index {index} is out of bounds for sheet '{sheet_name}'.")
                        return False
            elif fixed_value_refs:
                messagebox.showerror("Error", "Fixed value references (Sheet.[Col(index)]) are not available. Please load multiple sheets first.")
                return False

        # Basic syntax validation (unchanged)
        try:
            test_expression = expression
            # Replace [Field]
            for field in re.findall(field_pattern, expression):
                test_expression = test_expression.replace(f'[{field}]', '1')
            # Replace Sheet.Field
            for sheet_name, field_name in re.findall(cross_sheet_pattern, expression):
                test_expression = test_expression.replace(f'{sheet_name}.{field_name}', '1')
            # Replace Sheet.[Field(index)]
            for sheet_name, field_name, index in fixed_value_refs:
                test_expression = test_expression.replace(f'{sheet_name}.[{field_name}({index})]', '1')

            # Validate HAS_VALUE(Sheet, "Column", "Value")
            has_value_pattern = r'HAS_VALUE\(([^,]+),\s*["\']?([^,"]+)["\']?,\s*["\']?([^,"]+)["\']?\)'
            has_value_calls = re.findall(has_value_pattern, expression, re.IGNORECASE)

            if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                for sheet_name, column_name, value in has_value_calls:
                    sheet_name = sheet_name.strip()
                    column_name = column_name.strip()
                    if sheet_name.isdigit(): continue # Skip numbers
                    if sheet_name not in self.editor.sheet_ops.available_sheets:
                        messagebox.showerror("Error", f"HAS_VALUE function: Sheet '{sheet_name}' not found.")
                        return False
                    sheet_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    if column_name not in sheet_df.columns:
                        messagebox.showerror("Error", f"HAS_VALUE function: Column '{column_name}' not found in sheet '{sheet_name}'.")
                        return False
            elif has_value_calls:
                messagebox.showerror("Error", "HAS_VALUE function requires multiple sheets to be loaded.")
                return False

            # Validate LOOKUP(Sheet, "ColumnToGet", "FilterColumn", "FilterValue")
            lookup_pattern = r'LOOKUP\(([^,]+),\s*["\']?([^,"]+)["\']?,\s*["\']?([^,"]+)["\']?,\s*["\']?([^,"]+)["\']?\)'
            lookup_calls = re.findall(lookup_pattern, expression, re.IGNORECASE)

            if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                for sheet_name, col_to_get, filter_col, filter_val in lookup_calls:
                    sheet_name = sheet_name.strip()
                    col_to_get = col_to_get.strip()
                    filter_col = filter_col.strip()
                    if sheet_name.isdigit(): continue
                    if sheet_name not in self.editor.sheet_ops.available_sheets:
                        messagebox.showerror("Error", f"LOOKUP function: Sheet '{sheet_name}' not found.")
                        return False
                    sheet_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    if col_to_get not in sheet_df.columns:
                        messagebox.showerror("Error", f"LOOKUP function: Column '{col_to_get}' not found in sheet '{sheet_name}'.")
                        return False
                    if filter_col not in sheet_df.columns:
                        messagebox.showerror("Error", f"LOOKUP function: Filter column '{filter_col}' not found in sheet '{sheet_name}'.")
                        return False
            elif lookup_calls:
                messagebox.showerror("Error", "LOOKUP function requires multiple sheets to be loaded.")
                return False

            if '[' in test_expression or ']' in test_expression:
                # This check might be too simple now, but let's see
                pass
        except:
            pass
        
        return True
    
    def calculate_formula_field(self, field_name):
        """Calculate values for a formula field (supports cross-sheet references and all COUNT variations)"""
        if field_name not in self.editor.formula_fields:
            return
        
        formula_info = self.editor.formula_fields[field_name]
        expression = formula_info['expression']
        field_type = formula_info['type']
        
        # --- Pre-process and pre-calculate fixed values ---
        processed_expression = expression

        # --- New Logic: Determine implicit filter context from HAS_VALUE ---
        filter_context = None
        has_value_pattern_for_context = r'HAS_VALUE\(([^,]+),\s*["\']?([^,"]+)["\']?,\s*["\']?([^,"]+)["\']?\)'
        context_match = re.search(has_value_pattern_for_context, processed_expression, re.IGNORECASE)
        if context_match and hasattr(self.editor, 'sheet_ops'):
            sheet_name_ctx, col_name_ctx, val_ctx = [s.strip() for s in context_match.groups()]
            if sheet_name_ctx in self.editor.sheet_ops.available_sheets:
                filter_context = {
                    "sheet": sheet_name_ctx,
                    "column": col_name_ctx,
                    "value": val_ctx
                }
        # --- End New Logic ---

        # 1. Pre-calculate Sheet.[Column(index)] fixed values
        if hasattr(self.editor, 'sheet_ops'):
            fixed_value_pattern = r'(\w+)\.\[([\w\s-]+)\((\d+)\)\]'
            fixed_value_refs = re.findall(fixed_value_pattern, processed_expression)
            for sheet_name, field_name_ref, index_str in fixed_value_refs:
                if sheet_name.isdigit(): continue
                if sheet_name in self.editor.sheet_ops.available_sheets:
                    base_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    ref_df = base_df

                    # Apply implicit filter if context matches the sheet
                    if filter_context and filter_context["sheet"] == sheet_name:
                        try:
                            # Create a filtered view of the dataframe
                            filtered_df = base_df[base_df[filter_context["column"]].astype(str).str.strip() == filter_context["value"]]
                            if not filtered_df.empty:
                                ref_df = filtered_df
                        except Exception:
                            # If filtering fails, fall back to the original dataframe
                            ref_df = base_df
                    
                    index = int(index_str)
                    if field_name_ref in ref_df.columns and 0 <= index < len(ref_df):
                        value = ref_df.iloc[index][field_name_ref]
                        if pd.isna(value):
                            value = 0 if 'Number' in field_type else '""'
                        elif isinstance(value, str):
                            value = f'"{value}"'
                        
                        # Replace in the main expression string
                        full_ref = f'{sheet_name}.[{field_name_ref}({index_str})]'
                        processed_expression = processed_expression.replace(full_ref, str(value))

        # 2. Pre-calculate COUNT(Value) and COUNT(Sheet.Value) fixed values
        count_value_pattern = r'COUNT\(([^)]+)\)'
        count_value_matches = re.findall(count_value_pattern, processed_expression, re.IGNORECASE)
        data_for_counting = self.editor.filtered_df if self.editor.filtered_df is not None else self.editor.original_df

        for value_str in count_value_matches:
            # Skip if it's a COUNT([Field]) which is handled row-by-row
            if re.fullmatch(r'\[[^\]]+\]', value_str.strip()):
                continue

            value_str = value_str.strip()
            target_df = data_for_counting
            search_value = value_str
            
            if '.' in value_str and hasattr(self.editor, 'sheet_ops'):
                parts = value_str.split('.', 1)
                sheet_name, val = parts[0].strip(), parts[1].strip()
                if not sheet_name.isdigit() and sheet_name in self.editor.sheet_ops.available_sheets:
                    target_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    search_value = val

            total_count = sum((target_df[col].astype(str).str.strip() == search_value).sum() for col in target_df.columns)
            
            # Replace in the main expression string
            # Use a function for replacement to handle regex special characters in value_str
            def repl(m):
                return str(total_count)
            
            processed_expression = re.sub(
                r'COUNT\(\s*' + re.escape(value_str) + r'\s*\)',
                repl,
                processed_expression,
                count=1, # Replace only one instance at a time
                flags=re.IGNORECASE
            )

        # 3. Pre-calculate LOOKUP(Sheet, "ColumnToGet", "FilterColumn", "FilterValue") fixed values
        if hasattr(self.editor, 'sheet_ops'):
            lookup_pattern = r'LOOKUP\(([^,]+),\s*["\']?([^,"]+)["\']?,\s*["\']?([^,"]+)["\']?,\s*["\']?([^,"]+)["\']?\)'
            lookup_calls = re.findall(lookup_pattern, processed_expression, re.IGNORECASE)
            
            for sheet_name, col_to_get, filter_col, filter_val in lookup_calls:
                sheet_name = sheet_name.strip()
                col_to_get = col_to_get.strip()
                filter_col = filter_col.strip()
                filter_val = filter_val.strip()
                
                if sheet_name.isdigit(): continue
                if sheet_name in self.editor.sheet_ops.available_sheets:
                    target_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    
                    try:
                        # Filter the dataframe based on filter column and value
                        filtered_df = target_df[target_df[filter_col].astype(str).str.strip() == filter_val]
                        
                        # Find the first non-zero value in the specified column
                        result_value = 0
                        if not filtered_df.empty and col_to_get in filtered_df.columns:
                            for val in filtered_df[col_to_get]:
                                if pd.notna(val) and val != 0 and val != '0' and val != '':
                                    result_value = val
                                    break
                        
                        # Replace the LOOKUP call with the result value
                        original_call_regex = r'LOOKUP\(\s*' + re.escape(sheet_name) + r'\s*,\s*["\']?' + re.escape(col_to_get) + r'["\']?\s*,\s*["\']?' + re.escape(filter_col) + r'["\']?\s*,\s*["\']?' + re.escape(filter_val) + r'["\']?\s*\)'
                        processed_expression = re.sub(original_call_regex, str(result_value), processed_expression, flags=re.IGNORECASE)
                        
                    except Exception as e:
                        # If lookup fails, replace with 0
                        processed_expression = re.sub(original_call_regex, '0', processed_expression, flags=re.IGNORECASE)

        # 4. Pre-calculate HAS_VALUE(Sheet, "Column", "Value") - checks entire sheet once
        if hasattr(self.editor, 'sheet_ops') and 'HAS_VALUE' in processed_expression.upper():
            has_value_pattern = r'HAS_VALUE\(([^,]+),\s*["\']?([^,"]+)["\']?,\s*["\']?([^,"]+)["\']?\)'
            has_value_calls = re.findall(has_value_pattern, processed_expression, re.IGNORECASE)
            
            for sheet_name, column_name, value_to_check in has_value_calls:
                sheet_name = sheet_name.strip()
                column_name = column_name.strip()
                value_to_check = value_to_check.strip()
                
                result = "False"  # Default to False
                if not sheet_name.isdigit() and sheet_name in self.editor.sheet_ops.available_sheets:
                    ref_df = self.editor.sheet_ops.available_sheets[sheet_name]
                    if column_name in ref_df.columns:
                        # Check if ANY row in the sheet has the specified value
                        matching_rows = ref_df[ref_df[column_name].astype(str).str.strip() == value_to_check]
                        if len(matching_rows) > 0:
                            result = "True"
                
                # Replace the HAS_VALUE call with its boolean result
                original_call_regex = r'HAS_VALUE\(\s*' + re.escape(sheet_name) + r'\s*,\s*["\']?' + re.escape(column_name) + r'["\']?\s*,\s*["\']?' + re.escape(value_to_check) + r'["\']?\s*\)'
                processed_expression = re.sub(original_call_regex, result, processed_expression, flags=re.IGNORECASE)

        # --- Row-by-row evaluation using the processed expression ---
        
        # Regex for row-dependent COUNT([Field])
        count_field_pattern = r'COUNT\(\[([^\]]+)\]\)'
        count_field_matches = re.findall(count_field_pattern, processed_expression, re.IGNORECASE)
        
        # Cache for COUNT([Field]) results
        count_field_cache = {}
        if count_field_matches:
            for field in count_field_matches:
                if field in data_for_counting.columns:
                    value_counts = data_for_counting[field].value_counts().to_dict()
                    count_field_cache[field] = value_counts

        # --- Evaluate formula row by row ---
        # IMPORTANT: Only calculate for visible rows if filter is active
        target_indices = None
        if self.editor.filtered_df is not None:
            # Filter is active - only calculate for visible rows
            target_indices = set(self.editor.filtered_df.index)
        
        result_values = []
        for index, row in self.editor.original_df.iterrows():
            # Skip this row if filter is active and row is not visible
            if target_indices is not None and index not in target_indices:
                # Set empty/zero value for hidden rows
                result_values.append(None)
                continue
            
            try:
                evaluated_expression = processed_expression
                
                # A. Replace COUNT([Field]) with row-specific values
                for count_field in count_field_matches:
                    if count_field in count_field_cache and count_field in row:
                        current_value = row[count_field]
                        count_result = count_field_cache[count_field].get(current_value, 0)
                        evaluated_expression = re.sub(
                            r'COUNT\(\[' + re.escape(count_field) + r'\]\)',
                            str(count_result),
                            evaluated_expression,
                            flags=re.IGNORECASE
                        )
                
                # B. Replace standard field references ([Field] and Sheet.Field)
                # First, temporarily remove fixed value refs to not confuse the next regex
                temp_eval_expr = evaluated_expression
                fixed_value_pattern_full = r'(\w+)\.\[([\w\s-]+)\((\d+)\)\]'
                fixed_value_placeholders = {}
                placeholder_count = 0

                def replace_fixed_val(match):
                    nonlocal placeholder_count
                    placeholder = f"__FIXED_VAL_{placeholder_count}__"
                    fixed_value_placeholders[placeholder] = match.group(0)
                    placeholder_count += 1
                    return placeholder
                
                temp_eval_expr = re.sub(fixed_value_pattern_full, replace_fixed_val, temp_eval_expr)

                # Find bracket-style field references
                field_pattern = r'\[([^\]]+)\]'
                referenced_fields = re.findall(field_pattern, temp_eval_expr)
                
                # Find cross-sheet references (row-by-row)
                cross_sheet_pattern = r'(\w+)\.([\w\s-]+)(?!\[\w+\(\d+\)\])'
                cross_sheet_refs = re.findall(cross_sheet_pattern, temp_eval_expr)
                
                # Handle bracket-style references (current sheet)
                for ref_field in referenced_fields:
                    if ref_field in self.editor.original_df.columns:
                        value = row[ref_field]
                        if pd.isna(value):
                            value = 0 if 'Number' in field_type else '""'
                        elif isinstance(value, str):
                            value = f'"{value}"'
                        evaluated_expression = evaluated_expression.replace(f'[{ref_field}]', str(value))
                    elif ref_field in self.editor.formula_fields and ref_field in self.editor.original_df.columns:
                        value = self.editor.original_df.loc[index, ref_field]
                        if pd.isna(value):
                            value = 0 if 'Number' in field_type else '""'
                        evaluated_expression = evaluated_expression.replace(f'[{ref_field}]', str(value))
                
                # Handle cross-sheet references (row-by-row)
                if hasattr(self.editor, 'sheet_ops') and self.editor.sheet_ops.available_sheets:
                    for sheet_name, field_name_ref in cross_sheet_refs:
                        if sheet_name.isdigit(): continue
                        if sheet_name in self.editor.sheet_ops.available_sheets:
                            ref_df = self.editor.sheet_ops.available_sheets[sheet_name]
                            if field_name_ref in ref_df.columns and len(ref_df) > 0:
                                ref_index = min(index, len(ref_df) - 1)
                                value = ref_df.iloc[ref_index][field_name_ref]
                                if pd.isna(value):
                                    value = 0 if 'Number' in field_type else '""'
                                elif isinstance(value, str):
                                    value = f'"{value}"'
                                evaluated_expression = evaluated_expression.replace(f'{sheet_name}.{field_name_ref}', str(value))
                
                # C. HAS_VALUE has been pre-calculated in the preprocessing step (no action needed here)

                # D. Evaluate the final expression
                result = self.evaluate_expression(evaluated_expression, field_type)
                result_values.append(result)
                
            except Exception as e:
                default_value = 0 if field_type == "Number" else ""
                result_values.append(default_value)
        
        # Add the calculated column to dataframes
        self.editor.original_df[field_name] = result_values
        
        # Update working dataframe
        if self.editor.visible_columns and field_name not in self.editor.visible_columns:
            self.editor.visible_columns.append(field_name)
        
        if self.editor.visible_columns:
            # Ensure all visible columns exist in original_df
            available_visible = [col for col in self.editor.visible_columns if col in self.editor.original_df.columns]
            self.editor.df = self.editor.original_df[available_visible].copy()
        else:
            self.editor.df = self.editor.original_df.copy()
        
        # Reapply filters if any
        if self.editor.active_filters:
            self.editor.filter_ops.apply_filters()
    
    def evaluate_expression(self, expression, field_type):
        """Safely evaluate a formula expression"""
        try:
            # Handle common Excel-like functions
            expression = expression.replace('IF(', 'if_func(')
            expression = expression.replace('MAX(', 'max(')
            expression = expression.replace('MIN(', 'min(')
            expression = expression.replace('ROUND(', 'round(')
            expression = expression.replace('ABS(', 'abs(')
            expression = expression.replace('HAS_VALUE(', 'has_value_func(')
            
            # Define safe evaluation context
            safe_dict = {
                "__builtins__": {},
                "max": max,
                "min": min,
                "round": round,
                "abs": abs,
                "if_func": lambda condition, true_val, false_val: true_val if condition else false_val,
                "True": True,
                "False": False
            }
            
            # Evaluate the expression
            result = eval(expression, safe_dict)
            
            # Convert result based on field type
            if field_type == "Number":
                return float(result) if result != "" else 0
            elif field_type == "Text":
                return str(result)
            else:  # Auto
                # Try to determine type automatically
                if isinstance(result, (int, float)):
                    return result
                else:
                    return str(result)
                    
        except Exception as e:
            # Return default value on error
            return 0 if field_type == "Number" else str(expression)