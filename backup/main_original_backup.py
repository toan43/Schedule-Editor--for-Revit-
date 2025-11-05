import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
from typing import Optional
import re


class XLSEditor:
    def __init__(self, root):
        self.root = root
        self.current_language = "en"  # Default to English
        self.translations = self.load_translations()
        
        self.root.title(self.tr("XLS File Editor with Filtering"))
        self.root.geometry("1200x800")
        
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
        
        # Load formula templates from file
        self.load_formula_templates_from_file()
        
        self.create_menu()
        self.create_widgets()
        
    def load_translations(self):
        """Load translation dictionaries"""
        translations = {
            "en": {  # American English
                "XLS File Editor with Filtering": "XLS File Editor with Filtering",
                "File": "File",
                "Import XLS File": "Import XLS File",
                "Save": "Save",
                "Save As": "Save As",
                "Exit": "Exit",
                "Edit": "Edit",
                "Add Row": "Add Row",
                "Delete Row": "Delete Row",
                "Add Column": "Add Column",
                "Delete Column": "Delete Column",
                "Filter": "Filter",
                "Clear All Filters": "Clear All Filters",
                "Manage Filters": "Manage Filters",
                "Set Header Row": "Set Header Row",
                "Schedule": "Schedule",
                "Schedule Properties": "Schedule Properties",
                "Add Parameter": "Add Parameter",
                "Remove Parameter": "Remove Parameter",
                "Language": "Language",
                "English": "English",
                "Vietnamese": "Vietnamese",
                "File Information": "File Information",
                "Current File:": "Current File:",
                "No file loaded": "No file loaded",
                "Filters": "Filters",
                "Clear All": "Clear All",
                "Manage": "Manage",
                "Header Row": "Header Row",
                "No filters active": "No filters active",
                "Header: Row": "Header: Row",
                "Data Editor": "Data Editor",
                "Row": "Row",
                "Ready": "Ready",
                "Warning": "Warning",
                "Error": "Error",
                "Success": "Success",
                "Cancel": "Cancel",
                "Apply": "Apply",
                "OK": "OK",
                "Fields": "Fields",
                "Filter": "Filter",
                "Sorting/Grouping": "Sorting/Grouping",
                "Formula": "Formula",
                "Appearance": "Appearance",
                "Language changed to": "Language changed to",
                "Select XLS File": "Select XLS File",
                "Excel files": "Excel files",
                "All files": "All files",
                "File imported successfully": "File imported successfully",
                "Failed to import file": "Failed to import file",
                "Import failed": "Import failed"
            },
            "vi": {  # Vietnamese
                "XLS File Editor with Filtering": "Trình Chỉnh Sửa File XLS với Bộ Lọc",
                "File": "Tệp",
                "Import XLS File": "Nhập File XLS",
                "Save": "Lưu",
                "Save As": "Lưu Thành",
                "Exit": "Thoát",
                "Edit": "Chỉnh Sửa",
                "Add Row": "Thêm Dòng",
                "Delete Row": "Xóa Dòng",
                "Add Column": "Thêm Cột",
                "Delete Column": "Xóa Cột",
                "Filter": "Bộ Lọc",
                "Clear All Filters": "Xóa Tất Cả Bộ Lọc",
                "Manage Filters": "Quản Lý Bộ Lọc",
                "Set Header Row": "Đặt Dòng Tiêu Đề",
                "Schedule": "Lịch Trình",
                "Schedule Properties": "Thuộc Tính Lịch Trình",
                "Add Parameter": "Thêm Tham Số",
                "Remove Parameter": "Xóa Tham Số",
                "Language": "Ngôn Ngữ",
                "English": "Tiếng Anh",
                "Vietnamese": "Tiếng Việt",
                "File Information": "Thông Tin Tệp",
                "Current File:": "Tệp Hiện Tại:",
                "No file loaded": "Không có tệp nào được tải",
                "Filters": "Bộ Lọc",
                "Clear All": "Xóa Tất Cả",
                "Manage": "Quản Lý",
                "Header Row": "Dòng Tiêu Đề",
                "No filters active": "Không có bộ lọc nào hoạt động",
                "Header: Row": "Tiêu Đề: Dòng",
                "Data Editor": "Trình Chỉnh Sửa Dữ Liệu",
                "Row": "Dòng",
                "Ready": "Sẵn Sàng",
                "Warning": "Cảnh Báo",
                "Error": "Lỗi",
                "Success": "Thành Công",
                "Cancel": "Hủy",
                "Apply": "Áp Dụng",
                "OK": "Đồng Ý",
                "Fields": "Trường",
                "Filter": "Bộ Lọc",
                "Sorting/Grouping": "Sắp Xếp/Nhóm",
                "Formula": "Công Thức",
                "Appearance": "Giao Diện",
                "Language changed to": "Ngôn ngữ đã được thay đổi thành",
                "Select XLS File": "Chọn File XLS",
                "Save XLS File": "Lưu File XLS",
                "Excel files": "File Excel",
                "All files": "Tất cả file",
                "File imported successfully": "File đã được nhập thành công",
                "Failed to import file": "Lỗi khi nhập file",
                "Import failed": "Nhập file thất bại",
                "Warning": "Cảnh Báo",
                "No file is currently loaded": "Chưa có file nào được tải",
                "Please select a row to delete": "Vui lòng chọn một dòng để xóa",
                "Column name already exists": "Tên cột đã tồn tại",
                "Failed to save file": "Lỗi khi lưu file",
                "Please select a column": "Vui lòng chọn một cột",
                "Please enter a filter value": "Vui lòng nhập giá trị bộ lọc",
                "Info": "Thông Tin",
                "No filters are currently active": "Không có bộ lọc nào đang hoạt động",
                "Please select a filter to remove": "Vui lòng chọn bộ lọc để xóa",
                "Row number must be 1 or greater": "Số dòng phải lớn hơn hoặc bằng 1",
                "Please enter a valid row number": "Vui lòng nhập số dòng hợp lệ",
            }
        }
        return translations
        
    def tr(self, text):
        """Translate text based on current language"""
        return self.translations.get(self.current_language, {}).get(text, text)
        
    def change_language(self, language_code):
        """Change the application language"""
        self.current_language = language_code
        # Refresh the entire interface
        self.refresh_interface()
        messagebox.showinfo(self.tr("Success"), 
                           f"{self.tr('Language changed to')} {self.tr('English' if language_code == 'en' else 'Vietnamese')}")
    
    def refresh_interface(self):
        """Refresh the entire interface with new language"""
        # Update window title
        self.root.title(self.tr("XLS File Editor with Filtering"))
        
        # Recreate the menu and widgets
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
            self.populate_treeview()
            self.update_file_info()
            self.update_filter_display()
            self.update_header_display()
        
    def create_menu(self):
        """Create the application menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("File"), menu=file_menu)
        file_menu.add_command(label=self.tr("Import XLS File"), command=self.import_file)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("Save"), command=self.save_file)
        file_menu.add_command(label=self.tr("Save As"), command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("Exit"), command=self.on_closing)
        
        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("Edit"), menu=edit_menu)
        edit_menu.add_command(label=self.tr("Add Row"), command=self.add_row)
        edit_menu.add_command(label=self.tr("Delete Row"), command=self.delete_row)
        edit_menu.add_command(label=self.tr("Add Column"), command=self.add_column)
        edit_menu.add_command(label=self.tr("Delete Column"), command=self.delete_column)
        
        # Filter menu
        filter_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("Filter"), menu=filter_menu)
        filter_menu.add_command(label=self.tr("Clear All Filters"), command=self.clear_all_filters)
        filter_menu.add_command(label=self.tr("Manage Filters"), command=self.manage_filters)
        filter_menu.add_separator()
        filter_menu.add_command(label=self.tr("Set Header Row"), command=self.set_header_row)
        
        # Schedule menu (Revit-like features)
        schedule_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.tr("Schedule"), menu=schedule_menu)
        schedule_menu.add_command(label=self.tr("Schedule Properties"), command=self.open_schedule_properties)
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
        
        ttk.Button(button_frame, text=self.tr("Import XLS File"), command=self.import_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text=self.tr("Save"), command=self.save_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text=self.tr("Save As"), command=self.save_as_file).pack(side=tk.LEFT, padx=(0, 5))
        
        # Filter control frame
        filter_frame = ttk.LabelFrame(main_frame, text="Filters", padding="5")
        filter_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        filter_frame.columnconfigure(1, weight=1)
        
        ttk.Button(filter_frame, text="Clear All", command=self.clear_all_filters).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(filter_frame, text="Manage", command=self.manage_filters).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(filter_frame, text="Header Row", command=self.set_header_row).grid(row=0, column=2, padx=(0, 5))
        ttk.Button(filter_frame, text="Schedule Properties", command=self.open_schedule_properties).grid(row=0, column=3, padx=(0, 5))
        
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
        self.tree.bind('<Double-1>', self.on_cell_double_click)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def import_file(self):
        """Import an XLS file"""
        file_path = filedialog.askopenfilename(
            title=self.tr("Select XLS File"),
            filetypes=[
                (self.tr("Excel files"), "*.xlsx *.xls"),
                (self.tr("All files"), "*.*")
            ]
        )
        
        if file_path:
            try:
                # Read the Excel file with the specified header row
                if file_path.endswith('.xlsx'):
                    self.original_df = pd.read_excel(file_path, engine='openpyxl', header=self.header_row)
                else:
                    self.original_df = pd.read_excel(file_path, engine='xlrd', header=self.header_row)
                
                # Set working dataframe and visible columns
                self.df = self.original_df.copy()
                self.visible_columns = list(self.df.columns)  # Initially all columns are visible
                
                # Clear formula fields since data structure might have changed
                self.formula_fields = {}
                
                self.current_file = file_path
                self.modified = False
                self.filtered_df = None
                self.active_filters = {}
                self.update_file_info()
                self.update_filter_display()
                self.update_header_display()
                self.populate_treeview()
                self.status_var.set(f"{self.tr('File imported successfully')}: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror(self.tr("Error"), f"{self.tr('Failed to import file')}:\n{str(e)}")
                self.status_var.set(self.tr("Import failed"))
                
    def update_file_info(self):
        """Update the file information display"""
        if self.current_file:
            filename = os.path.basename(self.current_file)
            if self.modified:
                filename += " *"
            self.file_label.config(text=filename, foreground="black")
        else:
            self.file_label.config(text="No file loaded", foreground="gray")
            
    def populate_treeview(self):
        """Populate the treeview with DataFrame data"""
        # Use filtered data if filters are active, otherwise use working data
        display_df = self.filtered_df if self.filtered_df is not None else self.df
        
        if display_df is None:
            return
        
        # Only show visible columns
        if self.visible_columns:
            visible_cols = [col for col in self.visible_columns if col in display_df.columns]
            if visible_cols:
                display_df = display_df[visible_cols]
            
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Configure columns
        columns = list(display_df.columns)
        self.tree['columns'] = columns
        self.tree['show'] = 'tree headings'
        
        # Configure column widths and headings
        self.tree.column('#0', width=50, minwidth=50)
        self.tree.heading('#0', text='Row')
        
        for col in columns:
            self.tree.column(col, width=100, minwidth=80)
            self.tree.heading(col, text=str(col))
            
        # Insert data
        for index, row in display_df.iterrows():
            values = [str(val) if pd.notna(val) else '' for val in row]
            # Use original DataFrame index for editing purposes
            original_index = index if self.filtered_df is None else display_df.index[display_df.index == index][0]
            self.tree.insert('', 'end', text=str(original_index), values=values)
            
    def on_cell_double_click(self, event):
        """Handle double-click on a cell for editing"""
        if self.df is None:
            return
            
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        
        if column == '#0':  # Row number column
            return
            
        # Get column index
        col_index = int(column.replace('#', '')) - 1
        if col_index >= len(self.df.columns):
            return
            
        # Get row index
        row_index = int(self.tree.item(item, 'text'))
        
        # Get current value
        current_value = self.df.iloc[row_index, col_index]
        if pd.isna(current_value):
            current_value = ''
        else:
            current_value = str(current_value)
            
        # Create edit dialog
        self.edit_cell(row_index, col_index, current_value)
        
    def edit_cell(self, row_index, col_index, current_value):
        """Open dialog to edit cell value"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Cell")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (300 // 2)
        y = (dialog.winfo_screenheight() // 2) - (150 // 2)
        dialog.geometry(f"300x150+{x}+{y}")
        
        # Create widgets
        ttk.Label(dialog, text=f"Edit cell [{row_index}, {self.df.columns[col_index]}]:").pack(pady=10)
        
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
                self.df.iloc[row_index, col_index] = None
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
                self.df.iloc[row_index, col_index] = new_value
            
            self.modified = True
            self.update_file_info()
            self.populate_treeview()
            self.status_var.set("Cell updated")
            dialog.destroy()
            
        ttk.Button(button_frame, text="Save", command=save_edit).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Bind Enter key to save
        entry.bind('<Return>', lambda e: save_edit())
        
    def add_row(self):
        """Add a new row to the DataFrame"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        # Add empty row
        new_row = pd.Series([None] * len(self.df.columns), index=self.df.columns)
        self.df = pd.concat([self.df, new_row.to_frame().T], ignore_index=True)
        
        self.modified = True
        self.update_file_info()
        self.populate_treeview()
        self.status_var.set("Row added")
        
    def delete_row(self):
        """Delete selected row"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning(self.tr("Warning"), self.tr("Please select a row to delete."))
            return
            
        row_index = int(self.tree.item(selection[0], 'text'))
        
        if messagebox.askyesno("Confirm", f"Delete row {row_index}?"):
            self.df = self.df.drop(index=row_index).reset_index(drop=True)
            self.modified = True
            self.update_file_info()
            self.populate_treeview()
            self.status_var.set("Row deleted")
            
    def add_column(self):
        """Add a new column to the DataFrame"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        # Get column name
        column_name = tk.simpledialog.askstring("Add Column", "Enter column name:")
        if column_name and column_name not in self.df.columns:
            self.df[column_name] = None
            self.modified = True
            self.update_file_info()
            self.populate_treeview()
            self.status_var.set(f"Column '{column_name}' added")
        elif column_name in self.df.columns:
            messagebox.showwarning(self.tr("Warning"), self.tr("Column name already exists."))
            
    def delete_column(self):
        """Delete a column from the DataFrame"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        # Get column to delete
        columns = list(self.df.columns)
        if not columns:
            return
            
        # Create selection dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Delete Column")
        dialog.geometry("250x150")
        dialog.transient(self.root)
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
                    self.df = self.df.drop(columns=[column_var.get()])
                    self.modified = True
                    self.update_file_info()
                    self.populate_treeview()
                    self.status_var.set(f"Column '{column_var.get()}' deleted")
                    dialog.destroy()
                    
        ttk.Button(button_frame, text="Delete", command=delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
    def save_file(self):
        """Save the current file"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        if self.current_file is None:
            self.save_as_file()
            return
            
        try:
            self.df.to_excel(self.current_file, index=False, engine='openpyxl')
            self.modified = False
            self.update_file_info()
            self.status_var.set("File saved successfully")
        except Exception as e:
            messagebox.showerror(self.tr("Error"), f"{self.tr('Failed to save file')}:\n{str(e)}")
            
    def save_as_file(self):
        """Save the file with a new name"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        file_path = filedialog.asksaveasfilename(
            title=self.tr("Save XLS File"),
            defaultextension=".xlsx",
            filetypes=[
                (self.tr("Excel files"), "*.xlsx"),
                (self.tr("All files"), "*.*")
            ]
        )
        
        if file_path:
            try:
                self.df.to_excel(file_path, index=False, engine='openpyxl')
                self.current_file = file_path
                self.modified = False
                self.update_file_info()
                self.status_var.set(f"File saved as: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror(self.tr("Error"), f"{self.tr('Failed to save file')}:\n{str(e)}")
                
    def save_formula_templates_to_file(self):
        """Save formula templates to a file"""
        if not self.formula_templates:
            return
        
        try:
            import json
            templates_file = "formula_templates.json"
            with open(templates_file, 'w') as f:
                json.dump(self.formula_templates, f, indent=2)
        except Exception as e:
            print(f"Error saving formula templates: {e}")
    
    def load_formula_templates_from_file(self):
        """Load formula templates from file"""
        try:
            import json
            templates_file = "formula_templates.json"
            if os.path.exists(templates_file):
                with open(templates_file, 'r') as f:
                    self.formula_templates = json.load(f)
        except Exception as e:
            print(f"Error loading formula templates: {e}")
            self.formula_templates = {}

    def on_closing(self):
        """Handle application closing"""
        # Save formula templates before closing
        self.save_formula_templates_to_file()
        
        if self.modified:
            result = messagebox.askyesnocancel(
                "Unsaved Changes", 
                "You have unsaved changes. Do you want to save before closing?"
            )
            if result is True:  # Yes, save
                self.save_file()
                if not self.modified:  # Only close if save was successful
                    self.root.destroy()
            elif result is False:  # No, don't save
                self.root.destroy()
            # Cancel - do nothing
        else:
            self.root.destroy()
            
    # =============== FILTERING FUNCTIONALITY ===============
    
    def add_filter(self):
        """Add a new filter to the data"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        # Create filter dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Filter")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")
        
        # Column selection
        ttk.Label(dialog, text="Column:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(dialog, textvariable=column_var, values=list(self.df.columns), state="readonly")
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
                unique_values = self.df[column_var.get()].dropna().unique()
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
                messagebox.showwarning(self.tr("Warning"), self.tr("Please select a column."))
                return
                
            filter_type = filter_type_var.get()
            if filter_type not in ["is empty", "is not empty"] and not value_var.get():
                messagebox.showwarning(self.tr("Warning"), self.tr("Please enter a filter value."))
                return
                
            # Store the filter
            filter_id = f"{column_var.get()}_{len(self.active_filters)}"
            self.active_filters[filter_id] = {
                'column': column_var.get(),
                'type': filter_type,
                'value': value_var.get(),
                'case_sensitive': case_sensitive_var.get()
            }
            
            self.apply_filters()
            self.update_filter_display()
            self.status_var.set(f"Filter applied to column '{column_var.get()}'")
            dialog.destroy()
            
        ttk.Button(button_frame, text="Apply Filter", command=apply_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
    def apply_filters(self):
        """Apply all active filters to the DataFrame"""
        if self.df is None or not self.active_filters:
            self.filtered_df = None
            self.populate_treeview()
            return
            
        filtered_df = self.df.copy()
        
        for filter_info in self.active_filters.values():
            column = filter_info['column']
            filter_type = filter_info['type']
            value = filter_info['value']
            case_sensitive = filter_info['case_sensitive']
            
            if column not in filtered_df.columns:
                continue
                
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
                continue
                
            filtered_df = filtered_df[mask]
            
        self.filtered_df = filtered_df if len(filtered_df) < len(self.df) else None
        self.populate_treeview()
        
    def clear_all_filters(self):
        """Clear all active filters"""
        if not self.active_filters:
            messagebox.showinfo("Info", "No filters are currently active.")
            return
            
        if messagebox.askyesno("Confirm", "Clear all filters?"):
            self.active_filters = {}
            self.filtered_df = None
            self.apply_filters()
            self.update_filter_display()
            self.status_var.set("All filters cleared")
            
    def manage_filters(self):
        """Open filter management dialog"""
        if not self.active_filters:
            messagebox.showinfo("Info", "No filters are currently active.")
            return
            
        # Create filter management dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Manage Filters")
        dialog.geometry("500x300")
        dialog.transient(self.root)
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
        for filter_id, filter_info in self.active_filters.items():
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
                if item in self.active_filters:
                    del self.active_filters[item]
                filter_tree.delete(item)
                
            self.apply_filters()
            self.update_filter_display()
            self.status_var.set("Filter(s) removed")
            
        def close_dialog():
            """Close the dialog"""
            dialog.destroy()
            
        ttk.Button(button_frame, text="Remove Selected", command=remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=close_dialog).pack(side=tk.LEFT, padx=5)
        
    def update_filter_display(self):
        """Update the filter display label"""
        if not self.active_filters:
            self.filter_display.config(text="No filters active", foreground="gray")
        else:
            filter_count = len(self.active_filters)
            if self.filtered_df is not None:
                row_count = len(self.filtered_df)
                total_count = len(self.df)
                self.filter_display.config(
                    text=f"{filter_count} filter(s) active - Showing {row_count} of {total_count} rows", 
                    foreground="blue"
                )
            else:
                self.filter_display.config(
                    text=f"{filter_count} filter(s) active", 
                    foreground="blue"
                )
                
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
                
                self.populate_treeview()
                self.update_filter_display()
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
        
    # =============== SCHEDULE PROPERTIES (Revit-like) ===============
    
    def open_schedule_properties(self):
        """Open Schedule Properties dialog with tabs like Revit"""
        if self.df is None:
            messagebox.showwarning(self.tr("Warning"), self.tr("No file is currently loaded."))
            return
            
        # Create main dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Schedule Properties")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
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
            if hasattr(self, 'scheduled_fields_listbox'):
                # Get current scheduled fields order (these become the visible columns)
                new_visible_columns = []
                for i in range(self.scheduled_fields_listbox.size()):
                    field_name = self.scheduled_fields_listbox.get(i)
                    if field_name in self.original_df.columns:
                        new_visible_columns.append(field_name)
                
                # Update visible columns and working dataframe
                if new_visible_columns != self.visible_columns:
                    self.visible_columns = new_visible_columns
                    # Update working dataframe to show only visible columns in correct order
                    if self.visible_columns:
                        self.df = self.original_df[self.visible_columns].copy()
                    else:
                        self.df = self.original_df.copy()
                    self.modified = True
                    
            # Apply appearance settings
            if hasattr(self, 'appearance_vars'):
                for key, var in self.appearance_vars.items():
                    self.appearance_settings[key] = var.get()
            
            # Apply filters and sorting
            self.apply_filters()
            self.apply_sorting()
            self.populate_treeview()
            self.update_file_info()
            self.update_filter_display()
            self.status_var.set("Schedule properties applied")
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
        self.available_fields_listbox = tk.Listbox(left_frame, selectmode=tk.MULTIPLE)
        self.available_fields_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Populate available fields (all original columns)
        if self.original_df is not None:
            for col in self.original_df.columns:
                self.available_fields_listbox.insert(tk.END, col)
        
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
        self.scheduled_fields_listbox = tk.Listbox(right_frame)
        self.scheduled_fields_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Populate scheduled fields (currently visible columns)
        if self.visible_columns:
            for col in self.visible_columns:
                self.scheduled_fields_listbox.insert(tk.END, col)
        
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
        self.filter_tree = ttk.Treeview(list_container, columns=('Field', 'Type', 'Value'), show='headings', height=6)
        self.filter_tree.heading('Field', text='Field')
        self.filter_tree.heading('Type', text='Filter Type')
        self.filter_tree.heading('Value', text='Value')
        
        self.filter_tree.column('Field', width=150)
        self.filter_tree.column('Type', width=120)
        self.filter_tree.column('Value', width=150)
        
        self.filter_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar for filter tree
        filter_scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.filter_tree.yview)
        filter_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.filter_tree.configure(yscrollcommand=filter_scrollbar.set)
        
        # Populate existing filters
        self.refresh_filter_tree()
        
        # Add new filter section
        add_filter_frame = ttk.LabelFrame(main_frame, text="Add New Filter", padding="5")
        add_filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Filter row
        row_frame = ttk.Frame(add_filter_frame)
        row_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(row_frame, text="Field:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.new_filter_field = tk.StringVar()
        field_combo = ttk.Combobox(row_frame, textvariable=self.new_filter_field, width=15)
        if self.original_df is not None:    
            field_combo['values'] = list(self.original_df.columns)
        field_combo.grid(row=0, column=1, padx=5)
        
        ttk.Label(row_frame, text="Type:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.new_filter_type = tk.StringVar()
        filter_combo = ttk.Combobox(row_frame, textvariable=self.new_filter_type, width=15)
        filter_combo['values'] = ["equals", "not equals", "contains", "not contains", 
                                 "starts with", "ends with", "greater than", "less than",
                                 "greater or equal", "less or equal", "is empty", "is not empty"]
        filter_combo.set("equals")
        filter_combo.grid(row=0, column=3, padx=5)
        
        ttk.Label(row_frame, text="Value:").grid(row=0, column=4, sticky=tk.W, padx=(10, 5))
        self.new_filter_value = tk.StringVar()
        value_entry = ttk.Entry(row_frame, textvariable=self.new_filter_value, width=20)
        value_entry.grid(row=0, column=5, padx=5)
        
        # Case sensitive checkbox
        self.new_filter_case = tk.BooleanVar()
        ttk.Checkbutton(row_frame, text="Case sensitive", 
                       variable=self.new_filter_case).grid(row=0, column=6, padx=10)
        
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
            if self.new_filter_type.get() in ["is empty", "is not empty"]:
                value_entry.config(state="disabled")
                self.new_filter_value.set("")
            else:
                value_entry.config(state="normal")
                
        self.new_filter_type.trace('w', update_value_state)
        
    def refresh_filter_tree(self):
        """Refresh the filter tree view with current active filters"""
        # Clear existing items
        for item in self.filter_tree.get_children():
            self.filter_tree.delete(item)
            
        # Add current filters
        for filter_id, filter_info in self.active_filters.items():
            value_display = filter_info['value'] if filter_info['value'] else '(empty)'
            self.filter_tree.insert('', 'end', iid=filter_id, values=(
                filter_info['column'],
                filter_info['type'],
                value_display
            ))
            
    def add_filter_from_schedule_properties(self):
        """Add filter from Schedule Properties Filter tab"""
        if not self.new_filter_field.get():
            messagebox.showwarning("Warning", "Please select a field.")
            return
            
        filter_type = self.new_filter_type.get()
        filter_value = self.new_filter_value.get().strip()
        
        # Check if value is required
        if filter_type not in ["is empty", "is not empty"] and not filter_value:
            messagebox.showwarning("Warning", "Please enter a filter value.")
            return
            
        # Store the filter
        filter_id = f"{self.new_filter_field.get()}_{len(self.active_filters)}"
        self.active_filters[filter_id] = {
            'column': self.new_filter_field.get(),
            'type': filter_type,
            'value': filter_value,
            'case_sensitive': self.new_filter_case.get()
        }
        
        # Refresh the filter tree
        self.refresh_filter_tree()
        
        # Clear the input fields
        self.new_filter_field.set("")
        self.new_filter_type.set("equals")
        self.new_filter_value.set("")
        self.new_filter_case.set(False)
        
        # Update filter display
        self.update_filter_display()
        
    def remove_selected_filter(self):
        """Remove selected filter from tree"""
        selection = self.filter_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a filter to remove.")
            return
            
        for item in selection:
            if item in self.active_filters:
                del self.active_filters[item]
            self.filter_tree.delete(item)
            
        self.update_filter_display()
        
    def clear_all_filters_and_refresh(self):
        """Clear all filters and refresh the tree"""
        self.clear_all_filters()
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
        self.sort_rows = []
        for i in range(4):  # Allow up to 4 sort levels like Revit
            self.create_sort_row(main_frame, i + 1)
            
        # Options section
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="5")
        options_frame.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=10)
        
        # Grand totals
        self.grand_totals_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Grand totals", 
                       variable=self.grand_totals_var).pack(anchor=tk.W)
        
        # Itemize every instance
        self.itemize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Itemize every instance", 
                       variable=self.itemize_var).pack(anchor=tk.W)
                       
        # Test sorting button
        test_frame = ttk.Frame(main_frame)
        test_frame.grid(row=7, column=0, columnspan=4, pady=10)
        
        ttk.Button(test_frame, text="Apply Sort", 
                  command=self.preview_sorting).pack(side=tk.LEFT, padx=5)
        ttk.Button(test_frame, text="Clear Sort", 
                  command=self.clear_sorting).pack(side=tk.LEFT, padx=5)
                  
        # Current sort status
        self.sort_status_label = ttk.Label(main_frame, text="No sorting applied", 
                                          font=("Arial", 8), foreground="gray")
        self.sort_status_label.grid(row=8, column=0, columnspan=4, sticky=tk.W, pady=5)
        
    def preview_sorting(self):
        """Preview the sorting without closing the dialog"""
        self.apply_sorting()
        self.populate_treeview()
        self.update_sort_status()
        messagebox.showinfo("Preview", "Sorting preview applied! Check the main data view.")
        
    def clear_sorting(self):
        """Clear all sorting settings"""
        if hasattr(self, 'sort_rows'):
            for sort_info in self.sort_rows:
                sort_info['column_var'].set('(none)')
                sort_info['direction_var'].set('Ascending')
                sort_info['header_var'].set(False)
                sort_info['footer_var'].set(False)
                sort_info['blank_var'].set(False)
        
        # Reset dataframes to original order
        if self.original_df is not None:
            if self.visible_columns:
                self.df = self.original_df[self.visible_columns].copy()
            else:
                self.df = self.original_df.copy()
                
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
        if self.original_df is not None:
            available_columns.extend(list(self.original_df.columns))
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
        self.sort_rows.append(sort_info)
        
        # Add trace to update subsequent dropdowns when selection changes
        def on_column_change(*args):
            self.update_sort_status()
            
        column_var.trace('w', on_column_change)
        
    def update_sort_status(self):
        """Update sort status and provide feedback"""
        if hasattr(self, 'sort_rows'):
            sort_count = 0
            for sort_info in self.sort_rows:
                if sort_info['column_var'].get() != '(none)':
                    sort_count += 1
            
            if sort_count > 0:
                self.status_var.set(f"Sorting by {sort_count} column(s) - click Apply to see changes")
            else:
                self.status_var.set("No sorting applied")
        
    def create_formula_tab(self, notebook):
        """Create Formula tab for creating calculated fields"""
        formula_frame = ttk.Frame(notebook)
        notebook.add(formula_frame, text="Formula")
        
        main_frame = ttk.Frame(formula_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Instructions
        instructions = ttk.Label(main_frame, 
            text="Create calculated fields by combining existing fields with formulas.\n" +
                 "Use field names in brackets like [Field Name] and operators +, -, *, /, etc.",
            font=("Arial", 9), foreground="blue", wraplength=600)
        instructions.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Existing formula fields
        formula_fields_frame = ttk.LabelFrame(main_frame, text="Existing Formula Fields", padding="5")
        formula_fields_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        formula_fields_frame.columnconfigure(0, weight=1)
        formula_fields_frame.rowconfigure(0, weight=1)
        
        # Formula fields treeview
        columns = ('Field Name', 'Formula', 'Type')
        self.formula_tree = ttk.Treeview(formula_fields_frame, columns=columns, show='headings', height=6)
        for col in columns:
            self.formula_tree.heading(col, text=col)
            self.formula_tree.column(col, width=150)
        
        self.formula_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        formula_scroll = ttk.Scrollbar(formula_fields_frame, orient=tk.VERTICAL, command=self.formula_tree.yview)
        formula_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.formula_tree.configure(yscrollcommand=formula_scroll.set)
        
        # Populate existing formula fields
        self.refresh_formula_tree()
        
        # New formula section
        new_formula_frame = ttk.LabelFrame(main_frame, text="Create New Formula Field", padding="5")
        new_formula_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        new_formula_frame.columnconfigure(1, weight=1)
        
        # Field name
        ttk.Label(new_formula_frame, text="Field Name:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.new_formula_name = tk.StringVar()
        name_entry = ttk.Entry(new_formula_frame, textvariable=self.new_formula_name, width=20)
        name_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        # Formula expression
        ttk.Label(new_formula_frame, text="Formula:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        self.new_formula_expression = tk.StringVar()
        formula_entry = ttk.Entry(new_formula_frame, textvariable=self.new_formula_expression, width=50)
        formula_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        # Formula type
        ttk.Label(new_formula_frame, text="Type:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5))
        self.new_formula_type = tk.StringVar()
        type_combo = ttk.Combobox(new_formula_frame, textvariable=self.new_formula_type, width=15, state="readonly")
        type_combo['values'] = ["Number", "Text", "Auto"]
        type_combo.set("Auto")
        type_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Available fields
        fields_frame = ttk.LabelFrame(main_frame, text="Available Fields (double-click to insert)", padding="5")
        fields_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        self.available_fields_formula = tk.Listbox(fields_frame, height=8)
        self.available_fields_formula.pack(fill=tk.BOTH, expand=True)
        
        # Populate available fields
        if self.original_df is not None:
            for col in self.original_df.columns:
                self.available_fields_formula.insert(tk.END, col)
        
        # Bind double-click to insert field
        self.available_fields_formula.bind('<Double-Button-1>', self.insert_field_in_formula)
        
        # Formula templates
        templates_frame = ttk.LabelFrame(main_frame, text="Formula Templates", padding="5")
        templates_frame.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        self.formula_templates_listbox = tk.Listbox(templates_frame, height=8)
        self.formula_templates_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Populate templates
        self.refresh_formula_templates()
        
        # Bind double-click to load template
        self.formula_templates_listbox.bind('<Double-Button-1>', self.load_formula_template)
        
        # Buttons
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(buttons_frame, text="Create Formula Field", 
                  command=self.create_formula_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Update Selected", 
                  command=self.update_formula_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Delete Selected", 
                  command=self.delete_formula_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Save as Template", 
                  command=self.save_formula_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Refresh All Formulas", 
                  command=self.refresh_all_formulas).pack(side=tk.LEFT, padx=5)
        
        # Formula examples
        examples_frame = ttk.LabelFrame(main_frame, text="Formula Examples", padding="5")
        examples_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        examples_text = """Examples:
• [Length] * [Width]                    (Multiply two fields)
• [Price] * 1.1                        (Add 10% markup)
• [First Name] + " " + [Last Name]      (Concatenate text)
• IF([Status] = "Active", [Salary], 0)  (Conditional formula)
• MAX([Value1], [Value2], [Value3])     (Maximum of multiple values)
• ROUND([Price] * [Quantity], 2)        (Round to 2 decimal places)"""
        
        ttk.Label(examples_frame, text=examples_text, font=("Courier", 8), 
                 foreground="darkgreen").pack(anchor=tk.W)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
    def refresh_formula_tree(self):
        """Refresh the formula fields tree"""
        # Clear existing items
        for item in self.formula_tree.get_children():
            self.formula_tree.delete(item)
            
        # Add current formula fields
        for field_name, formula_info in self.formula_fields.items():
            self.formula_tree.insert('', 'end', values=(
                field_name,
                formula_info['expression'],
                formula_info['type']
            ))
    
    def refresh_formula_templates(self):
        """Refresh the formula templates listbox"""
        self.formula_templates_listbox.delete(0, tk.END)
        for template_name in self.formula_templates.keys():
            self.formula_templates_listbox.insert(tk.END, template_name)
    
    def insert_field_in_formula(self, event):
        """Insert selected field into formula expression"""
        selection = self.available_fields_formula.curselection()
        if selection:
            field_name = self.available_fields_formula.get(selection[0])
            current_formula = self.new_formula_expression.get()
            # Insert field name in brackets
            new_formula = current_formula + f"[{field_name}]"
            self.new_formula_expression.set(new_formula)
    
    def load_formula_template(self, event):
        """Load selected template into formula fields"""
        selection = self.formula_templates_listbox.curselection()
        if selection:
            template_name = self.formula_templates_listbox.get(selection[0])
            if template_name in self.formula_templates:
                template = self.formula_templates[template_name]
                self.new_formula_name.set(template['name'])
                self.new_formula_expression.set(template['expression'])
                self.new_formula_type.set(template['type'])
    
    def create_formula_field(self):
        """Create a new formula field"""
        field_name = self.new_formula_name.get().strip()
        expression = self.new_formula_expression.get().strip()
        field_type = self.new_formula_type.get()
        
        if not field_name:
            messagebox.showwarning("Warning", "Please enter a field name.")
            return
            
        if not expression:
            messagebox.showwarning("Warning", "Please enter a formula expression.")
            return
            
        # Check if field name already exists
        if field_name in self.df.columns or field_name in self.formula_fields:
            messagebox.showwarning("Warning", f"Field '{field_name}' already exists.")
            return
        
        # Validate formula
        if not self.validate_formula(expression):
            return
            
        # Store formula information
        self.formula_fields[field_name] = {
            'expression': expression,
            'type': field_type
        }
        
        # Calculate and add the new field
        try:
            self.calculate_formula_field(field_name)
            self.refresh_formula_tree()
            self.populate_treeview()  # Refresh the main view
            self.modified = True
            
            # Clear input fields
            self.new_formula_name.set("")
            self.new_formula_expression.set("")
            self.new_formula_type.set("Auto")
            
            messagebox.showinfo("Success", f"Formula field '{field_name}' created successfully!")
            
        except Exception as e:
            del self.formula_fields[field_name]  # Remove if calculation failed
            messagebox.showerror("Error", f"Failed to create formula field:\n{str(e)}")
    
    def update_formula_field(self):
        """Update selected formula field"""
        selection = self.formula_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a formula field to update.")
            return
            
        # Get the selected field name
        selected_item = self.formula_tree.item(selection[0])
        old_field_name = selected_item['values'][0]
        
        field_name = self.new_formula_name.get().strip()
        expression = self.new_formula_expression.get().strip()
        field_type = self.new_formula_type.get()
        
        if not field_name or not expression:
            messagebox.showwarning("Warning", "Please enter field name and formula expression.")
            return
        
        # Validate formula
        if not self.validate_formula(expression):
            return
        
        try:
            # Remove old field if name changed
            if field_name != old_field_name:
                if old_field_name in self.df.columns:
                    self.df = self.df.drop(columns=[old_field_name])
                if old_field_name in self.original_df.columns:
                    self.original_df = self.original_df.drop(columns=[old_field_name])
                del self.formula_fields[old_field_name]
            
            # Update formula information
            self.formula_fields[field_name] = {
                'expression': expression,
                'type': field_type
            }
            
            # Recalculate the field
            self.calculate_formula_field(field_name)
            self.refresh_formula_tree()
            self.populate_treeview()
            self.modified = True
            
            messagebox.showinfo("Success", f"Formula field '{field_name}' updated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update formula field:\n{str(e)}")
    
    def delete_formula_field(self):
        """Delete selected formula field"""
        selection = self.formula_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a formula field to delete.")
            return
            
        selected_item = self.formula_tree.item(selection[0])
        field_name = selected_item['values'][0]
        
        if messagebox.askyesno("Confirm", f"Delete formula field '{field_name}'?"):
            try:
                # Remove from dataframes
                if field_name in self.df.columns:
                    self.df = self.df.drop(columns=[field_name])
                if field_name in self.original_df.columns:
                    self.original_df = self.original_df.drop(columns=[field_name])
                
                # Remove from formula fields
                del self.formula_fields[field_name]
                
                # Remove from visible columns if present
                if field_name in self.visible_columns:
                    self.visible_columns.remove(field_name)
                
                self.refresh_formula_tree()
                self.populate_treeview()
                self.modified = True
                
                messagebox.showinfo("Success", f"Formula field '{field_name}' deleted successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete formula field:\n{str(e)}")
    
    def save_formula_template(self):
        """Save current formula as a template"""
        field_name = self.new_formula_name.get().strip()
        expression = self.new_formula_expression.get().strip()
        field_type = self.new_formula_type.get()
        
        if not field_name or not expression:
            messagebox.showwarning("Warning", "Please enter field name and formula expression.")
            return
        
        template_name = tk.simpledialog.askstring("Save Template", 
                                                 "Enter template name:", 
                                                 initialvalue=field_name)
        if template_name:
            self.formula_templates[template_name] = {
                'name': field_name,
                'expression': expression,
                'type': field_type
            }
            self.refresh_formula_templates()
            messagebox.showinfo("Success", f"Template '{template_name}' saved successfully!")
    
    def refresh_all_formulas(self):
        """Refresh all formula fields with current data"""
        if not self.formula_fields:
            messagebox.showinfo("Info", "No formula fields to refresh.")
            return
        
        try:
            # Recalculate all formula fields
            for field_name in self.formula_fields.keys():
                self.calculate_formula_field(field_name)
            
            self.populate_treeview()
            self.modified = True
            messagebox.showinfo("Success", "All formula fields refreshed successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh formula fields:\n{str(e)}")
    
    def validate_formula(self, expression):
        """Validate formula expression"""
        if not expression:
            return False
        
        # Check for field references in brackets
        import re
        field_pattern = r'\[([^\]]+)\]'
        referenced_fields = re.findall(field_pattern, expression)
        
        # Check if all referenced fields exist
        if self.original_df is not None:
            available_fields = list(self.original_df.columns) + list(self.formula_fields.keys())
            for field in referenced_fields:
                if field not in available_fields:
                    messagebox.showerror("Error", f"Field '{field}' not found in available fields.")
                    return False
        
        # Basic syntax validation
        try:
            # Replace field references with dummy values for syntax check
            test_expression = expression
            for field in referenced_fields:
                test_expression = test_expression.replace(f'[{field}]', '1')
            
            # Try to evaluate basic syntax (won't work for all functions but catches basic errors)
            # This is a simplified validation
            if any(char in test_expression for char in ['[', ']']):
                messagebox.showerror("Error", "Formula contains unmatched brackets.")
                return False
                
        except:
            pass  # Skip detailed syntax validation for now
        
        return True
    
    def calculate_formula_field(self, field_name):
        """Calculate values for a formula field"""
        if field_name not in self.formula_fields:
            return
        
        formula_info = self.formula_fields[field_name]
        expression = formula_info['expression']
        field_type = formula_info['type']
        
        # Parse and evaluate the formula
        result_values = []
        
        for index, row in self.original_df.iterrows():
            try:
                # Replace field references with actual values
                evaluated_expression = expression
                
                # Find all field references
                import re
                field_pattern = r'\[([^\]]+)\]'
                referenced_fields = re.findall(field_pattern, expression)
                
                for ref_field in referenced_fields:
                    if ref_field in self.original_df.columns:
                        value = row[ref_field]
                        # Handle different data types
                        if pd.isna(value):
                            value = 0 if 'Number' in field_type else ""
                        elif isinstance(value, str):
                            value = f'"{value}"'
                        
                        evaluated_expression = evaluated_expression.replace(f'[{ref_field}]', str(value))
                    elif ref_field in self.formula_fields:
                        # Handle references to other formula fields
                        if ref_field in self.original_df.columns:
                            value = self.original_df.loc[index, ref_field]
                            if pd.isna(value):
                                value = 0 if 'Number' in field_type else ""
                            evaluated_expression = evaluated_expression.replace(f'[{ref_field}]', str(value))
                
                # Evaluate the expression (simplified evaluation)
                result = self.evaluate_expression(evaluated_expression, field_type)
                result_values.append(result)
                
            except Exception as e:
                # Use default value on error
                default_value = 0 if field_type == "Number" else ""
                result_values.append(default_value)
        
        # Add the calculated column to dataframes
        self.original_df[field_name] = result_values
        
        # Update working dataframe
        if self.visible_columns and field_name not in self.visible_columns:
            self.visible_columns.append(field_name)
        
        if self.visible_columns:
            # Ensure all visible columns exist in original_df
            available_visible = [col for col in self.visible_columns if col in self.original_df.columns]
            self.df = self.original_df[available_visible].copy()
        else:
            self.df = self.original_df.copy()
        
        # Reapply filters if any
        if self.active_filters:
            self.apply_filters()
    
    def evaluate_expression(self, expression, field_type):
        """Safely evaluate a formula expression"""
        try:
            # Handle common Excel-like functions
            expression = expression.replace('IF(', 'if_func(')
            expression = expression.replace('MAX(', 'max(')
            expression = expression.replace('MIN(', 'min(')
            expression = expression.replace('ROUND(', 'round(')
            expression = expression.replace('ABS(', 'abs(')
            
            # Define safe evaluation context
            safe_dict = {
                "__builtins__": {},
                "max": max,
                "min": min,
                "round": round,
                "abs": abs,
                "if_func": lambda condition, true_val, false_val: true_val if condition else false_val
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
        """Create Formatting tab like Revit"""
        formatting_frame = ttk.Frame(notebook)
        notebook.add(formatting_frame, text="Formatting")
        
        main_frame = ttk.Frame(formatting_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Fields list
        left_frame = ttk.LabelFrame(main_frame, text="Fields:", padding="5")
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        self.formatting_fields_listbox = tk.Listbox(left_frame)
        self.formatting_fields_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Populate fields
        if self.df is not None:
            for col in self.df.columns:
                self.formatting_fields_listbox.insert(tk.END, col)
        
        # Formatting options
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Heading
        heading_frame = ttk.LabelFrame(right_frame, text="Heading:", padding="5")
        heading_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(heading_frame, text="Heading:").grid(row=0, column=0, sticky=tk.W)
        heading_entry = ttk.Entry(heading_frame, width=30)
        heading_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(heading_frame, text="Heading orientation:").grid(row=1, column=0, sticky=tk.W, pady=5)
        orientation_combo = ttk.Combobox(heading_frame, values=["Horizontal", "Vertical"], width=27)
        orientation_combo.set("Horizontal")
        orientation_combo.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(heading_frame, text="Alignment:").grid(row=2, column=0, sticky=tk.W)
        alignment_combo = ttk.Combobox(heading_frame, values=["Left", "Center", "Right"], width=27)
        alignment_combo.set("Left")
        alignment_combo.grid(row=2, column=1, padx=5)
        
        # Field formatting
        field_frame = ttk.LabelFrame(right_frame, text="Field formatting:", padding="5")
        field_frame.pack(fill=tk.X)
        
        ttk.Checkbutton(field_frame, text="Hidden field").pack(anchor=tk.W)
        ttk.Checkbutton(field_frame, text="Show conditional format on sheets").pack(anchor=tk.W)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=2)
        main_frame.rowconfigure(0, weight=1)
        
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
        grid_var = tk.BooleanVar(value=self.appearance_settings['grid_lines'])
        ttk.Checkbutton(graphics_frame, text="Grid lines", variable=grid_var).pack(anchor=tk.W)
        
        # Outline
        outline_var = tk.BooleanVar(value=self.appearance_settings['outline'])
        ttk.Checkbutton(graphics_frame, text="Outline", variable=outline_var).pack(anchor=tk.W)
        
        # Text section
        text_frame = ttk.LabelFrame(main_frame, text="Text", padding="5")
        text_frame.pack(fill=tk.X)
        
        # Show title
        title_var = tk.BooleanVar(value=self.appearance_settings['show_title'])
        ttk.Checkbutton(text_frame, text="Show Title", variable=title_var).pack(anchor=tk.W)
        
        # Show headers
        headers_var = tk.BooleanVar(value=self.appearance_settings['show_headers'])
        ttk.Checkbutton(text_frame, text="Show Headers", variable=headers_var).pack(anchor=tk.W)
        
        # Store variables for later use
        self.appearance_vars = {
            'grid_lines': grid_var,
            'outline': outline_var,
            'show_title': title_var,
            'show_headers': headers_var
        }
        
    # Field management methods
    def add_field_to_schedule(self):
        """Add selected field to schedule (make it visible)"""
        selection = self.available_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to add.")
            return
            
        for index in selection:
            field_name = self.available_fields_listbox.get(index)
            # Check if field is already in scheduled fields
            scheduled_items = [self.scheduled_fields_listbox.get(i) for i in range(self.scheduled_fields_listbox.size())]
            if field_name not in scheduled_items:
                self.scheduled_fields_listbox.insert(tk.END, field_name)
                
    def remove_field_from_schedule(self):
        """Remove selected field from schedule (hide it, but keep in available fields)"""
        selection = self.scheduled_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to remove.")
            return
            
        # Remove selected items (in reverse order to maintain indices)
        for index in reversed(selection):
            self.scheduled_fields_listbox.delete(index)
            
        # Note: Field remains in original_df and available_fields_listbox
        # It's just hidden from the schedule view
            
    def move_field_up(self):
        """Move selected field up in schedule"""
        selection = self.scheduled_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to move.")
            return
            
        if selection[0] > 0:
            index = selection[0]
            item = self.scheduled_fields_listbox.get(index)
            self.scheduled_fields_listbox.delete(index)
            self.scheduled_fields_listbox.insert(index - 1, item)
            self.scheduled_fields_listbox.selection_set(index - 1)
            
    def move_field_down(self):
        """Move selected field down in schedule"""
        selection = self.scheduled_fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a field to move.")
            return
            
        if selection[0] < self.scheduled_fields_listbox.size() - 1:
            index = selection[0]
            item = self.scheduled_fields_listbox.get(index)
            self.scheduled_fields_listbox.delete(index)
            self.scheduled_fields_listbox.insert(index + 1, item)
            self.scheduled_fields_listbox.selection_set(index + 1)
            
    def add_filter_row(self):
        """Legacy method - redirects to Schedule Properties"""
        messagebox.showinfo("Filter", "Please use 'Schedule Properties' to add filters.\n\nClick the 'Schedule Properties' button and go to the 'Filter' tab.")
        self.open_schedule_properties()
        
    def apply_sorting(self):
        """Apply sorting based on sort settings"""
        if self.df is None or not hasattr(self, 'sort_rows'):
            return
            
        # Get active sort criteria
        sort_columns = []
        sort_ascending = []
        
        for sort_info in self.sort_rows:
            column = sort_info['column_var'].get()
            if column and column != '(none)':
                # Check if column exists in visible columns or original dataframe
                if column in self.visible_columns or (self.visible_columns and column in self.original_df.columns):
                    sort_columns.append(column)
                    sort_ascending.append(sort_info['direction_var'].get() == "Ascending")
        
        # Apply sorting to the appropriate dataframe
        if sort_columns:
            try:
                # Determine which dataframe to sort
                if self.filtered_df is not None:
                    # Sort the filtered dataframe
                    sorted_df = self.filtered_df.sort_values(by=sort_columns, ascending=sort_ascending)
                    self.filtered_df = sorted_df
                else:
                    # Sort the main working dataframe
                    sorted_df = self.df.sort_values(by=sort_columns, ascending=sort_ascending)
                    self.df = sorted_df
                    
                # Also sort the original dataframe to maintain consistency
                if sort_columns:
                    available_sort_cols = [col for col in sort_columns if col in self.original_df.columns]
                    if available_sort_cols:
                        available_ascending = [sort_ascending[i] for i, col in enumerate(sort_columns) if col in available_sort_cols]
                        self.original_df = self.original_df.sort_values(by=available_sort_cols, ascending=available_ascending)
                        
                        # Update the working dataframe with new order
                        if self.visible_columns:
                            self.df = self.original_df[self.visible_columns].copy()
                        else:
                            self.df = self.original_df.copy()
                            
                self.modified = True
                
            except Exception as e:
                messagebox.showerror("Sorting Error", f"Failed to apply sorting:\n{str(e)}")
                
        # Store current sort settings for future use
        self.current_sort_columns = sort_columns
        self.current_sort_ascending = sort_ascending
                
    def add_parameter(self):
        """Add a new parameter/column"""
        self.add_column()
        
    def remove_parameter(self):
        """Remove a parameter/column"""
        self.delete_column()


def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = XLSEditor(root)
    
    # Handle window closing
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Start the application
    root.mainloop()


if __name__ == "__main__":
    main()