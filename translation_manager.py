"""
Translation Manager Module
Handles internationalization and language switching for the XLS Editor
"""

class TranslationManager:
    def __init__(self):
        self.current_language = "en"  # Default to English
        self.translations = self.load_translations()
    
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
        return f"{self.tr('Language changed to')} {self.tr('English' if language_code == 'en' else 'Vietnamese')}"
    
    def get_current_language(self):
        """Get current language code"""
        return self.current_language