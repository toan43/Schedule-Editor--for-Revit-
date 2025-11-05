# XLS File Editor with Revit Schedule Features

A powerful GUI application for importing, editing, and filtering Excel (.xls/.xlsx) files with complete Revit Schedule-like functionality including **Formula/Calculated Fields**.

## Features

- **Import XLS/XLSX files**: Load Excel files into the application
- **Edit data**: Double-click any cell to edit its value
- **Add/Delete rows**: Add new rows or delete selected rows
- **Add/Delete columns**: Add new columns or delete existing ones
- **Advanced Filtering**: Revit Schedule-like filtering with multiple filter types
- **Schedule Properties**: Complete Revit-style schedule management
- **Fields Management**: Add/remove/reorder fields like Revit parameters
- **Advanced Sorting**: Multi-level sorting with grouping options
- **🆕 Formula Columns**: Create calculated fields using mathematical expressions (Revit-like)
- **Appearance Settings**: Graphics, text, and layout options
- **Save functionality**: Save changes to the original file or save as a new file
- **Visual feedback**: Shows file modification status and operation feedback

## 🚀 NEW: Enhanced Formula Feature (Revit-Like Calculated Fields)

### **Formula Menu:**
- **Add Formula Column** - Create new calculated columns with enhanced interface
- **Manage Formulas** - View, refresh, and remove formula columns

### **Enhanced Formula Interface:**
- **Interactive Field Selection** - Choose from available fields list
- **One-Click Operators** - Insert +, -, *, /, **, (), abs(), round() with buttons
- **Real-time Preview** - See formula as you build it
- **Smart Cursor Positioning** - Optimal cursor placement after insertions
- **Double-click Insertion** - Quick field reference insertion

### **Formula Syntax:**
- Use square brackets for column references: `[Column_Name]`
- Supported operations: `+`, `-`, `*`, `/`, `**`, `()`, `abs()`, `round()`, `max()`, `min()`
- Examples:
  - `[Length] * [Width]` - Calculate area
  - `[Price] + [Tax]` - Add tax to price
  - `round([Area] / 10.764, 2)` - Convert sq ft to sq m
  - `abs([Value1] - [Value2])` - Absolute difference

### **Enhanced Formula Creation:**
- **Available Fields Panel** - Lists all data columns for easy selection
- **Operator Buttons** - Click to insert mathematical operators
- **Function Buttons** - Pre-formatted function templates
- **Formula Testing** - Validate before applying
- **Error Prevention** - Guided field selection reduces typing errors

## Schedule Properties (Revit-Style Features)

### **4-Tab Interface:** *(Formatting tab removed, Formula tab added)*
1. **Fields Tab** - Manage available and scheduled fields
2. **Filter Tab** - Advanced multi-condition filtering  
3. **Sorting/Grouping Tab** - Multi-level sorting with headers/footers
4. **🆕 Formula Tab** - Create and manage calculated fields
5. **Appearance Tab** - Graphics, grid lines, and text settings

### **Fields Management:**
- **Available Fields List** - Shows all possible columns/parameters
- **Scheduled Fields List** - Shows active columns in order
- **Drag & Drop Style** - Add/Remove buttons between lists
- **Reorder Fields** - Move up/down buttons for field order
- **Real-time Preview** - Changes apply immediately

### **Advanced Filtering:**
- **Multiple Filter Rows** - Add unlimited filter conditions
- **AND Logic** - All conditions must be met
- **12 Filter Types** - Same as before (equals, contains, etc.)
- **Dynamic Fields** - Filter dropdown updates with available columns
- **Remove Filters** - Individual filter row removal

### **Sorting & Grouping:**
- **4-Level Sorting** - Sort by multiple columns simultaneously
- **Ascending/Descending** - Per-column sort direction
- **Headers & Footers** - Group headers and summary footers
- **Blank Lines** - Visual separation between groups
- **Grand Totals** - Summary calculations
- **Itemize Every Instance** - Show individual records

### **Formula Tab (NEW):**
1. **Available Columns** - Shows all columns with [bracket] syntax
2. **Column Name** - Enter name for new calculated column
3. **Formula Expression** - Enter mathematical formula
4. **Test Formula** - Validate formula before applying
5. **Add Formula** - Create the calculated column
6. **Manage Existing** - View, refresh, or remove formulas

### **Formatting Options (REMOVED):**
*The Formatting tab has been removed and replaced with the Formula tab to focus on calculated field functionality like Revit.*

### **Appearance Control:**
- **Graphics Settings**:
  - Grid lines on/off
  - Outline borders
  - Line styles
- **Text Settings**:
  - Show/hide title
  - Show/hide headers
  - Font selections

## Requirements

- Python 3.7+
- Required packages (install via `pip install -r requirements.txt`):
  - pandas
  - openpyxl
  - xlrd

## How to Run

1. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

2. Run the application:
   ```
   python main.py
   ```

## How to Use

1. **Import a file**: Click "Import XLS File" or use File → Import XLS File menu
2. **Edit cells**: Double-click any cell in the data grid to edit its value
3. **Add/Delete rows**: Use the Edit menu or toolbar buttons
4. **Add/Delete columns**: Use the Edit menu options
5. **🆕 Create Formulas**: Use Formula menu or Schedule Properties → Formula tab
6. **Apply filters**: Use Filter menu or filter buttons to narrow down data
7. **Schedule Properties**: Click "Schedule Properties" for advanced Revit-like features
8. **Save**: Use Ctrl+S, the Save button, or File → Save menu
9. **Save As**: Use File → Save As to save with a new filename

## 🔧 Formula Usage Guide

### **Creating Formula Columns:**

#### **Method 1: Enhanced Formula Menu**
1. Go to **Formula → Add Formula Column**
2. **Enhanced Interface**:
   - Enter column name in left panel
   - Select fields from "Available Fields" list
   - Double-click fields to insert or use "Insert Field" button
   - Use operator buttons: +, -, *, /, **, (, )
   - Use function buttons: abs(), round(), max(), min()
   - Preview formula in real-time
3. Click "Test Formula" to validate
4. Click "Add Column" to create

#### **Method 2: Schedule Properties - Enhanced Formula Tab**
1. Click **"Schedule Properties"** button
2. Go to **"Formula"** tab
3. **Interactive Formula Building**:
   - Use "Available Fields" panel for field selection
   - Click operator and function buttons
   - Build formulas with guided assistance
4. Test and add formula
5. Click "Apply" to apply changes

### **Formula Examples:**
- **Area**: `[Length] * [Width]`
- **Volume**: `[Length] * [Width] * [Height]`
- **Total Cost**: `[Price] * [Quantity]`
- **Cost per Area**: `round([Total_Cost] / [Area], 2)`
- **Max Dimension**: `max([Length], [Width])`
- **Tax Included**: `[Price] * 1.08`

### **Managing Formulas:**
- Use **Formula → Manage Formulas** to view all formula columns
- Refresh formulas when source data changes
- Remove individual formulas or clear all formulas
- All formula definitions are preserved when saving files

## Schedule Properties Guide (Revit-Style)

### **Opening Schedule Properties:**
- Click **"Schedule Properties"** button in filter section
- Or use **Schedule → Schedule Properties** menu
- Opens tabbed dialog with 5 tabs (Formula tab replaces Formatting tab)

### **Fields Tab:**
1. **Left List** - Available fields/columns
2. **Add/Remove** - Move fields between lists  
3. **Right List** - Active scheduled fields
4. **Reorder** - Use ↑↓ buttons to change column order
5. **Apply** - Updates the data view with new field order

### **Filter Tab:**
1. **Add Filter Row** - Click to add new filter condition
2. **Select Field** - Choose column to filter
3. **Select Type** - Choose filter operation (equals, contains, etc.)
4. **Enter Value** - Type filter value
5. **Multiple Filters** - All conditions work together (AND logic)

### **Sorting/Grouping Tab:**
1. **Sort by** - Primary sort column
2. **Then by** - Secondary, tertiary, quaternary sorts
3. **Ascending/Descending** - Choose sort direction
4. **Headers/Footers** - Add group separators (Headers and Footers options)
5. **Blank Line** - Visual separation between groups
6. **Grand Totals** - Enable summary calculations option
7. **Itemize Every Instance** - Show individual records option

### **🆕 Formula Tab:**
1. **Existing Formulas** - View all current formula columns
2. **Column Name** - Enter name for new calculated field
3. **Formula Expression** - Enter mathematical formula using [Column] syntax
4. **Available Columns** - Reference list showing all available columns
5. **Test Formula** - Validate formula without creating column
6. **Add Formula** - Create the calculated column
7. **Remove Selected** - Delete existing formula columns

### **Appearance Tab:**
1. **Grid Lines** - Show/hide cell borders
2. **Outline** - Show/hide table border
3. **Show Title** - Display table title
4. **Show Headers** - Display column headers

## Features Overview

- **File Management**: Import, save, and save-as functionality
- **Data Editing**: In-place cell editing with type conversion
- **Row Operations**: Add empty rows or delete selected rows
- **Column Operations**: Add new columns or delete existing ones
- **Advanced Filtering**: 12 filter types with case sensitivity options
- **Filter Management**: Add, remove, and manage multiple simultaneous filters
- **Unique Values Preview**: See all unique values in a column before filtering
- **Real-time Updates**: Filters apply immediately with row count display
- **Visual Indicators**: Modified files marked (*), active filter count shown
- **Error Handling**: Comprehensive error messages and validation

## Supported File Formats

- .xlsx (Excel 2007+)
- .xls (Excel 97-2003)
*CONDITION
-add condition in formula (like the column,row must satisfied the conditon like have this character eg:"C1-400x550" folowing the
horizontal row not the column
Omniclass	    Type        Element	Type	Diameter	L1	L2	Steel-Total-Length	Belt-Step
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C1-400x550	Main-Steel	16	640	100		
23-1331211313	C2-350x500	Main-Steel	8			1470	200
23-1331211313	C2-350x501	Main-Steel	8			1470	200
23-1331211313	C2-350x502	Main-Steel	8			1470	200

like in here i want cal the L1 + L2 but only in if they have C1-400x550 in the Type column

### **Cross-Sheet Formulas with HAS_VALUE:**

The `HAS_VALUE` function checks if a value exists ANYWHERE in a specified column of another sheet. This is perfect for conditional calculations that depend on whether certain data exists.

**Example Formula** (can be placed on Sheet2 to calculate from Sheet1 data):
```
IF(HAS_VALUE(Sheet1, "Type Element", "C1-400x550"), (COUNT(Sheet1.Main-Steel) * Sheet1.[Diameter(0)] * (Sheet1.[Diameter(0)]/1000) * (Sheet1.[Diameter(0)]/1000) * 3.141592 / 4 * 7850.3 * ((10300 + Sheet1.[L1(0)] + Sheet1.[L2(0)])/1000))+(COUNT(Sheet1.Belt Steel)*(Sheet1.[Steel-Total-Length(14)]/1000)*(10300/Sheet1.[Belt-Step(14)])*(Sheet1.[Diameter(14)]/1000) * (Sheet1.[Diameter(14)]/1000)*3.141592/4*7850.3),0)
```

**How it works:**
1. `HAS_VALUE(Sheet1, "Type Element", "C1-400x550")` - Checks if "C1-400x550" exists in Sheet1's "Type Element" column
2. If TRUE (data exists):
   - Calculates main steel weight using filtered data
   - Calculates belt steel weight using filtered data
   - Returns the sum
3. If FALSE (data doesn't exist):
   - Returns 0

**Key Features:**
- ✅ Works cross-sheet (formula on Sheet2, data from Sheet1)
- ✅ HAS_VALUE checks the ENTIRE sheet, not just current row
- ✅ Implicit filtering: `Sheet1.[Diameter(0)]` gets the first value WHERE Type Element="C1-400x550"
- ✅ Can combine multiple COUNT and fixed value lookups in one formula

**Important Notes:**
- When using cross-sheet references, always specify the sheet name (e.g., `Sheet1.Main-Steel`, not just `Main-Steel`)
- Column names with spaces or special characters should use the exact name from the CSV/Excel file
- `Sheet.[Column(index)]` with HAS_VALUE will automatically filter to matching rows before indexing"# Schedule-Editor--for-Revit-" 
