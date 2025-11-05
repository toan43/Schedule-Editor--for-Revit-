# Project Roadmap and Function Reference

This document provides a high-level roadmap for the repository and a detailed reference of the public functions/classes in each Python module.

Date: 2025-11-01

## Purpose
This workspace contains a modular XLS editor with filtering, multi-sheet support, schedule (Revit-like) properties and a formula engine. Use this roadmap to understand module responsibilities and find specific functions quickly.

---

## Files & Responsibilities (overview)

- `main.py` — Application entrypoint and central coordinator (XLSEditor). Wires modules together and builds the GUI.
- `file_operations.py` — Import/export, save, save-as and closing logic.
- `data_management.py` — Data display and editing logic (Treeview population, cell edit, add/delete rows/columns).
- `filter_operations.py` — Filter dialogs and applying/clearing filters.
- `formula_operations.py` — Formula validation, parsing and calculation engine. Manages formula templates and formula fields.
- `schedule_properties.py` — The Revit-like dialog (Fields, Filter, Sorting, Formula, Appearance tabs) and related UI handlers.
- `sheet_operations.py` — Multi-sheet Excel handling, loading multiple sheets, sheet switching and cross-sheet formula support.
- `translation_manager.py` — Simple i18n manager for English/Vietnamese translations.
- `debug_formula.py` — Local helper script used to debug and test formulas using `2D_element.xlsx`.
- `2D_element.xlsx` — Example dataset used during debugging (not listed here as a code file).

---

## Detailed Function Reference

Below each module lists exported classes and their important methods/functions with a short description and any important notes or edge-cases.

### main.py — XLSEditor
Class: `XLSEditor`
- `__init__(self, root)` — Initialize application, modules and GUI.
- `tr(self, text)` — Localization helper (delegates to `TranslationManager.tr`).
- `change_language(self, language_code)` — Change language and refresh UI.
- `refresh_interface(self)` — Rebuilds menus/widgets and refreshes displays.
- `create_menu(self)` — Builds the menubar and routes commands to modules.
- `create_widgets(self)` — Builds main window layout including treeview, controls and status bar.
- `update_header_display(self)` — Refresh header row label.
- `set_header_row(self)` — Dialog to let user pick header row; reloads current file with chosen header.
- `add_parameter(self)` / `remove_parameter(self)` — Glue to `DataManagement.add_column/delete_column`.
- `sync_current_sheet_data(self)` — Save current in-memory df into `sheet_ops.available_sheets` when multi-sheet active.
- `create_cross_sheet_formula_dialog(self)` — GUI dialog to create a formula field on a selected target sheet.
- `save_all_sheets(self)` — Ask-for-path then save all loaded sheets via `SheetOperations.save_all_sheets`.
- `main()` (module-level) — Starts the Tk main loop and app.

Notes: `XLSEditor` is the central orchestrator — UI actions call into module instances stored on `self`.

---

### file_operations.py — FileOperations
Class: `FileOperations`
- `__init__(self, editor_instance)` — Keep reference to editor.
- `import_file(self)` — Open file dialog and import Excel into `editor.original_df`; sets up `editor.df` and resets filters/formula fields.
- `save_file(self)` — Save current working df or all sheets (via `SheetOperations`) back to `editor.current_file`.
- `save_as_file(self)` — Save-as flow, supports both single and multi-sheet saves.
- `update_file_info(self)` — Update filename label and modified indicator.
- `on_closing(self)` — Graceful shutdown: save templates, prompt for unsaved changes and close the window.

Notes: uses pandas + openpyxl/xlrd. Handles both single and multi-sheet saving.

---

### data_management.py — DataManagement
Class: `DataManagement`
- `__init__(self, editor_instance)`
- `populate_treeview(self)` — Populate the Treeview widget with `editor.df` or `editor.filtered_df`. Applies `visible_columns`.
- `on_cell_double_click(self, event)` — Map GUI double-click to `edit_cell` for that row/column.
- `edit_cell(self, row_index, col_index, current_value)` — Dialog to edit a specific cell; converts numeric strings to numbers when possible.
- `add_row(self)` / `delete_row(self)` — Add or delete rows, update `editor.df` and refresh view.
- `add_column(self)` / `delete_column(self)` — Add or drop columns via simple dialogs.

Notes: `populate_treeview` uses the `editor.visible_columns` ordering. Edits call `editor.file_ops.update_file_info()` to mark modified state.

---

### filter_operations.py — FilterOperations
Class: `FilterOperations`
- `__init__(self, editor_instance)`
- `add_filter(self)` — Dialog to build and add a new filter; provides a preview of unique values.
- `apply_filters(self)` — Applies all `editor.active_filters` to `editor.df` and sets `editor.filtered_df`.
- `clear_all_filters(self)` — Clear all active filters with confirmation.
- `manage_filters(self)` — Dialog to view and remove active filters.
- `update_filter_display(self)` — Update the filter status label.

Notes: Supports many filter operations (equals, contains, numeric comparisons, is empty etc.). Uses case-insensitive matching by default.

---

### formula_operations.py — FormulaOperations (formula engine)
Class: `FormulaOperations`
- `__init__(self, editor_instance)` — Loads formula templates on init.
- `save_formula_templates_to_file(self)` / `load_formula_templates_from_file(self)` — Persist templates to `formula_templates.json`.
- `refresh_formula_tree(self)` / `refresh_formula_templates(self)` — Update UI lists with current formulas/templates.
- `insert_field_in_formula(self, event)` — Insert clicked/selected field into the formula expression input.
- `load_formula_template(self, event)` — Load saved template into the creation form.
- `create_formula_field(self)` — Validate and add a new formula field; calls `calculate_formula_field`.
- `update_formula_field(self)` / `delete_formula_field(self)` — Update or remove existing formula fields (keeps dataframes consistent).
- `save_formula_template(self)` — Save a template via a simple dialog.
- `refresh_all_formulas(self)` — Recalculate all formula fields.

Validation & Calculation:
- `validate_formula(self, expression)` — Validates multiple syntaxes:
    - `[Field]`: Standard field on the current sheet.
    - `Sheet.Field`: Row-by-row reference to another sheet.
    - `Sheet.[Column(index)]`: **New!** Fixed value reference with implicit filtering support.
    - `COUNT(...)` variants: `COUNT([Field])`, `COUNT(Value)`, `COUNT(Sheet.Value)`.
    - `HAS_VALUE(Sheet, "Column", "Value")`: Conditional check on another sheet.
    - `LOOKUP(Sheet, "ColumnToGet", "FilterColumn", "FilterValue")`: **New!** Lookup function to find first non-zero value.
- `calculate_formula_field(self, field_name)` — Core calculation logic, now with intelligent pre-processing:
  - **Implicit Filter Context**: When a formula contains `HAS_VALUE(Sheet, "Column", "Value")`, the system automatically creates a filter context. All `Sheet.[Column(index)]` lookups on that sheet will use the filtered data instead of the full sheet.
  - **Pre-processing Steps** (executed once before row-by-row evaluation):
    1. Detects `HAS_VALUE` and creates filter context for the referenced sheet.
    2. Replaces `Sheet.[Column(index)]` with actual values from the filtered or full dataframe.
    3. Replaces `COUNT(Value)` and `COUNT(Sheet.Value)` with their calculated counts.
    4. Replaces `LOOKUP(...)` calls with the first non-zero value found in the filtered data.
  - **Row-by-row Evaluation**:
    - Handles `IF(HAS_VALUE(...), ..., ...)` by checking the condition for each row.
    - Replaces row-dependent references like `[Field]` and `Sheet.Field` with values from the current row index.
    - Replaces `COUNT([Field])` with the count of the current row's value in that field.
    - Evaluates the final expression safely.
- `evaluate_expression(self, expression, field_type)` — Safe eval wrapper supporting IF, MAX, MIN, ROUND, ABS and other functions.

Formula Syntax Reference:

**1. Fixed Value Lookup with Implicit Filtering: `Sheet.[Column(index)]`**
- When used with `HAS_VALUE`, the index is relative to the filtered data.
- Example: `IF(HAS_VALUE(Sheet1, "Type", "Belt Steel"), Sheet1.[Diameter(0)] * 100, 0)`
  - If "Belt Steel" exists in Sheet1's Type column, `Sheet1.[Diameter(0)]` will return the Diameter value from the *first row where Type="Belt Steel"*, not the first row of the entire sheet.
- Without `HAS_VALUE`, index refers to the absolute row position in the sheet.

**2. LOOKUP Function: `LOOKUP(Sheet, "ColumnToGet", "FilterColumn", "FilterValue")`**
- Filters the specified sheet to find rows where `FilterColumn` equals `FilterValue`.
- Returns the first non-zero, non-empty value found in `ColumnToGet` from the filtered rows.
- Returns 0 if no matching rows or no non-zero values are found.
- Example: `LOOKUP(Sheet1, "main-steel-weight", "Type", "Main-Steel") + LOOKUP(Sheet1, "belt-steel-weight", "Type", "Belt Steel")`
  - This finds the calculated weight for "Main-Steel" and "Belt Steel" separately, then adds them together.
  - Perfect for summing results from different formula columns that have been pre-calculated for different categories.

**3. COUNT Variants**
- `COUNT([Field])`: Counts how many rows have the same value as the current row in the specified field (row-dependent).
- `COUNT(Value)`: Counts how many cells in the entire current sheet equal the literal value (constant).
- `COUNT(Sheet.Value)`: Counts how many cells in another sheet equal the literal value (constant).

**4. HAS_VALUE Function: `HAS_VALUE(Sheet, "Column", "Value")`**
- **Checks if ANY row in the entire specified sheet has the given value in the specified column.**
- Returns `True` if at least one matching row is found, `False` otherwise.
- **Works across sheets**: Can be used on Sheet2 to check conditions on Sheet1.
- Used in IF statements to conditionally execute formulas.
- Creates an implicit filter context for `Sheet.[Column(index)]` lookups on the same sheet.
- Example: `IF(HAS_VALUE(Sheet1, "Type", "C1-400x550"), ... , 0)`
  - This checks if "C1-400x550" exists ANYWHERE in Sheet1's "Type" column.
  - If found, executes the true branch; otherwise returns 0.
  - **This works even when the formula is placed on Sheet2 or any other sheet.**

Notes & edge-cases:
- **HAS_VALUE is Sheet-Global**: Unlike row-by-row functions, `HAS_VALUE(Sheet, "Column", "Value")` checks the **entire sheet** for any matching row. This means:
  - You can use it on Sheet2 to check conditions on Sheet1.
  - It returns the **same result** for all rows (True or False based on whether the value exists anywhere in the specified sheet).
  - Perfect for conditional calculations that depend on whether certain data exists in another sheet.
- **Implicit Filter Context**: The filter context from `HAS_VALUE` only applies to `Sheet.[Column(index)]` lookups for the *same sheet* mentioned in `HAS_VALUE`. Other sheets are not affected.
- **LOOKUP vs Sheet.[Column(index)]**: Use `LOOKUP` when you want to find a calculated result from another formula column. Use `Sheet.[Column(index)]` when you want to reference raw data values.
- **Performance**: Fixed-value lookups (`Sheet.[Column(index)]` and `LOOKUP`) are pre-calculated once per formula field, making them very efficient even with large datasets.
- Cross-sheet `Sheet.Field` substitution (without `[index]`) maps by row index for row-by-row calculations.
- If formula evaluation fails on a row, it defaults to 0 for numeric fields.
- **Cross-Sheet Formula Usage**: When placing formulas on Sheet2 that reference Sheet1:
  - Always use explicit sheet references (e.g., `Sheet1.[Column(index)]`, `HAS_VALUE(Sheet1, ...)`, `COUNT(Sheet1.Value)`).
  - The formula will be evaluated for each row on Sheet2, but the lookups will pull data from Sheet1.
  - This allows you to create summary or calculation sheets that aggregate data from multiple source sheets.

---

### schedule_properties.py — ScheduleProperties (UI heavy)
Class: `ScheduleProperties`
- `__init__(self, editor_instance)`
- `open_schedule_properties(self)` — Build the main Schedule Properties dialog and tabs.
- Tab builders: `create_fields_tab`, `create_filter_tab`, `create_sorting_tab`, `create_formula_tab`, `create_appearance_tab` — Each builds UI for that tab and hooks to editor/formula/filter modules.
- Field management: `add_field_to_schedule`, `remove_field_from_schedule`, `move_field_up`, `move_field_down` — Manage scheduled (visible) fields ordering.
- Sorting helpers: `create_sort_row`, `apply_sorting`, `preview_sorting`, `clear_sorting`, `update_sort_status`.
- Filter helpers inside tab: `refresh_filter_tree`, `add_filter_from_schedule_properties`, `remove_selected_filter`, `clear_all_filters_and_refresh`.

Notes: This module is UI-centric and coordinates other modules when user applies changes.

---

### sheet_operations.py — SheetOperations
Class: `SheetOperations`
- `__init__(self, editor_instance)` — Holds `available_sheets` dict and `current_sheet` state.
- `get_sheet_names(self, file_path)` — Return sheet names for a file using pandas.
- `import_file_with_sheet_selection(self)` and `show_sheet_selection_dialog(self, file_path, sheet_names)` — UI to select sheets to import.
- `load_multiple_sheets(self, file_path, sheet_names, primary_sheet)` — Load DataFrames for each sheet, set `primary_sheet` as current working df and refresh editor state.
- `load_sheet(self, file_path, sheet_name)` — Load a single sheet into `available_sheets` and set as current.
- `add_sheet_switcher(self)` / `switch_sheet(self, event=None)` — Add sheet selector to main UI and handle switching, saving previous sheet data back to `available_sheets`.
- `get_available_sheets_for_formula`, `get_sheet_columns`, `get_sheet_data` — Helpers to expose sheet metadata.
- `create_cross_sheet_formula(self, target_sheet, formula_field_name, formula_expression)` — Basic cross-sheet formula processor (currently replaces references and evaluates via `eval` and inserts results into target sheet).
- `get_cross_sheet_fields_for_schedule_properties(self)`, `save_all_sheets(self, file_path)` — Utilities for schedule UI and saving.

Notes: `create_cross_sheet_formula` currently uses simplistic replacement and uses the first row for cross-sheet lookups by default — formula engine in `formula_operations` offers more flexible features.

---

### translation_manager.py — TranslationManager
Class: `TranslationManager`
- `__init__(self)` — Load translations (hard-coded dict for 'en' and 'vi').
- `load_translations(self)` — Return the translations structure.
- `tr(self, text)` — Translate text according to `current_language`.
- `change_language(self, language_code)` — Switch language and return status message.
- `get_current_language(self)` — Get language code.

Notes: Lightweight i18n implementation, centralizes UI strings.

---

### debug_formula.py
- `calculate_formula(row_index_to_debug=0)` — Example script that reads `2D_element.xlsx` and computes the complex steel-weight formula for a specific row, printing intermediate values. Useful for reproducing and debugging formula behavior outside the GUI.

---

## How to use this Roadmap

- To locate a function: search for the file name above, then scan the listed functions.
- To update behavior of formula parsing/counting: edit `formula_operations.py` — validation and calculation code are concentrated in `validate_formula` and `calculate_formula_field`.
- To adjust multi-sheet loading behavior (row-alignment, lookups): check `sheet_operations.create_cross_sheet_formula` and `sheet_operations.load_multiple_sheets`.

---

## Next steps / Recommendations

1. Add unit tests for `formula_operations.calculate_formula_field` to cover COUNT variations and cross-sheet references.
2. Improve cross-sheet lookups: support key-based joins instead of index-based substitution.
3. Harden `evaluate_expression` to avoid `eval` where practical or to add stricter sandboxing for unusual inputs.
4. Add documentation comments (docstrings) to functions missing them (many UI callbacks are not fully documented).

---

If you want, I can:
- Generate a more detailed per-function signature list (parameters, return types and exceptions).
- Create unit tests for the formula engine.
- Create quick-start usage examples for the most common formulas.

