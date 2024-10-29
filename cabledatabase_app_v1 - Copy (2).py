import pandas as pd
import PySimpleGUI as sg
import openpyxl
import os
import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
from thefuzz import fuzz
import re
import time

# Constants
DEFAULT_SETTINGS = {
    'last_file_path': '',
    'default_file_path': '',
    'auto_load_default': True,
    'last_directory': '',
}

# UI Constants
UI_CONSTANTS = {
    'WINDOW_SIZE': (1200, 800),
    'WINDOW_LOCATION': (100, 100),
    'INPUT_SIZE': (15, 1),
    'LABEL_SIZE': (8, 1),
    'BUTTON_COLORS': ('white', 'navy'),
    'FILTER_KEYS': {
        'NUMBER': '-NUM-START-',
        'DWG': '-DWG-',
        'ORIGIN': '-ORIGIN-',
        'DEST': '-DEST-',
        'Wire Type': '-WIRE-TYPE-',
        'Length': '-LENGTH-',
        'Project ID': '-PROJECT-'
    },
    'TABLE_COLS': [
        'NUMBER', 'DWG', 'ORIGIN', 'DEST',
        'Alternate Dwg', 'Wire Type', 'Length',
        'Note', 'Project ID'
    ]
}

# Add these functions at the module level (near the top of the file)
def load_column_mapping() -> Dict[str, str]:
    """Load saved column mapping"""
    try:
        with open('config/column_mapping.json', 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_column_mapping(mapping: Dict[str, str]):
    """Save column mapping to settings file"""
    settings_path = Path('config/column_mapping.json')
    settings_path.parent.mkdir(exist_ok=True)
    with open(settings_path, 'w') as f:
        json.dump(mapping, f, indent=4)

def show_column_mapping_dialog(excel_columns: List[str], required_columns: List[str]) -> Dict[str, str]:
    """Show dialog for mapping Excel columns to required database fields"""
    layout = [
        [sg.Text("Map Your Excel Columns to Required Fields", font=('Any', 12, 'bold'))],
        [sg.Text("Please match each required field with the corresponding column from your Excel file:")],
        [sg.Text("_" * 80)],
    ]
    
    # Create mapping inputs for each required column
    mappings = {}
    for req_col in required_columns:
        # Try to find a close match in excel_columns
        default_match = next(
            (col for col in excel_columns 
             if col.upper().replace(" ", "") == req_col.upper().replace(" ", "")),
            excel_columns[0] if excel_columns else ""
        )
        
        layout.append([
            sg.Text(f"{req_col}:", size=(15, 1)),
            sg.Combo(
                excel_columns,
                default_value=default_match,
                key=f'-MAP-{req_col}-',
                size=(30, 1)
            ),
            sg.Checkbox("Skip this field", key=f'-SKIP-{req_col}-')
        ])
    
    layout.extend([
        [sg.Text("_" * 80)],
        [sg.Checkbox("Save this mapping for future use", key='-SAVE-MAPPING-', default=True)],
        [sg.Button("Apply Mapping"), sg.Button("Cancel")]
    ])
    
    window = sg.Window("Column Mapping", layout, modal=True, finalize=True)
    
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, "Cancel"):
            window.close()
            return None
            
        if event == "Apply Mapping":
            # Create mapping dictionary
            mapping = {}
            for req_col in required_columns:
                if not values[f'-SKIP-{req_col}-']:  # If field is not skipped
                    excel_col = values[f'-MAP-{req_col}-']
                    if excel_col:
                        mapping[req_col] = excel_col
            
            # Save mapping if requested
            if values['-SAVE-MAPPING-']:
                save_column_mapping(mapping)
            
            window.close()
            return mapping
    
    window.close()
    return None

class Settings:
    def __init__(self):
        self.settings_file = 'cable_db_settings.json'
        self.settings = self.load_settings()
        
    def load_settings(self) -> Dict:
        """Load settings from file or create default"""
        try:
            with open(self.settings_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            # First time setup - create default settings
            default_settings = {
                'first_run': False,  # Add this flag
                'theme': 'Dark',
                'last_directory': os.getcwd(),
                # ... other settings ...
            }
            self.save_settings(default_settings)
            return default_settings
        except json.JSONDecodeError:
            # Handle corrupted settings file
            return self.create_default_settings()

    def save_settings(self, settings: Dict = None) -> None:
        """Save settings without showing dialog"""
        if settings:
            self.settings = settings
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f)
        except Exception as e:
            print(f"Error saving settings: {str(e)}")

    def save_color_settings(self, values):
        """Save color settings to config"""
        for i in range(6):
            color_key = f'color{i+1}'
            self.settings['color_categories'][color_key] = {
                'color': values[f'-COLOR{i+1}-'],
                'keywords': values[f'-KEYWORDS{i+1}-'].split(',')
            }
        self.save_settings()

# Basic utility functions
def load_last_file_path():
    try:
        with open('last_file_path.json', 'r') as f:
            return json.load(f).get('last_path', '')
    except FileNotFoundError:
        return ''

def save_last_file_path(file_path):
    with open('last_file_path.json', 'w') as f:
        json.dump({'last_path': file_path}, f)

class DataManager:
    def __init__(self):
        """Initialize the DataManager"""
        self.df = None
        self.length_matrix = None
        self.force_show_mapping = False  # Add this attribute
        self.filtered_df = None
        self.filter_keys = {
            'NUMBER': '-NUM-START-',  # Remove any reference to -NUM-END-
            'DWG': '-DWG-',
            'ORIGIN': '-ORIGIN-',
            'DEST': '-DEST-',
            'Wire Type': '-WIRE-TYPE-',
            'Length': '-LENGTH-',
            'Project ID': '-PROJECT-'
        }
    
    def validate_length_values(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
        """
        Validate Length column values and return list of validation issues
        """
        issues = []
        if 'Length' not in df.columns:
            return df, ["Length column not found in data"]
        
        # Store original values before conversion
        original_lengths = df['Length'].copy()
        
        # Convert to numeric, with NaN for invalid values
        df['Length'] = pd.to_numeric(df['Length'], errors='coerce')
        
        # Find rows with invalid length values (non-numeric)
        invalid_mask = df['Length'].isna() & original_lengths.notna()
        invalid_rows = df[invalid_mask]
        
        if not invalid_rows.empty:
            for idx, row in invalid_rows.iterrows():
                cable_id = row['NUMBER']
                invalid_value = original_lengths[idx]
                issues.append(f"Invalid Length value '{invalid_value}' for cable {cable_id}")
        
        # Handle float to int conversion
        valid_mask = df['Length'].notna()
        if valid_mask.any():
            # Check for non-integer values
            non_integer_mask = (df['Length'] % 1 != 0) & valid_mask
            non_integer_rows = df[non_integer_mask]
            
            if not non_integer_rows.empty:
                for idx, row in non_integer_rows.iterrows():
                    cable_id = row['NUMBER']
                    float_value = row['Length']
                    issues.append(f"Non-integer Length value '{float_value}' for cable {cable_id}")
            
            # Round float values to nearest integer
            df.loc[valid_mask, 'Length'] = df.loc[valid_mask, 'Length'].round()
        
        # Convert to nullable integer type
        df['Length'] = df['Length'].astype('Int64')
        
        return df, issues

    def show_validation_dialog(self, issues: List[str]) -> bool:
        """Show validation issues dialog with Yes as default"""
        issues_text = "\n".join(issues[:20])  # Show first 20 issues
        if len(issues) > 20:
            issues_text += f"\n\n...and {len(issues) - 20} more issues."
        
        layout = [
            [sg.Text("Data Loading Issues Found:")],
            [sg.Multiline(issues_text, size=(60, 20), disabled=True)],  # Increased size
            [sg.Text("\nWould you like to continue loading the data?")],
            [sg.Button('Yes', bind_return_key=True), sg.Button('No')]
        ]
        
        window = sg.Window(
            "Validation Issues",
            layout,
            modal=True,
            finalize=True
        )
        
        window['Yes'].set_focus()
        response = window.read()
        window.close()
        
        return response[0] == 'Yes'

    def load_excel_file(self, settings, show_dialog=False):
        """Load data from Excel file"""
        try:
            if show_dialog:
                file_path = sg.popup_get_file('Select Excel file', 
                                            file_types=(("Excel Files", "*.xls*"),),
                                            initial_folder=settings.settings.get('last_directory', ''))
                if not file_path:
                    print("No file selected in dialog")
                    return None, None
            else:
                file_path = settings.settings.get('default_file_path')
                if not file_path or not Path(file_path).exists():
                    print(f"Default file path invalid or not found: {file_path}")
                    return None, None

            print(f"\nAttempting to load data from {file_path}")
            
            with pd.ExcelFile(file_path) as xls:
                available_sheets = xls.sheet_names
                print(f"Available sheets: {available_sheets}")

                cable_sheet = 'CableList'
                print(f"Reading headers from sheet: {cable_sheet}")
                
                # Read all data as strings initially
                cable_list = pd.read_excel(
                    xls, 
                    sheet_name=cable_sheet,
                    dtype=str  # Force all columns to be read as strings
                )
                
                print(f"Found columns: {cable_list.columns.tolist()}")
                
                # Required and optional columns
                required_columns = ['NUMBER', 'DWG', 'ORIGIN', 'DEST']
                optional_columns = ['Alternate Dwg', 'Wire Type', 'Length', 'Note', 'Project ID']
                
                # Add missing optional columns with None values
                all_columns = required_columns + optional_columns
                for col in all_columns:
                    if col not in cable_list.columns:
                        cable_list[col] = None
                
                # Select only the columns we want
                cable_list = cable_list[all_columns]
                
                # Clean up whitespace in string columns
                string_columns = ['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 
                                'Wire Type', 'Note', 'Project ID']
                for col in string_columns:
                    if col in cable_list.columns:
                        cable_list[col] = cable_list[col].str.strip()
                
                # Validate length values
                cable_list, length_issues = self.validate_length_values(cable_list)
                
                # Show validation issues if any
                if length_issues:
                    if not self.show_validation_dialog(length_issues):
                        print("Data loading cancelled due to validation issues")
                        return None, None
                
                # Try to load length matrix if available
                length_matrix = None
                if 'LengthMatrix' in available_sheets:
                    print("Reading LengthMatrix sheet...")
                    length_matrix = pd.read_excel(xls, sheet_name='LengthMatrix', index_col=0)
                
                print(f"Successfully loaded {len(cable_list)} records")
                return cable_list, length_matrix
                
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f"Error loading file: {str(e)}")
            return None, None

    def load_data(self, file_path, settings, show_dialog=False):
        """Load and validate data"""
        try:
            self.df, self.length_matrix = self.load_excel_file(settings, show_dialog)
            if self.df is not None:
                self.validate_data()
                # Update the table immediately after loading
                print(f"Successfully loaded {len(self.df)} records")
                return True
            else:
                print("No data loaded - file loading failed")
                return False
        except Exception as e:
            print(f"Data load error: {str(e)}")
            return False
    
    def validate_data(self):
        """Ensure data meets requirements"""
        required_cols = set(UI_CONSTANTS['TABLE_COLS'])
        if not required_cols.issubset(self.df.columns):
            missing = required_cols - set(self.df.columns)
            raise ValueError(f"Missing required columns: {missing}")

    def apply_filters(self, filters: Dict[str, Any], use_exact: bool, use_fuzzy: bool) -> pd.DataFrame:
        """Apply filters to the dataframe"""
        if self.df is None:
            return pd.DataFrame()
            
        filtered_df = self.df.copy()
        
        for field, value in filters.items():
            if not value:  # Skip empty filters
                continue
                
            if field == 'NUMBER':
                # Handle number range separately
                start = value.get('start', '')
                end = value.get('end', '')
                if start:
                    filtered_df = filtered_df[filtered_df['NUMBER'] >= start]
                if end:
                    filtered_df = filtered_df[filtered_df['NUMBER'] <= end]
                continue
                
            # For all other fields
            search_value = str(value).strip()
            if not search_value:
                continue
                
            if use_exact:
                # Exact match takes precedence over fuzzy
                filtered_df = filtered_df[filtered_df[field].astype(str).str.upper() == search_value.upper()]
            elif use_fuzzy:
                # Fuzzy search using partial string matching
                filtered_df = filtered_df[
                    filtered_df[field].astype(str).str.contains(
                        search_value, 
                        case=False, 
                        na=False, 
                        regex=False
                    )
                ]
            else:
                # Default behavior: case-insensitive contains
                filtered_df = filtered_df[
                    filtered_df[field].astype(str).str.contains(
                        search_value, 
                        case=False, 
                        na=False, 
                        regex=False
                    )
                ]
        
        return filtered_df

    def _fuzzy_match(self, text: str, pattern: str, threshold: int = 80) -> bool:
        """
        Perform fuzzy matching with multiple strategies
        Returns True if any matching strategy succeeds
        """
        text = str(text).lower()
        pattern = str(pattern).lower()
        
        # Exact substring match
        if pattern in text:
            return True
            
        # Fuzzy ratio match
        if fuzz.ratio(text, pattern) >= threshold:
            return True
            
        # Partial ratio match (for substrings)
        if fuzz.partial_ratio(text, pattern) >= threshold:
            return True
            
        # Token sort ratio (handles word order differences)
        if fuzz.token_sort_ratio(text, pattern) >= threshold:
            return True
            
        # Token set ratio (handles extra/missing words)
        if fuzz.token_set_ratio(text, pattern) >= threshold:
            return True
        
        return False

    def group_by_field(self, field: str) -> Dict[str, int]:
        """Group records by field and count occurrences"""
        if self.df is None or field not in self.df.columns:
            return {}
            
        # Group by the field, handling both string and non-string data
        grouped = self.df[field].astype(str).value_counts()
        
        # Convert to dictionary and sort by value (count)
        result = dict(sorted(grouped.items(), key=lambda x: (-x[1], x[0].lower())))
        
        return result

    def update_filtered_data(self, filters: Dict[str, str], use_fuzzy: bool = False) -> Tuple[List[List], List[str]]:
        """Update filtered data based on current filters"""
        filtered_df = self.apply_filters(filters, use_fuzzy)
        
        if filtered_df.empty:
            return [], []
            
        headers = filtered_df.columns.tolist()
        values = filtered_df.values.tolist()
        
        return values, headers

    def apply_sorting(self, df, sort_column, ascending=True):
        """Apply sorting to the dataframe"""
        if sort_column:
            return df.sort_values(by=sort_column, ascending=ascending)
        return df

    def apply_grouping(self, df, group_by):
        """Apply grouping to the dataframe"""
        if group_by:
            return df.sort_values(by=group_by)
        return df

    def get_filter_values(self, values: Dict[str, Any]) -> Dict[str, str]:
        """Get filter values from UI values dictionary"""
        return {
            'NUMBER': values.get('-NUM-START-', ''),
            'DWG': values.get('-DWG-', ''),
            'ORIGIN': values.get('-ORIGIN-', ''),
            'DEST': values.get('-DEST-', ''),
            'Wire Type': values.get('-WIRE-TYPE-', ''),
            'Length': values.get('-LENGTH-', ''),
            'Project ID': values.get('-PROJECT-', '')
        }

    def group_and_sort_data(self, group_by: str = None, sort_by: str = None, 
                           sort_ascending: bool = True) -> Dict[str, Any]:
        """Group and sort the data"""
        if self.filtered_df is None:
            return {}
        
        result = {}
        
        # Handle grouping
        if group_by:
            groups = self.filtered_df.groupby(group_by).size()
            result['groups'] = dict(sorted(groups.items(), 
                                         key=lambda x: (-x[1], x[0].lower())))
        
        # Handle sorting
        if sort_by:
            self.filtered_df.sort_values(by=sort_by, 
                                       ascending=sort_ascending, 
                                       inplace=True)
        
        return result

def show_color_config_window(settings):
    """Show color configuration window"""
    
    # Load existing color settings or use defaults
    color_settings = settings.settings.get('colors', {
        'Color 1': {'color': '#FFFF00', 'keywords': []},
        'Color 2': {'color': '#E6E6FA', 'keywords': []},
        'Color 3': {'color': '#90EE90', 'keywords': []},
        'Color 4': {'color': '#ADD8E6', 'keywords': []},
        'Color 5': {'color': '#FFB6C1', 'keywords': []}
    })
    
    layout = [
        [sg.Text('Color Categories Configuration', font='Any 12 bold')],
        *[
            [
                sg.Text(f'Color {i+1}:', size=(8,1)),
                sg.Input(key=f'-COLOR{i+1}-', size=(10,1), enable_events=True),
                sg.ColorChooserButton('Pick', target=f'-COLOR{i+1}-'),
                sg.Text('Keywords:'),
                sg.Input(key=f'-KEYWORDS{i+1}-', size=(30,1))
            ] for i in range(6)
        ],
        [sg.Text('_' * 80)],
        [sg.Text('Add New Category:')],
        [sg.Input(key='-NEW-COLOR-NAME-', size=(20,1)), 
         sg.ColorChooserButton('Pick Color'),
         sg.Input(key='-NEW-KEYWORDS-', size=(30,1))],
        [sg.Button('Add Category')],
        [sg.Button('Save'), sg.Button('Cancel')]
    ]
    return sg.Window('Color Settings', layout, finalize=True, modal=True)

def create_export_options_window():
    """Create export options popup window"""
    layout = [
        [sg.Text('Export Options', font='Any 12 bold')],
        [sg.Text('Export Format:')],
        [sg.Radio('Excel', 'FORMAT', key='-EXCEL-', default=True),
         sg.Radio('CSV', 'FORMAT', key='-CSV-')],
        [sg.Text('Include:')],
        [sg.Checkbox('Headers', key='-HEADERS-', default=True),
         sg.Checkbox('Row Numbers', key='-ROW-NUMS-')],
        [sg.Text('Columns to Export:')],
        [sg.Listbox(values=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 
                           'Wire Type', 'Length', 'Note', 'Project ID'],
                   select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE,
                   size=(30, 6),
                   key='-EXPORT-COLS-')],
        [sg.Button('Export'), sg.Button('Cancel')]
    ]
    return sg.Window('Export Options', layout, finalize=True, modal=True)

def show_settings_window(settings):
    """Show settings dialog with comprehensive configuration options"""
    layout = [
        [sg.Frame('File Settings', [
            [sg.Text("Default Excel File:", size=(15, 1))],
            [sg.Input(settings.settings.get('default_file_path', ''), 
                     key='-DEFAULT-FILE-', 
                     size=(50, 1)),
             sg.FileBrowse(file_types=(("Excel Files", "*.xlsx;*.xlsm"),))],
            [sg.Checkbox("Auto-load default file on startup", 
                        key='-AUTO-LOAD-',
                        default=settings.settings.get('auto_load_default', True))]
        ])],
        
        [sg.Frame('Display Settings', [
            [sg.Text("Table Row Height:", size=(15, 1)),
             sg.Input(settings.settings.get('row_height', '25'), 
                     key='-ROW-HEIGHT-', 
                     size=(5, 1)),
             sg.Text("pixels")],
            [sg.Text("Font Size:", size=(15, 1)),
             sg.Combo(values=[8, 9, 10, 11, 12, 14], 
                     default_value=settings.settings.get('font_size', 10),
                     key='-FONT-SIZE-',
                     size=(5, 1))]
        ])],
        
        [sg.Frame('Color Theme', [
            [sg.Text("Application Theme:", size=(15, 1)),
             sg.Combo(values=['Default', 'Dark', 'Light', 'System'], 
                     default_value=settings.settings.get('theme', 'Default'),
                     key='-THEME-',
                     size=(10, 1))],
            [sg.Text("Table Colors:")],
            [sg.Text("Background:", size=(12, 1)),
             sg.Input(settings.settings.get('table_bg', '#232323'), 
                     key='-TABLE-BG-', 
                     size=(10, 1)),
             sg.ColorChooserButton('Pick')],
            [sg.Text("Alternate Row:", size=(12, 1)),
             sg.Input(settings.settings.get('table_alt', '#191919'), 
                     key='-TABLE-ALT-', 
                     size=(10, 1)),
             sg.ColorChooserButton('Pick')]
        ])],
        
        [sg.Frame('Startup Behavior', [
            [sg.Checkbox("Remember window position", 
                        key='-REMEMBER-POS-',
                        default=settings.settings.get('remember_position', True))],
            [sg.Checkbox("Remember last filters", 
                        key='-REMEMBER-FILTERS-',
                        default=settings.settings.get('remember_filters', False))],
            [sg.Checkbox("Show startup tips", 
                        key='-SHOW-TIPS-',
                        default=settings.settings.get('show_tips', True))]
        ])],
        
        [sg.Button("Save", bind_return_key=True), 
         sg.Button("Cancel"),
         sg.Button("Reset to Defaults")]
    ]
    
    window = sg.Window("Settings", layout, modal=True, finalize=True)
    window["Save"].set_focus()
    
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, "Cancel"):
            break
            
        if event == "Reset to Defaults":
            # Confirm before resetting
            if sg.popup_yes_no("Are you sure you want to reset all settings to defaults?",
                             title="Confirm Reset") == "Yes":
                settings.settings = DEFAULT_SETTINGS.copy()
                settings.save_settings()
                break
        
        if event == "Save":
            try:
                # Validate numeric inputs
                row_height = int(values['-ROW-HEIGHT-'])
                if not (10 <= row_height <= 100):
                    raise ValueError("Row height must be between 10 and 100")
                
                # Update settings
                settings.settings.update({
                    'default_file_path': values['-DEFAULT-FILE-'],
                    'auto_load_default': values['-AUTO-LOAD-'],
                    'row_height': row_height,
                    'font_size': values['-FONT-SIZE-'],
                    'theme': values['-THEME-'],
                    'table_bg': values['-TABLE-BG-'],
                    'table_alt': values['-TABLE-ALT-'],
                    'remember_position': values['-REMEMBER-POS-'],
                    'remember_filters': values['-REMEMBER-FILTERS-'],
                    'show_tips': values['-SHOW-TIPS-']
                })
                
                settings.save_settings()
                sg.popup("Settings saved successfully!", title="Success")
                break
                
            except ValueError as e:
                sg.popup_error(f"Invalid input: {str(e)}", title="Error")
            except Exception as e:
                sg.popup_error(f"Error saving settings: {str(e)}", title="Error")
    
    window.close()

def show_export_options_window():
    """Show export options window"""
    layout = [
        [sg.Text('Export Options')],
        [sg.Checkbox('Include Headers', default=True, key='-HEADERS-')],
        [sg.Checkbox('Include Row Numbers', default=False, key='-ROW_NUMS-')],
        [sg.Text('Export Format:')],
        [sg.Radio('Excel', 'FORMAT', default=True, key='-EXCEL-'),
         sg.Radio('CSV', 'FORMAT', key='-CSV-')],
        [sg.Text('Sheet Name:'), sg.Input('Sheet1', key='-SHEET_NAME-')],
        [sg.Button('Export'), sg.Button('Cancel')]
    ]
    
    window = sg.Window('Export Options', layout, modal=True, finalize=True)
    
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, 'Cancel'):
            break
            
        if event == 'Export':
            # Return export settings
            window.close()
            return values
    
    window.close()
    return None

class EventHandler:
    def __init__(self, window, data_manager):
        self.window = window
        self.data_manager = data_manager
        self.current_group_field = None
        # Update valid filter keys
        self.valid_filter_keys = [
            '-NUM-START-',
            '-DWG-',
            '-ORIGIN-',
            '-DEST-',
            '-WIRE-TYPE-',
            '-LENGTH-',
            '-PROJECT-',
            '-FUZZY-SEARCH-'
        ]

    def handle_event(self, event: str, values: Dict[str, Any]) -> bool:
        """Handle all events"""
        if event == '-FUZZY-HELP-':
            sg.popup(
                "Fuzzy Search Help\n\n"
                "When enabled, searches will match similar text:\n"
                "- Partial matches ('cat' matches 'catch')\n"
                "- Case insensitive\n"
                "- Typo-tolerant\n"
                "- More flexible matching",
                title="Fuzzy Search Help",
                background_color='#191919',
                text_color='white'
            )
            return True
        
        if event == '-APPLY-SORT-':
            sort_by = values.get('-SORT-BY-')
            if sort_by:
                self.handle_sort_event(values)
            return True
            
        if event == '-APPLY-GROUP-':
            group_by = values.get('-GROUP-BY-')
            if group_by:
                self.handle_group_event(values)
            return True
            
        if event == '-RESET-GROUP-':
            self.reset_grouping()
            return True
            
        if event == '-APPLY-FILTER-':
            self.handle_filter_event(values)
            return True
            
        if event == '-CLEAR-FILTER-':
            self.clear_filters()
            return True
            
        # Debug print for event handling
        if event != '__TIMEOUT__':
            print(f"Event received: {event}")
            
        if event == 'Filter' or '_Enter' in str(event):  # Changed event check
            return self.handle_filter(values)
        elif event == 'Sort':
            return self.handle_sort(values)
        elif event == 'Reset Sort':
            return self.handle_reset_sort()
        elif event == 'Clear Filter':
            return self.handle_clear_filter()
        elif event == 'Apply Grouping':
            return self.handle_grouping(values)
        elif event == 'Reset Grouping':
            return self.handle_reset_grouping()
        elif event in ['-GROUP-BY-', '-SORT-BY-', '-SORT-ASC-', '-SORT-DESC-']:
            self.handle_group_sort_event(event, values)
            return True
        
        if event == '-CLEAR-GROUP-':
            self.window['-GROUP-BY-'].update('')
            self.window['-GROUP-DISPLAY-'].update('')
            return True
        
        return False
    
    def handle_filter(self, values):
        """Handle filter events"""
        try:
            self.filtered_df = self.data_manager.apply_filters(values)
            self.window['-TABLE-'].update(values=self.filtered_df.values.tolist())
            return True
        except Exception as e:
            print(f"Filter error: {str(e)}")
            return False
    
    def handle_sort(self, values):
        """Handle sort events"""
        try:
            if values['-SORT-']:
                sort_col = values['-SORT-']
                ascending = values['-SORT-ASC-']
                self.filtered_df = self.filtered_df.sort_values(by=sort_col, ascending=ascending)
                self.window['-TABLE-'].update(values=self.filtered_df.values.tolist())
            return True
        except Exception as e:
            print(f"Sort error: {str(e)}")
            return False
    
    def handle_clear_filter(self):
        """Handle clear filter event"""
        try:
            self.filtered_df = self.data_manager.df.copy()
            self.window['-TABLE-'].update(values=self.filtered_df.values.tolist())
            return True
        except Exception as e:
            print(f"Clear filter error: {str(e)}")
            return False
    
    def handle_grouping(self, values):
        """Handle grouping event"""
        try:
            if values['-GROUP-']:
                # Add grouping logic here
                pass
            return True
        except Exception as e:
            print(f"Grouping error: {str(e)}")
            return False
    
    def handle_reset_grouping(self):
        """Handle reset grouping event"""
        try:
            # Add reset grouping logic here
            return True
        except Exception as e:
            print(f"Reset grouping error: {str(e)}")
            return False
    
    def handle_reset_sort(self):
        """Handle reset sort event"""
        try:
            self.filtered_df = self.filtered_df.sort_index()
            self.window['-TABLE-'].update(values=self.filtered_df.values.tolist())
            return True
        except Exception as e:
            print(f"Reset sort error: {str(e)}")
            return False

    def handle_filter_event(self, values: Dict[str, Any]) -> None:
        """Handle filter application"""
        filters = {}
        use_exact = values['-EXACT-']
        use_fuzzy = values['-FUZZY-SEARCH-']
        
        # Number range handling
        num_start = values['-NUM-START-'].strip()
        num_end = values['-NUM-END-'].strip()
        if num_start or num_end:
            filters['NUMBER'] = {
                'start': num_start,
                'end': num_end
            }
        
        # Other fields
        field_mapping = {
            'DWG': '-DWG-',
            'ORIGIN': '-ORIGIN-',
            'DEST': '-DEST-',
            'Wire Type': '-WIRE-TYPE-',
            'Length': '-LENGTH-',
            'Note': '-NOTE-',
            'Project ID': '-PROJECT-'
        }
        
        for field, key in field_mapping.items():
            if values[key].strip():
                filters[field] = values[key].strip()
        
        filtered_data = self.data_manager.apply_filters(filters, use_exact, use_fuzzy)
        
        # Convert to list of lists for table display
        table_data = filtered_data.values.tolist()
        
        self.window['-TABLE-'].update(values=table_data)
        self.window['-RECORD-COUNT-'].update(f'Records: {len(table_data)}')

    def update_group_display(self, field: str) -> None:
        """Update the group display for the selected field"""
        if not field:
            return
            
        self.current_group_field = field
        groups = self.data_manager.group_by_field(field)
        
        # Format the display text
        if groups:
            display_text = "\n".join([f"{k}: {v}" for k, v in groups.items()])
        else:
            display_text = "No groups found"
        
        self.window['-GROUP-DISPLAY-'].update(display_text)

    def clear_filters(self):
        """Clear all filters"""
        # Update list of filter keys to clear
        filter_keys = [
            '-NUM-START-',
            '-DWG-',
            '-ORIGIN-',
            '-DEST-',
            '-WIRE-TYPE-',
            '-LENGTH-',
            '-PROJECT-'
        ]
        for key in filter_keys:
            if self.window.find_element(key):
                self.window[key].update('')

    def handle_group_sort_event(self, event: str, values: Dict[str, Any]) -> None:
        """Handle group by and sort events"""
        group_by = values.get('-GROUP-BY-', None)
        sort_by = values.get('-SORT-BY-', None)
        sort_ascending = values.get('-SORT-ASC-', True)
        
        result = self.data_manager.group_and_sort_data(
            group_by, sort_by, sort_ascending
        )
        
        # Update group display
        if 'groups' in result:
            display_text = "\n".join([f"{k}: {v}" for k, v in result['groups'].items()])
            self.window['-GROUP-DISPLAY-'].update(display_text)
        
        # Update table with sorted data
        if sort_by:
            values = self.data_manager.filtered_df.values.tolist()
            self.window['-TABLE-'].update(values=values)

    def handle_menu_event(self, event: str, values: Dict[str, Any]) -> bool:
        """Handle menu-related events"""
        if '+ENTER+' in event:
            # Handle menu hover
            return True
            
        if '+LEAVE+' in event:
            # Handle menu leave
            return True
            
        if event in ('Clear Filters', 'Clear Groups', 'Default', 'Dark', 'Light'):
            # Handle menu actions
            return True
        
        return False

class UIBuilder:
    def __init__(self):
        self.constants = UI_CONSTANTS
        self.filter_keys = {
            'NUMBER': '-NUM-START-',  # Single key for number filter
            'DWG': '-DWG-',
            'ORIGIN': '-ORIGIN-',
            'DEST': '-DEST-',
            'Wire Type': '-WIRE-TYPE-',
            'Length': '-LENGTH-',
            'Project ID': '-PROJECT-'
        }

    def create_filter_input(self, key, label):
        """Standardized filter input creation"""
        return [
            sg.Text(label, size=self.constants['LABEL_SIZE']),
            sg.Input(size=self.constants['INPUT_SIZE'], 
                    key=key, 
                    enable_events=True),
            sg.Checkbox('Exact', key=f'{key}EXACT-')
        ]
    
    def create_menu(self):
        """Standardized menu creation"""
        menu_def = [
            ['&File', [
                'Open::open_key\tCtrl+O', 
                'Save\tCtrl+S', 
                'Settings\tCtrl+,', 
                '---',  # Separator
                'E&xit\tAlt+F4'
            ]],
            ['&Actions', [
                'Clear &Filters\tCtrl+F', 
                'Clear &Groups\tCtrl+G'
            ]],
            ['&Colors', [
                'Default\tCtrl+D', 
                'Dark\tCtrl+K', 
                'Light\tCtrl+L'
            ]],
            ['&Help', ['About\tF1']]
        ]
        return sg.Menu(menu_def, key='-MENU-', tearoff=False,
                      background_color='white', text_color='black',
                      font=('Arial', 10), pad=(0,0))
    
    def create_main_layout(self):
        """Create the main application layout"""
        # Menu definition - using standard sg.Menu instead of MenubarCustom
        menu_def = [
            ['&File', [
                'Open::open_key', 
                'Save', 
                'Settings', 
                '---',  # Separator
                'E&xit'
            ]],
            ['&Actions', [
                'Clear &Filters', 
                'Clear &Groups'
            ]],
            ['&Colors', [
                'Default', 
                'Dark', 
                'Light'
            ]],
            ['&Help', ['About']]
        ]

        # Filter frame - without border and label
        filter_frame = [
            [sg.Checkbox('Exact', key='-EXACT-'), 
             sg.Checkbox('Fuzzy Search', key='-FUZZY-SEARCH-')],
            [sg.Text('NUMBER:', size=(8, 1)), 
             sg.Input(size=(8, 1), key='-NUM-START-'),
             sg.Input(size=(8, 1), key='-NUM-END-')],
            [sg.Text('DWG:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-DWG-')],
            [sg.Text('ORIGIN:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-ORIGIN-')],
            [sg.Text('DEST:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-DEST-')],
            [sg.Text('Wire Type:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-WIRE-TYPE-')],
            [sg.Text('Length:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-LENGTH-')],
            [sg.Text('Note:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-NOTE-')],
            [sg.Text('Project ID:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-PROJECT-')]
        ]

        # Sort and Group controls - without border and label
        sort_group_controls = [
            [sg.Text('Sort By:', size=(8, 1)),
             sg.Combo(
                 values=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Project ID'],
                 key='-SORT-BY-',
                 size=(15, 1)
             )],
            [sg.Radio('Sort Up', 'SORT', key='-SORT-ASC-', default=True),
             sg.Radio('Sort Down', 'SORT', key='-SORT-DESC-'),
             sg.Button('Sort', key='-APPLY-SORT-', size=(8, 1))],
            [sg.Text('Group By:', size=(8, 1)),
             sg.Combo(
                 values=['DWG', 'ORIGIN', 'DEST', 'Project ID'],
                 key='-GROUP-BY-',
                 size=(15, 1)
             )],
            [sg.Button('Apply Grouping', key='-APPLY-GROUP-', size=(12, 1)),
             sg.Button('Reset Grouping', key='-RESET-GROUP-', size=(12, 1))],
            [sg.VPush()],
            [sg.Button('Filter', key='-APPLY-FILTER-', size=(8, 1)),
             sg.Button('Clear Filter', key='-CLEAR-FILTER-', size=(8, 1))]
        ]

        # Main layout with side-by-side arrangement
        layout = [
            [sg.Menu(menu_def, key='-MENU-', tearoff=False)],
            [sg.Column(filter_frame, pad=((10, 20), (10, 10))),
             sg.Column(sort_group_controls, pad=((20, 10), (10, 10)))],
            [sg.Text('Records:', key='-RECORD-COUNT-', pad=(10, 5))],
            [sg.Table(
                values=[],
                headings=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate DWG', 
                         'Wire Type', 'Length', 'Note', 'Project ID'],
                auto_size_columns=True,
                justification='left',
                key='-TABLE-',
                enable_events=True,
                expand_x=True,
                expand_y=True,
                enable_click_events=True,
                row_colors=((0, '#191919'), (1, '#212121')),
                text_color='white',
                background_color='#191919',
                pad=(10, 10)
            )]
        ]
        
        return layout

    def create_window(self):
        """Create the main window"""
        sg.theme('DarkGrey13')  # Or your preferred dark theme
        
        window = sg.Window(
            'Cable Database Interface',
            self.create_main_layout(),
            resizable=True,
            finalize=True,
            return_keyboard_events=True,
            use_default_focus=False,
            enable_close_attempted_event=True
        )
        
        # Bind keyboard shortcuts
        window.bind('<Control-o>', 'Open::open_key')
        window.bind('<Control-s>', 'Save')
        window.bind('<Control-,>', 'Settings')
        window.bind('<Alt-F4>', 'Exit')
        window.bind('<Control-f>', 'Clear Filters')
        window.bind('<Control-g>', 'Clear Groups')
        window.bind('<Control-d>', 'Default')
        window.bind('<Control-k>', 'Dark')
        window.bind('<Control-l>', 'Light')
        window.bind('<F1>', 'About')
        
        return window

class CableDatabaseApp:
    def __init__(self):
        self.settings = Settings()
        self.data_manager = DataManager()
        self.ui_builder = UIBuilder()
        self.window = None
        self.event_handler = None
        
        # Check if this is first run (no default file path set)
        if not self.settings.settings.get('default_file_path'):
            self.show_first_run_dialog()
    
    def show_first_run_dialog(self):
        """Show first-run dialog to set up initial settings"""
        layout = [
            [sg.Text("Welcome to Cable Database Interface!", font=('Any', 12, 'bold'))],
            [sg.Text("Please select your Excel file to get started:")],
            [sg.Input(key='-DEFAULT-FILE-', size=(50, 1)),
             sg.FileBrowse(file_types=(("Excel Files", "*.xlsx;*.xlsm"),))],
            [sg.Checkbox("Auto-load this file on startup", 
                        key='-AUTO-LOAD-', default=True)],
            [sg.Button("Save"), sg.Button("Skip")]
        ]
        
        window = sg.Window("First Time Setup", layout, modal=True, finalize=True)
        
        while True:
            event, values = window.read()
            
            if event in (sg.WIN_CLOSED, "Skip"):
                break
                
            if event == "Save":
                if values['-DEFAULT-FILE-']:
                    self.settings.settings.update({
                        'default_file_path': values['-DEFAULT-FILE-'],
                        'auto_load_default': values['-AUTO-LOAD-']
                    })
                    self.settings.save_settings()
                    sg.popup("Settings saved successfully!", title="Success")
                    break
                else:
                    sg.popup_error("Please select an Excel file.", title="Error")
        
        window.close()
    
    def initialize(self):
        """Initialize the application"""
        # First try to load data
        load_success = self.data_manager.load_data(
            self.settings.settings.get('default_file_path', ''),
            self.settings,
            show_dialog=not self.settings.settings.get('auto_load_default', True)
        )
        
        if not load_success:
            # If loading fails, show file selection dialog
            response = sg.popup_yes_no(
                "Failed to load data file. Would you like to select a file now?",
                title="Load Error"
            )
            if response == "Yes":
                file_path = sg.popup_get_file(
                    'Select Excel file',
                    file_types=(("Excel Files", "*.xls*"),),
                    initial_folder=self.settings.settings.get('last_directory', '')
                )
                if file_path:
                    self.settings.settings['default_file_path'] = file_path
                    self.settings.settings['last_directory'] = str(Path(file_path).parent)
                    self.settings.save_settings()
                    load_success = self.data_manager.load_data(file_path, self.settings, False)
        
        if not load_success:
            sg.popup_error("Could not load any data file. Application will close.")
            return False
            
        self.window = self.ui_builder.create_window()
        
        # Bind return key to all input fields
        input_keys = [
            '-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', '-DEST-',
            '-WIRE-TYPE-', '-LENGTH-', '-NOTE-', '-PROJECT-'
        ]
        for key in input_keys:
            if key in self.window.key_dict:  # Check if key exists before binding
                self.window[key].bind('<Return>', '_Enter')
            
        # Initialize the event handler
        self.event_handler = EventHandler(self.window, self.data_manager)
        
        # Update the table with initial data
        self.window['-TABLE-'].update(values=self.data_manager.df.values.tolist())
        
        return True
    
    def run(self):
        """Main application loop"""
        if not self.initialize():
            return
            
        while True:
            event, values = self.window.read(timeout=100)
            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            try:
                self.event_handler.handle_event(event, values)
            except Exception as e:
                print(f"Application error: {str(e)}")
                traceback.print_exc()
        
        if self.window:
            self.window.close()

if __name__ == "__main__":
    try:
        app = CableDatabaseApp()
        app.run()
    except Exception as e:
        sg.popup_error(f"Application Error: {str(e)}")
        traceback.print_exc()
    
    
























































































