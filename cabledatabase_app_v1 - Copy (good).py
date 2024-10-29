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

def show_column_mapping_dialog(excel_columns: List[str], missing_columns: List[str]) -> Optional[Dict[str, str]]:
    """Show dialog for mapping Excel columns to required database fields"""
    layout = [
        [sg.Text("Column Mapping Required", font=('Any', 12, 'bold'))],
        [sg.Text("Some required columns are missing. Please map them to existing columns:")],
        [sg.Text("_" * 80)],
    ]
    
    # Create mapping inputs for each missing column
    mappings = {}
    for col in missing_columns:
        # Try to find a close match in excel_columns
        default_match = next(
            (ecol for ecol in excel_columns 
             if col.lower().replace(" ", "") in ecol.lower().replace(" ", "")),
            excel_columns[0] if excel_columns else ""
        )
        
        layout.append([
            sg.Text(f"{col}:", size=(15, 1)),
            sg.Combo(
                excel_columns,
                default_value=default_match,
                key=f'-MAP-{col}-',
                size=(30, 1),
                enable_events=True
            ),
            sg.Checkbox("Skip this column", key=f'-SKIP-{col}-', enable_events=True)
        ])
    
    layout.extend([
        [sg.Text("_" * 80)],
        [sg.Button("Apply Mapping"), sg.Button("Cancel")],
        [sg.Text("Note: Skipped columns will be created as empty", font=('Any', 9, 'italic'))]
    ])
    
    window = sg.Window("Column Mapping", layout, modal=True, finalize=True)
    
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, "Cancel"):
            window.close()
            return None
            
        # Handle checkbox events to disable/enable combos
        if event.startswith('-SKIP-'):
            col = event.replace('-SKIP-', '').replace('-', ' ')
            window[f'-MAP-{col}-'].update(disabled=values[event])
            continue
            
        if event == "Apply Mapping":
            # Create mapping dictionary
            mapping = {}
            for col in missing_columns:
                if not values[f'-SKIP-{col}-']:  # If not skipped
                    excel_col = values[f'-MAP-{col}-']
                    if excel_col:  # If a mapping was selected
                        mapping[excel_col] = col
            
            window.close()
            return mapping
    
    window.close()
    return None

class Settings:
    def __init__(self):
        """Initialize settings with proper file paths"""
        self.settings_file = Path('config/settings.json')  # Changed to config directory
        self.settings = self.load_settings()

    def create_default_settings(self) -> Dict:
        """Create default settings with proper paths"""
        return {
            'default_db_path': '',  # Empty by default
            'last_directory': os.getcwd(),
            'auto_load_default': True,
            'theme': 'DarkGrey13',
            'window_location': None,
            'window_size': None
        }

    def load_settings(self) -> Dict:
        """Load settings from file or create default"""
        try:
            # Ensure config directory exists
            self.settings_file.parent.mkdir(exist_ok=True)
            
            if self.settings_file.exists():
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    default_settings = self.create_default_settings()
                    default_settings.update(settings)
                    return default_settings
            else:
                default_settings = self.create_default_settings()
                self.save_settings(default_settings)
                return default_settings
                
        except Exception as e:
            print(f"Error loading settings: {str(e)}")
            traceback.print_exc()
            return self.create_default_settings()

    def save_settings(self, settings: Dict = None) -> None:
        """Save settings to file"""
        try:
            # Ensure config directory exists
            self.settings_file.parent.mkdir(exist_ok=True)
            
            if settings is not None:
                self.settings = settings
            
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
                
            print(f"Settings saved successfully to {self.settings_file}")
            
        except Exception as e:
            print(f"Error saving settings: {str(e)}")
            traceback.print_exc()

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
        self.df = None
        self.original_df = None
        self.filtered_df = None  # Add this to track filtered state
        self.current_filters = {}  # Track active filters
        self.expected_columns = [
            'NUMBER', 'DWG', 'ORIGIN', 'DEST', 
            'Alternate Dwg', 'Wire Type', 'Length', 
            'Note', 'ProjectID'
        ]

    def normalize_length(self, length_value: Any) -> str:
        """Normalize length values, preserving text annotations"""
        if pd.isna(length_value):
            return ''
            
        # Convert to string and clean up
        length_str = str(length_value).strip().upper()
        
        # Try to extract numeric portion
        numeric_part = ''.join(c for c in length_str if c.isdigit() or c == '.')
        text_part = ''.join(c for c in length_str if not c.isdigit() and c != '.')
        
        try:
            if numeric_part:
                # Convert to integer if possible
                num = int(float(numeric_part))
                return f"{num}{text_part}"
            return length_str
        except ValueError:
            return length_str

    def format_number(self, value) -> str:
        """Format NUMBER field to 10-digit string with leading zeros"""
        try:
            if pd.isna(value) or value == '' or value is None:
                return '0000000000'
            
            # Remove any existing leading zeros and non-numeric characters
            clean_num = ''.join(filter(str.isdigit, str(value)))
            # Pad with zeros to 10 digits
            return clean_num.zfill(10)
        except (ValueError, TypeError):
            print(f"Warning: Invalid NUMBER value: {value}")
            return '0000000000'

    def load_data(self, filename: str, settings) -> bool:
        """Load data from Excel file"""
        try:
            print(f"Loading file: {filename}")
            
            if not os.path.exists(filename):
                print(f"File not found: {filename}")
                return False

            xl = pd.ExcelFile(filename)
            print(f"Available sheets: {xl.sheet_names}")

            # Try CableList sheet first, then LengthMatrix
            if 'CableList' in xl.sheet_names:
                sheet_name = 'CableList'
            elif 'LengthMatrix' in xl.sheet_names:
                sheet_name = 'LengthMatrix'
            else:
                print("No valid sheet found in Excel file")
                return False

            print(f"Reading {sheet_name} sheet...")
            df = pd.read_excel(filename, sheet_name=sheet_name)
            
            # Ensure NUMBER is properly formatted
            if 'NUMBER' in df.columns:
                df['NUMBER'] = df['NUMBER'].apply(self.format_number)
                print("NUMBER column formatted with leading zeros")
            
            # Normalize Length values
            if 'Length' in df.columns:
                df['Length'] = df['Length'].apply(self.normalize_length)
            
            # Check for missing columns
            missing_cols = [col for col in self.expected_columns if col not in df.columns]
            if missing_cols:
                print(f"Missing columns: {missing_cols}")
                mapping = show_column_mapping_dialog(df.columns.tolist(), missing_cols)
                
                if mapping:
                    # Rename columns according to mapping
                    df = df.rename(columns=mapping)
                    
                    # Add any remaining missing columns as empty
                    remaining_missing = [col for col in missing_cols 
                                      if col not in df.columns]
                    for col in remaining_missing:
                        print(f"Adding empty column: {col}")
                        df[col] = ''
                else:
                    print("Column mapping cancelled")
                    return False

            # Keep only expected columns in the correct order
            self.df = df[self.expected_columns].copy()
            self.original_df = self.df.copy()
            
            print(f"Successfully loaded {len(self.df)} records")
            return True

        except Exception as e:
            print(f"Error loading file: {str(e)}")
            traceback.print_exc()
            return False

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

    def apply_filters(self, filters: Dict[str, Any], use_exact: bool = False, 
                     use_fuzzy: bool = False) -> pd.DataFrame:
        """Apply filters and store the filtered state"""
        if self.df is None:
            return pd.DataFrame()
            
        self.current_filters = filters  # Store current filters
        filtered_df = self.df.copy()
        
        for field, value in filters.items():
            if not value:
                continue
                
            if field == 'NUMBER':
                start = value.get('start', '').strip()
                end = value.get('end', '').strip()
                
                if start:
                    start = self.format_number(start)
                    # Use string comparison with proper zero padding
                    filtered_df = filtered_df[filtered_df['NUMBER'].astype(str).str.zfill(10) >= start]
                if end:
                    end = self.format_number(end)
                    filtered_df = filtered_df[filtered_df['NUMBER'].astype(str).str.zfill(10) <= end]
            
            elif field == 'Length':
                search_value = str(value).strip()
                if use_exact:
                    filtered_df = filtered_df[filtered_df['Length'].astype(str) == search_value]
                else:
                    filtered_df = filtered_df[filtered_df['Length'].astype(str).str.contains(search_value, case=False, na=False)]
            
            else:
                if use_exact:
                    filtered_df = filtered_df[filtered_df[field].astype(str) == str(value)]
                elif use_fuzzy:
                    filtered_df = filtered_df[filtered_df[field].astype(str).apply(
                        lambda x: fuzz.partial_ratio(str(x).lower(), str(value).lower()) > 80)]
                else:
                    filtered_df = filtered_df[filtered_df[field].astype(str).str.contains(str(value), case=False, na=False)]

        self.filtered_df = filtered_df  # Store filtered result
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

    def group_by_field(self, field: str) -> pd.DataFrame:
        """Group data by field while maintaining original order within groups"""
        if self.df is None or field not in self.df.columns:
            return self.df

        try:
            # Store original order
            self.df['_original_index'] = range(len(self.df))
            
            # Create groups but maintain order within each group
            grouped = self.df.groupby(field, sort=False) \
                            .apply(lambda x: x.sort_values('_original_index')) \
                            .reset_index(drop=True)
            
            # Remove helper column
            grouped = grouped.drop('_original_index', axis=1)
            
            # Calculate group statistics for display
            group_stats = self.df.groupby(field, sort=False).size()
            print(f"Groups created: {dict(group_stats)}")
            
            return grouped

        except Exception as e:
            print(f"Error in grouping: {str(e)}")
            traceback.print_exc()
            return self.df

    def handle_group_event(self, values: Dict[str, Any]) -> None:
        """Handle grouping event"""
        group_by = values.get('-GROUP-BY-')
        if not group_by or self.df is None:
            return

        try:
            print(f"Grouping by: {group_by}")
            self.df = self.group_by_field(group_by)
            
            # Update display
            if hasattr(self, 'window'):
                self.window['-TABLE-'].update(values=self.df.values.tolist())
                
                # Show group statistics
                group_counts = self.df.groupby(group_by, sort=False).size()
                group_display = "\n".join([f"{k}: {v}" for k, v in group_counts.items()])
                if '-GROUP-DISPLAY-' in self.window.AllKeysDict:
                    self.window['-GROUP-DISPLAY-'].update(group_display)
            
        except Exception as e:
            print(f"Error handling group event: {str(e)}")
            traceback.print_exc()

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

    def apply_grouping(self, group_by: str) -> bool:
        """Apply grouping while maintaining filtered state"""
        # Work with filtered_df if it exists, otherwise use full df
        working_df = self.filtered_df if self.filtered_df is not None else self.df
        
        if working_df is None or group_by not in working_df.columns:
            print(f"Cannot group: invalid column {group_by}")
            return False
            
        try:
            print(f"Sorting by group column: {group_by}")
            
            # Sort the working dataset
            sorted_df = working_df.sort_values(by=group_by, na_position='first')
            
            # Update the appropriate dataframe
            if self.filtered_df is not None:
                self.filtered_df = sorted_df
            else:
                self.df = sorted_df
                
            print(f"Grouped data has {len(sorted_df)} rows")
            return True
            
        except Exception as e:
            print(f"Error in grouping: {str(e)}")
            traceback.print_exc()
            return False

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

    def load_file(self, filename: str) -> bool:
        """Load and validate Excel file"""
        try:
            print(f"Attempting to load file: {filename}")
            
            # First check if file exists
            if not os.path.exists(filename):
                print(f"File not found: {filename}")
                return False

            # Read Excel file
            xl = pd.ExcelFile(filename)
            print(f"Available sheets: {xl.sheet_names}")

            # Try CableList sheet first, then LengthMatrix
            if 'CableList' in xl.sheet_names:
                sheet_name = 'CableList'
            elif 'LengthMatrix' in xl.sheet_names:
                sheet_name = 'LengthMatrix'
            else:
                print("No valid sheet found in Excel file")
                return False

            print(f"Reading from sheet: {sheet_name}")
            df = pd.read_excel(filename, sheet_name=sheet_name)
            
            # Print found columns for debugging
            print(f"Found columns: {df.columns.tolist()}")

            # Validate required columns exist
            missing_cols = [col for col in self.expected_columns if col not in df.columns]
            if missing_cols:
                print(f"Missing required columns: {missing_cols}")
                return False

            # Keep only required columns
            self.df = df[self.expected_columns].copy()
            self.original_df = self.df.copy()
            
            print(f"Successfully loaded {len(self.df)} records")
            return True

        except Exception as e:
            print(f"Error loading file: {str(e)}")
            traceback.print_exc()  # Print full stack trace
            return False

    def reset_grouping(self) -> None:
        """Reset to original order while maintaining filters"""
        if hasattr(self, 'original_df'):
            print("Resetting to original order")
            self.df = self.original_df.copy()
            # Reapply filters if they exist
            if self.current_filters:
                self.apply_filters(self.current_filters)

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

    def update_table(self):
        """Update the table display with current data"""
        try:
            display_df = self.data_manager.filtered_df if self.data_manager.filtered_df is not None else self.data_manager.df
            if display_df is not None:
                self.window['-TABLE-'].update(values=display_df.values.tolist())
                self.window['-RECORD-COUNT-'].update(f'Records: {len(display_df)}')
        except Exception as e:
            print(f"Error updating table: {str(e)}")
            traceback.print_exc()

    def handle_event(self, event: str, values: Dict[str, Any]) -> bool:
        """Handle all events"""
        try:
            # Handle input validation
            if event in ('-NUM-START-', '-NUM-END-'):
                if not self.validate_input(event, values[event]):
                    # Clear invalid input
                    self.window[event].update('')
                    return True

            # Handle table clicks
            if isinstance(event, tuple) and event[0] == '-TABLE-':
                return True

            # Handle sort events
            if event == '-APPLY-SORT-':
                sort_by = values.get('-SORT-BY-')
                if sort_by and self.data_manager.df is not None:
                    ascending = values.get('-SORT-ASC-', True)
                    working_df = self.data_manager.filtered_df if self.data_manager.filtered_df is not None else self.data_manager.df
                    sorted_df = working_df.sort_values(by=sort_by, ascending=ascending)
                    
                    if self.data_manager.filtered_df is not None:
                        self.data_manager.filtered_df = sorted_df
                    else:
                        self.data_manager.df = sorted_df
                    self.update_table()
                return True

            # Handle group events
            if event == '-APPLY-GROUP-':
                group_by = values.get('-GROUP-BY-')
                if group_by:
                    success = self.data_manager.apply_grouping(group_by)
                    if success:
                        self.update_table()
                return True

            elif event == '-RESET-GROUP-':
                self.data_manager.reset_grouping()
                self.update_table()
                return True

            # Handle filter events
            if event in ('-APPLY-FILTER-', '-CLEAR-FILTER-'):
                self.handle_filter_event(values if event == '-APPLY-FILTER-' else {})
                return True

            return True

        except Exception as e:
            print(f"Error in event handler: {str(e)}")
            traceback.print_exc()
            return True

    def handle_filter_event(self, values: Dict[str, Any]) -> None:
        """Handle filter application"""
        try:
            filters = {}
            use_exact = values.get('-EXACT-', False)
            use_fuzzy = values.get('-FUZZY-SEARCH-', False)

            # Number range handling
            num_start = values.get('-NUM-START-', '').strip()
            num_end = values.get('-NUM-END-', '').strip()
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
                'Project ID': '-PROJECT-'
            }

            for field, key in field_mapping.items():
                if values.get(key, '').strip():
                    filters[field] = values[key].strip()

            # Apply filters
            filtered_data = self.data_manager.apply_filters(filters, use_exact, use_fuzzy)
            self.window['-TABLE-'].update(values=filtered_data.values.tolist())
            self.window['-RECORD-COUNT-'].update(f'Records: {len(filtered_data)}')

        except Exception as e:
            print(f"Error in filter event: {str(e)}")
            traceback.print_exc()

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

    def validate_input(self, key: str, value: str) -> bool:
        """Validate input fields"""
        if not value:
            return True
            
        if key == '-NUM-START-' or key == '-NUM-END-':
            # Allow only digits for NUMBER fields
            if not value.isdigit():
                sg.popup_error('Please enter only digits for NUMBER field', title='Input Error')
                return False
            if len(value) > 10:
                sg.popup_error('NUMBER cannot exceed 10 digits', title='Input Error')
                return False
        
        return True

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

        # Filter frame with adjusted input sizes
        filter_frame = [
            [sg.Checkbox('Exact', key='-EXACT-'), 
             sg.Checkbox('Fuzzy Search', key='-FUZZY-SEARCH-')],
            [sg.Text('NUMBER:', size=(8, 1)), 
             sg.Input(size=(10, 1), key='-NUM-START-'),
             sg.Text('to'),
             sg.Input(size=(10, 1), key='-NUM-END-')],
            [sg.Text('DWG:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-DWG-')],
            [sg.Text('ORIGIN:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-ORIGIN-')],
            [sg.Text('DEST:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-DEST-')],
            [sg.Text('Wire Type:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-WIRE-TYPE-')],
            [sg.Text('Length:', size=(8, 1)), 
             sg.Input(size=(10, 1), key='-LENGTH-')],
            [sg.Text('Project ID:', size=(8, 1)), 
             sg.Input(size=(15, 1), key='-PROJECT-')]
        ]
        
        # Sort and Group controls with increased vertical padding
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
            [sg.VPush()],  # Add vertical spacing
            [sg.Text('Group By:', size=(8, 1)),
             sg.Combo(
                 values=['DWG', 'ORIGIN', 'DEST', 'Project ID'],
                 key='-GROUP-BY-',
                 size=(15, 1)
             )],
            [sg.Button('Apply Grouping', key='-APPLY-GROUP-', size=(12, 1)),
             sg.Button('Reset Grouping', key='-RESET-GROUP-', size=(12, 1))],
            [sg.VPush()],  # Add vertical spacing
            [sg.Button('Filter', key='-APPLY-FILTER-', size=(8, 1), pad=((0, 10), (20, 0))),
             sg.Button('Clear Filter', key='-CLEAR-FILTER-', size=(8, 1), pad=((0, 0), (20, 0)))]
        ]

        # Table with adjusted column widths
        table = sg.Table(
            values=[],
            headings=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate DWG', 
                     'Wire Type', 'Length', 'Note', 'Project ID'],
            auto_size_columns=False,
            col_widths=[10, 15, 15, 15, 15, 15, 10, 20, 15],
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
        )

        # Main layout with side-by-side arrangement
        layout = [
            [sg.Menu(menu_def, key='-MENU-', tearoff=False)],
            [sg.Column(
                [[sg.Frame('', filter_frame, relief=sg.RELIEF_FLAT)],
                 [sg.Frame('Sort and Group', sort_group_controls)]],
                pad=((10, 10), (10, 10))
            )],
            [sg.Text('Records:', key='-RECORD-COUNT-', pad=(10, 5)),
             sg.Text('', key='-STATUS-', pad=(10, 5))],  # Add status line
            [table]
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
        print("Starting initialization...")
        try:
            # Initialize basic components first
            self.settings = Settings()
            self.data_manager = DataManager()
            self.ui_builder = UIBuilder()
            sg.theme('DarkGrey13')
            
            # Create window
            print("Creating window...")
            self.window = self.ui_builder.create_window()
            
            # Create event handler after window is created
            print("Creating event handler...")
            self.event_handler = EventHandler(self.window, self.data_manager)
            
            # Try to load default file if exists
            self.load_initial_file()
            
            print("Initialization complete")
            
        except Exception as e:
            print(f"Initialization error: {str(e)}")
            traceback.print_exc()
            raise

    def load_initial_file(self):
        """Attempt to load the default file on startup"""
        try:
            default_path = self.settings.settings.get('default_db_path', '')
            if default_path and os.path.exists(default_path):
                print(f"Loading default file: {default_path}")
                if self.data_manager.load_data(default_path, self.settings):
                    print("Default file loaded successfully")
                    # Update table with loaded data
                    if self.data_manager.df is not None:
                        self.window['-TABLE-'].update(values=self.data_manager.df.values.tolist())
                        self.window['-RECORD-COUNT-'].update(f'Records: {len(self.data_manager.df)}')
                        self.window['-STATUS-'].update('Data loaded successfully')
                else:
                    print("Failed to load default file")
                    self.window['-STATUS-'].update('Failed to load default file')
            else:
                print("No default file found or file doesn't exist")
                self.window['-STATUS-'].update('No default file loaded')
        except Exception as e:
            print(f"Error loading initial file: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update('Error loading file')

    def run(self):
        """Main application loop"""
        print("Starting main application loop...")
        try:
            while True:
                event, values = self.window.read()
                print(f"Event received: {event}")  # Debug line
                
                if event in (sg.WIN_CLOSED, 'Exit', None):
                    print("Exit condition met")
                    break
                
                # Handle file-related events
                if event == 'Open::open_key':
                    filename = sg.popup_get_file(
                        'Select Excel file',
                        file_types=(("Excel Files", "*.xlsx;*.xlsm"),),
                        initial_folder=self.settings.settings.get('last_directory', '')
                    )
                    if filename:
                        if self.data_manager.load_data(filename, self.settings):
                            self.window['-TABLE-'].update(values=self.data_manager.df.values.tolist())
                            self.window['-RECORD-COUNT-'].update(f'Records: {len(self.data_manager.df)}')
                            self.window['-STATUS-'].update('File loaded successfully')
                            # Update settings
                            self.settings.settings['last_directory'] = os.path.dirname(filename)
                            self.settings.save_settings()
                
                # Handle other events through event handler
                self.event_handler.handle_event(event, values)
                
        except Exception as e:
            print(f"Error in main loop: {str(e)}")
            traceback.print_exc()
        finally:
            print("Closing window...")
            self.window.close()

if __name__ == "__main__":
    print("Application starting...")
    try:
        app = CableDatabaseApp()
        print("App instance created, starting run...")
        app.run()
        print("App run completed")
    except Exception as e:
        print(f"Critical error: {str(e)}")
        traceback.print_exc()
        sg.popup_error(f"Critical Error: {str(e)}")
    
    























































































