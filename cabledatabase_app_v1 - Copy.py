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
import pandas as pd

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
        self.filtered_df = None
        self.current_filters = {}
        self.current_file = None
        self.window = None  # Will be set later
        
        # Add column name mapping
        self.column_names = {
            'NUMBER': 'NUMBER',
            'DWG': 'DWG',
            'ORIGIN': 'ORIGIN',
            'DEST': 'DEST',
            'Alternate DWG': 'Alternate Dwg',
            'Wire Type': 'Wire Type',
            'Length': 'Length',
            'Note': 'Note',
            'Project ID': 'ProjectID'
        }
        
        self.expected_columns = list(self.column_names.values())
        self.debug_log = []  # Store recent debug messages
        self.max_log_entries = 50  # Keep last 50 messages

    def set_window(self, window):
        """Set the window reference for status updates"""
        self.window = window

    def reset_grouping(self) -> None:
        """Reset to original data order"""
        if hasattr(self, 'original_df'):
            self.df = self.original_df.copy()
            print("Reset to original order")
            self.update_status("Reset to original order")

    def load_file(self, filename: str) -> bool:
        """Load and validate Excel file"""
        try:
            self.log_debug(f"Attempting to load file: {filename}")
            self.current_file = filename
            
            if not os.path.exists(filename):
                self.update_status(f"File not found: {filename}")
                return False

            self.update_status("Opening Excel file...")
            self.update_progress(10)
            
            # Show immediate feedback before potentially long operation
            self.update_status("Reading Excel file (this may take a few moments)...")
            self.update_progress(20)
            xl = pd.ExcelFile(filename)
            
            self.update_status("Identifying target sheet...")
            self.update_progress(30)
            
            if 'CableList' in xl.sheet_names:
                sheet_name = 'CableList'
            elif 'LengthMatrix' in xl.sheet_names:
                sheet_name = 'LengthMatrix'
            else:
                self.update_status("No valid sheet found")
                return False

            # Critical long operation - show detailed status
            self.update_status(f"Reading sheet '{sheet_name}' (this may take several seconds)...")
            self.update_progress(40)
            
            # Force window refresh before long operation
            if self.window:
                self.window.refresh()
            
            df = pd.read_excel(filename, sheet_name=sheet_name)
            
            # Convert NUMBER column to integer, handling any decimal places
            if 'NUMBER' in df.columns:
                self.update_status("Converting NUMBER column to integers...")
                df['NUMBER'] = df['NUMBER'].fillna(-1)  # Handle NaN values
                df['NUMBER'] = df['NUMBER'].astype(float).astype(int)
                
            self.update_status("Processing data columns...")
            self.update_progress(70)

            # Map Project ID column name
            if 'Project ID' in df.columns:
                df.rename(columns={'Project ID': 'ProjectID'}, inplace=True)

            # Validate required columns exist
            missing_cols = [col for col in self.expected_columns if col not in df.columns]
            if missing_cols:
                self.update_status(f"Error: Missing columns {', '.join(missing_cols)}")
                return False

            self.update_status("Finalizing data load...")
            self.update_progress(90)

            # Keep only required columns
            self.df = df[self.expected_columns].copy()
            self.original_df = self.df.copy()
            
            self.update_status(f"Successfully loaded {len(self.df)} records")
            self.update_progress(100)
            return True

        except Exception as e:
            error_msg = f"Error loading file: {str(e)}"
            self.log_debug(error_msg)
            traceback.print_exc()
            return False

    def log_debug(self, message: str):
        """Add debug message to log and update status"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        debug_msg = f"[{timestamp}] {message}"
        print(debug_msg)  # Still print to console
        
        # Add to debug log
        self.debug_log.append(debug_msg)
        if len(self.debug_log) > self.max_log_entries:
            self.debug_log.pop(0)  # Remove oldest message
            
        # Update status with latest message
        self.update_status(message)

    def update_status(self, message: str):
        """Update status bar with message"""
        if self.window:
            try:
                self.window['-STATUS-TEXT-'].update(message)
                self.window.refresh()
            except Exception as e:
                print(f"Error updating status: {str(e)}")

    def update_progress(self, value: int):
        """Update progress bar"""
        if self.window:
            try:
                self.window['-PROGRESS-'].update(current_count=value, visible=True)
                self.window.refresh()
            except Exception as e:
                print(f"Error updating progress: {str(e)}")

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

    def apply_fuzzy_filter(self, df, column, search_term, threshold=75):
        """Apply fuzzy matching with improved accuracy"""
        try:
            if not search_term or df.empty:
                return df

            # Convert search term to lowercase
            search_term_lower = str(search_term).lower()
            
            # Convert column to string and handle NaN
            str_series = df[column].fillna('').astype(str)
            
            # Initialize result mask
            mask = pd.Series(False, index=df.index)
            
            # First do direct substring matching (faster)
            for idx, value in str_series.items():
                value_lower = value.lower()
                # If direct substring match, mark as match
                if search_term_lower in value_lower:
                    mask[idx] = True
                else:
                    # Only do fuzzy matching if no direct match
                    ratio = fuzz.partial_ratio(value_lower, search_term_lower)
                    mask[idx] = ratio >= threshold
            
            # Get matching results
            matches = df[mask]
            
            # Debug info
            print(f"Fuzzy search for '{search_term}' in {column}:")
            print(f"Total records: {len(df)}")
            print(f"Matches found: {len(matches)}")
            
            return matches
                
        except Exception as e:
            print(f"Error in fuzzy matching: {str(e)}")
            traceback.print_exc()
            return df

    def apply_filters(self, filters: Dict[str, Any], use_exact: bool = False, use_fuzzy: bool = False) -> pd.DataFrame:
        """Apply filters to the DataFrame"""
        try:
            if self.df is None:
                return pd.DataFrame()
            
            filtered_df = self.df.copy()
            
            for field, value in filters.items():
                if not value:  # Skip empty filters
                    continue
                
                if field == 'NUMBER':
                    try:
                        num_value = int(float(value))
                        filtered_df = filtered_df[filtered_df['NUMBER'] == num_value]
                    except ValueError:
                        print(f"Invalid number value: {value}")
                        continue
                else:
                    if use_fuzzy:
                        filtered_df = self.apply_fuzzy_filter(filtered_df, field, str(value))
                    elif use_exact:
                        # Exact matching (case-insensitive)
                        mask = filtered_df[field].astype(str).str.lower() == str(value).lower()
                        filtered_df = filtered_df[mask]
                    else:
                        # Simple contains matching (case-insensitive)
                        # Escape special regex characters and treat as literal string
                        search_value = re.escape(str(value))
                        mask = filtered_df[field].fillna('').astype(str).apply(
                            lambda x: search_value.lower() in x.lower()
                        )
                        filtered_df = filtered_df[mask]
            
            print(f"Filter applied: {field}={value}, matches: {len(filtered_df)}")
            return filtered_df
            
        except Exception as e:
            print(f"Error in filtering: {str(e)}")
            traceback.print_exc()
            return self.df.copy()

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

    def handle_sort(self, sort_by: str, ascending: bool = True) -> bool:
        """Handle sorting with proper column name mapping"""
        try:
            working_df = self.filtered_df if self.filtered_df is not None else self.df
            if working_df is None:
                self.update_status("No data to sort")
                return False

            # Map UI column name to actual DataFrame column name
            actual_column = self.column_names.get(sort_by, sort_by)
            
            if actual_column not in working_df.columns:
                error_msg = f"Column '{sort_by}' not found in data"
                self.update_status(error_msg)
                print(error_msg)
                return False

            self.update_status(f"Sorting by {sort_by}...")
            sorted_df = working_df.sort_values(by=actual_column, ascending=ascending)
            
            # Update the appropriate dataframe
            if self.filtered_df is not None:
                self.filtered_df = sorted_df
            else:
                self.df = sorted_df
                
            self.update_status(f"Sorted by {sort_by}")
            return True

        except Exception as e:
            error_msg = f"Error in sorting: {str(e)}"
            self.update_status(error_msg)
            print(error_msg)
            traceback.print_exc()
            return False

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

    window = sg.Window('Export Options', layout, finalize=True, modal=True)
    
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Cancel'):
            window.close()
            return None
        if event == 'Export':
            window.close()
            return values
            
    window.close()
    return None

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

class EventHandler:
    def __init__(self, window, data_manager):
        self.window = window
        self.data_manager = data_manager
        self.current_group_field = None

    def handle_event(self, event, values) -> bool:
        """Handle all window events"""
        try:
            print(f"Handling event: {event}")
            
            if event in (None, 'Exit', sg.WIN_CLOSED):
                return False
                
            if event == '-APPLY-FILTER-':
                self.handle_filter_event(values)
                
            elif event == '-CLEAR-FILTER-':
                self.clear_filters()
                # Reset to original data
                if self.data_manager.original_df is not None:
                    self.window['-TABLE-'].update(
                        values=self.data_manager.original_df.values.tolist()
                    )
                    self.window['-RECORD-COUNT-'].update(
                        f'Records: {len(self.data_manager.original_df)}'
                    )
                    
            elif event in ('-APPLY-SORT-', '-APPLY-GROUP-', '-RESET-GROUP-'):
                self.handle_group_sort_event(event, values)
                
            elif event.startswith('Open'):
                # Handle file open
                filename = sg.popup_get_file('Choose Excel file', 
                                           file_types=(("Excel Files", "*.xlsx"),))
                if filename:
                    if self.data_manager.load_file(filename):
                        self.update_table()
                        
            return True
            
        except Exception as e:
            print(f"Error handling event: {str(e)}")
            traceback.print_exc()
            return True

    def update_table(self):
        """Update the table with current data"""
        if self.data_manager.df is not None:
            self.window['-TABLE-'].update(values=self.data_manager.df.values.tolist())
            self.window['-RECORD-COUNT-'].update(f'Records: {len(self.data_manager.df)}')

    def handle_filter_event(self, values):
        """Handle filter application with improved search"""
        try:
            # Get filter values
            filters = {
                'NUMBER': values['-NUM-START-'],
                'DWG': values['-DWG-'],
                'ORIGIN': values['-ORIGIN-'],
                'DEST': values['-DEST-'],
                'Wire Type': values['-WIRE-TYPE-'],
                'ProjectID': values['-PROJECT-']
            }
            
            # Get search mode
            exact_match = values['-EXACT-']
            fuzzy_search = values['-FUZZY-SEARCH-']
            
            # Start with full dataset
            df = self.data_manager.df.copy() if self.data_manager.df is not None else None
            if df is None or df.empty:
                return
                
            records_before = len(df)
            
            # Apply each non-empty filter
            for field, value in filters.items():
                if not value:  # Skip empty filters
                    continue
                    
                print(f"Applying filter: {field}={value}")
                
                if field == 'NUMBER':
                    try:
                        num_value = int(float(value))
                        df = df[df['NUMBER'] == num_value]
                    except ValueError:
                        print(f"Invalid number value: {value}")
                        continue
                else:
                    if fuzzy_search:
                        df = self.apply_fuzzy_filter(df, field, str(value))
                    elif exact_match:
                        # Exact match (case-insensitive)
                        df = df[df[field].fillna('').astype(str).str.lower() == str(value).lower()]
                    else:
                        # Simple contains search (case-insensitive)
                        search_value = str(value).lower()
                        df = df[df[field].fillna('').astype(str).apply(
                            lambda x: search_value in x.lower()
                        )]
                    
                print(f"Records after {field} filter: {len(df)}")
            
            # Update the filtered dataframe
            self.data_manager.filtered_df = df
            
            # Update the table
            self.window['-TABLE-'].update(values=df.values.tolist())
            self.window['-RECORD-COUNT-'].update(f'Records: {len(df)}')
            
            # Update status
            records_filtered = records_before - len(df)
            self.window['-STATUS-TEXT-'].update(
                f'Filter applied: {records_filtered} records filtered out'
            )
            
        except Exception as e:
            print(f"Error applying filter: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f"Error applying filter: {str(e)}")

    def apply_fuzzy_filter(self, df, field, search_term, threshold=75):
        """Apply fuzzy matching to filter"""
        if not search_term:
            return df
            
        # Convert values to strings and handle NaN
        str_series = df[field].fillna('').astype(str)
        
        # First try exact substring match (case-insensitive)
        exact_matches = str_series.str.contains(search_term, case=False, na=False)
        
        # Then do fuzzy matching
        fuzzy_scores = []
        for val in str_series:
            if search_term.lower() in val.lower():
                fuzzy_scores.append(100)
            else:
                ratio = fuzz.ratio(val.lower(), search_term.lower())
                partial = fuzz.partial_ratio(val.lower(), search_term.lower())
                fuzzy_scores.append(max(ratio, partial))
        
        # Create mask for matches
        mask = (pd.Series(fuzzy_scores) >= threshold) | exact_matches
        return df[mask]

    def clear_filters(self):
        """Clear all filters"""
        try:
            # Clear all filter inputs
            filter_keys = ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', 
                         '-DEST-', '-WIRE-TYPE-', '-PROJECT-']
            for key in filter_keys:
                self.window[key].update('')
            
            # Reset checkboxes
            self.window['-EXACT-'].update(False)
            self.window['-FUZZY-SEARCH-'].update(False)
            
            # Reset to original data
            if self.data_manager.original_df is not None:
                self.data_manager.filtered_df = self.data_manager.original_df.copy()
                self.window['-TABLE-'].update(
                    values=self.data_manager.original_df.values.tolist()
                )
                self.window['-RECORD-COUNT-'].update(
                    f'Records: {len(self.data_manager.original_df)}'
                )
                self.window['-STATUS-TEXT-'].update('Filters cleared')
                
        except Exception as e:
            print(f"Error clearing filters: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f"Error clearing filters: {str(e)}")

    def handle_group_sort_event(self, event, values):
        """Handle sorting and grouping events"""
        try:
            if self.data_manager.df is None:
                sg.popup_error("No data loaded")
                return

            if event == '-APPLY-SORT-':
                sort_by = values['-SORT-BY-']
                ascending = values['-SORT-ASC-']
                
                if not sort_by:
                    sg.popup_error("Please select a column to sort by")
                    return
                    
                print(f"Sorting by {sort_by} ({'ascending' if ascending else 'descending'})")
                self.data_manager.handle_sort(sort_by, ascending)
                self.update_table()
                self.window['-STATUS-TEXT-'].update(f"Sorted by {sort_by}")
                
            elif event == '-APPLY-GROUP-':
                group_by = values['-GROUP-BY-']
                
                if not group_by:
                    sg.popup_error("Please select a column to group by")
                    return
                    
                print(f"Grouping by {group_by}")
                self.data_manager.apply_grouping(group_by)
                self.update_table()
                self.window['-STATUS-TEXT-'].update(f"Grouped by {group_by}")
                
            elif event == '-RESET-GROUP-':
                print("Resetting groups")
                # Reset to original order if filtered, otherwise use original dataframe
                if self.data_manager.filtered_df is not None:
                    self.data_manager.filtered_df = self.data_manager.filtered_df.sort_index()
                else:
                    self.data_manager.df = self.data_manager.df.sort_index()
                self.update_table()
                self.window['-STATUS-TEXT-'].update("Groups reset")
                
        except Exception as e:
            error_msg = f"Error in sorting/grouping: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            sg.popup_error(error_msg)

    def update_table(self):
        """Update the table display with current data"""
        try:
            # Use filtered_df if it exists, otherwise use main df
            df_to_display = (self.data_manager.filtered_df 
                           if self.data_manager.filtered_df is not None 
                           else self.data_manager.df)
            
            if df_to_display is not None:
                self.window['-TABLE-'].update(values=df_to_display.values.tolist())
                self.window['-RECORD-COUNT-'].update(f'Records: {len(df_to_display)}')
            
        except Exception as e:
            print(f"Error updating table: {str(e)}")
            traceback.print_exc()

class UIBuilder:
    def create_filter_frame(self):
        """Create the filter section of the UI"""
        FIELD_LENGTH = 20  # Standard length for all input fields
        
        layout = [
            [sg.Frame('Search Mode', [
                [sg.Checkbox('Exact Match', key='-EXACT-', enable_events=True),
                 sg.Checkbox('Fuzzy Search', key='-FUZZY-SEARCH-', enable_events=True)],
                [sg.Text('(Fuzzy search finds similar matches with 75% similarity)',
                        size=(40, 1), font=('Helvetica', 8, 'italic'))]
            ])],
            [sg.Text('Number Range:', size=(12, 1)),
             sg.Input(key='-NUM-START-', size=(8, 1)),
             sg.Text('to'),
             sg.Input(key='-NUM-END-', size=(8, 1))],
            [sg.Text('DWG:', size=(12, 1)),
             sg.Input(key='-DWG-', size=(FIELD_LENGTH, 1))],
            [sg.Text('Origin:', size=(12, 1)),
             sg.Input(key='-ORIGIN-', size=(FIELD_LENGTH, 1))],
            [sg.Text('Destination:', size=(12, 1)),
             sg.Input(key='-DEST-', size=(FIELD_LENGTH, 1))],
            [sg.Text('Wire Type:', size=(12, 1)),
             sg.Input(key='-WIRE-TYPE-', size=(FIELD_LENGTH, 1))],
            [sg.Text('ProjectID:', size=(12, 1)),
             sg.Input(key='-PROJECT-', size=(FIELD_LENGTH, 1))],
            [sg.Button('Apply Filter', key='-APPLY-FILTER-', bind_return_key=True),
             sg.Button('Clear Filter', key='-CLEAR-FILTER-')]
        ]
        return layout

    def create_sort_group_frame(self):
        """Create the sort and group section of the UI"""
        layout = [
            [sg.Text('Sort by:', size=(8, 1)),
             sg.Combo(['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Wire Type', 'Length'],
                     key='-SORT-BY-', size=(15, 1)),
             sg.Checkbox('Ascending', key='-SORT-ASC-', default=True)],
            [sg.Text('Group by:', size=(8, 1)),
             sg.Combo(['DWG', 'ORIGIN', 'DEST', 'ProjectID'],
                     key='-GROUP-BY-', size=(15, 1))],
            [sg.Button('Apply Sort', key='-APPLY-SORT-'),
             sg.Button('Apply Group', key='-APPLY-GROUP-'),
             sg.Button('Reset Group', key='-RESET-GROUP-')]
        ]
        return layout

    def create_main_layout(self):
        """Create the main application layout"""
        menu_def = [
            ['File', ['Open::open_key', 'Save', 'Settings', 'Exit']],
            ['Help', ['About']]
        ]
        
        table = sg.Table(
            values=[],
            headings=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 
                     'Wire Type', 'Length', 'Note', 'ProjectID'],
            auto_size_columns=False,
            col_widths=[10, 15, 60, 60, 15, 15, 10, 20, 10],
            justification='left',
            num_rows=25,
            key='-TABLE-',
            selected_row_colors='white on blue',
            enable_events=True,
            expand_x=True,
            expand_y=True,
            enable_click_events=True
        )

        status_bar = [
            [sg.Text('Records:', size=(8, 1)),
             sg.Text('0', key='-RECORD-COUNT-', size=(8, 1)),
             sg.VerticalSeparator(),
             sg.Text('Status:', size=(8, 1)),
             sg.Text('Ready', key='-STATUS-TEXT-', size=(50, 1), relief=sg.RELIEF_SUNKEN)],
            [sg.ProgressBar(100, orientation='h', size=(20, 20), 
                              key='-PROGRESS-', visible=False)]
        ]

        layout = [
            [sg.Menu(menu_def, key='-MENU-', tearoff=False)],
            [
                sg.Column([
                    [sg.Frame('Filters', self.create_filter_frame(), pad=(10, 10))],
                    [sg.Frame('Sort and Group', self.create_sort_group_frame(), pad=(10, 10))]
                ], pad=(10, 10)),
                sg.Column([[table]], expand_x=True, expand_y=True, pad=(10, 10))
            ],
            [sg.Frame('Status', status_bar, relief=sg.RELIEF_SUNKEN, pad=(10, 10))]
        ]
        
        return layout

    def create_window(self):
        """Create the main window"""
        sg.theme('DarkGrey13')
        
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
        print("Application starting...")
        print("Starting initialization...")
        self.settings = Settings()
        self.data_manager = DataManager()
        self.ui_builder = UIBuilder()
        print("Creating window...")
        self.window = self.ui_builder.create_window()
        self.event_handler = EventHandler(self.window, self.data_manager)
        print("Initialization complete")

    def update_status(self, message: str) -> None:
        """Update status bar message"""
        try:
            if hasattr(self, 'window') and self.window is not None:
                self.window['-STATUS-TEXT-'].update(message)
        except Exception as e:
            print(f"Error updating status: {str(e)}")

    def run(self):
        """Main application loop"""
        try:
            # Load initial file if configured
            self.load_initial_file()
            
            while True:
                event, values = self.window.read()
                print(f"Event received: {event}")
                
                if event in (None, 'Exit'):
                    print("Exit condition met")
                    break
                    
                # Handle events
                if not self.event_handler.handle_event(event, values):
                    break
                    
            print("Closing window...")
            self.window.close()
            print("App run completed")
            
        except Exception as e:
            print(f"Critical error in run: {str(e)}")
            traceback.print_exc()
            if hasattr(self, 'window') and self.window is not None:
                self.window.close()

    def load_initial_file(self) -> None:
        """Load initial file if configured"""
        try:
            default_file = self.settings.settings.get('default_file_path', '')
            if default_file and os.path.exists(default_file):
                if self.data_manager.load_file(default_file):
                    self.update_status('File loaded successfully')
                    self.event_handler.update_table()
                else:
                    self.update_status('Error loading default file')
            else:
                self.update_status('No default file loaded')
        except Exception as e:
            print(f"Error in load_initial_file: {str(e)}")
            self.update_status(f'Error: {str(e)}')

    def show_loading_screen(self):
        """Create a loading screen with debug output"""
        # Create debug output window with transparency
        debug_layout = [
            [sg.Multiline(size=(60, 20), key='-DEBUG-LOG-', autoscroll=True, 
                         background_color='#1a1a1a', text_color='#00ff00',
                         reroute_stdout=True, reroute_stderr=True, 
                         write_only=True, expand_x=True, expand_y=True)]
        ]
        
        # Create loading animation layout
        loading_layout = [
            [sg.Text('Loading Cable Database...', font=('Any', 16, 'bold'), 
                    text_color='white', background_color='#1a1a1a')],
            [sg.ProgressBar(100, orientation='h', size=(30, 20), key='-PROGRESS-',
                           bar_color=('green', '#1a1a1a'))],
            [sg.Text('Initializing...', key='-LOAD-STATUS-', font=('Any', 10),
                    text_color='white', background_color='#1a1a1a')]
        ]
        
        # Combined layout with both elements
        layout = [
            [sg.Column(loading_layout, background_color='#1a1a1a', 
                      element_justification='center', pad=(20, 20))],
            [sg.Column(debug_layout, background_color='#1a1a1a',
                      pad=(20, 20), expand_x=True, expand_y=True)]
        ]
        
        # Create window with transparency
        window = sg.Window('Loading', layout,
                          no_titlebar=True,
                          alpha_channel=0.9,
                          background_color='#1a1a1a',
                          finalize=True,
                          keep_on_top=True)
        
        return window

    def initialize_with_loading(self):
        """Initialize application with loading screen"""
        loading_window = self.show_loading_screen()
        
        try:
            # Update progress bar in steps
            for i in range(0, 101, 10):
                loading_window['-PROGRESS-'].update(i)
                if i == 0:
                    loading_window['-LOAD-STATUS-'].update('Loading settings...')
                    self.settings = Settings()
                    
                elif i == 20:
                    loading_window['-LOAD-STATUS-'].update('Initializing data manager...')
                    self.data_manager = DataManager()
                    
                elif i == 40:
                    loading_window['-LOAD-STATUS-'].update('Building user interface...')
                    self.ui_builder = UIBuilder()
                    
                elif i == 60:
                    loading_window['-LOAD-STATUS-'].update('Creating main window...')
                    self.window = self.ui_builder.create_window()
                    
                elif i == 80:
                    loading_window['-LOAD-STATUS-'].update('Setting up event handlers...')
                    self.event_handler = EventHandler(self.window, self.data_manager)
                    
                elif i == 100:
                    loading_window['-LOAD-STATUS-'].update('Ready!')
                    
                time.sleep(0.1)  # Add small delay for visual effect
                loading_window.refresh()
                
            # Close loading window after short delay
            time.sleep(0.5)
            loading_window.close()
            
        except Exception as e:
            print(f"Error during initialization: {str(e)}")
            loading_window.close()
            raise

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
   
 