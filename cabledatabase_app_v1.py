import pandas as pd
import PySimpleGUI as sg
import openpyxl
import os
import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional

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
    'TABLE_COLS': ['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 
                  'Wire Type', 'Length', 'Note', 'Project ID'],
    'COL_WIDTHS': [8, 12, 25, 25, 12, 12, 8, 25, 25],
    'FILTER_FIELDS': [
        ('-NUM-START-', 'NUMBER'),
        ('-DWG-', 'DWG'),
        ('-ORIGIN-', 'ORIGIN'),
        # ... etc
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
        self.settings_file = Path('config/settings.json')
        self.settings = self.load_settings()
    
    def load_settings(self):
        try:
            if self.settings_file.exists():
                with open(self.settings_file, 'r') as f:
                    return {**DEFAULT_SETTINGS, **json.load(f)}
            return DEFAULT_SETTINGS
        except Exception as e:
            print(f"Error loading settings: {e}")
            return DEFAULT_SETTINGS
    
    def save_settings(self):
        try:
            self.settings_file.parent.mkdir(exist_ok=True)
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

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
                # Get available sheets
                available_sheets = xls.sheet_names
                print(f"Available sheets: {available_sheets}")

                # Find the cable list sheet (try variations of the name)
                cable_sheet = 'CableList'  # Your file already has this sheet
                
                # Read headers from selected sheet
                print(f"Reading headers from sheet: {cable_sheet}")
                excel_headers = pd.read_excel(xls, sheet_name=cable_sheet, nrows=0).columns.tolist()
                print(f"Found columns: {excel_headers}")
                
                # Required columns (make Project ID optional)
                required_columns = ['NUMBER', 'DWG', 'ORIGIN', 'DEST']
                optional_columns = ['Alternate Dwg', 'Wire Type', 'Length', 'Note', 'Project ID']
                
                # Check if we need column mapping
                need_mapping = False
                for col in required_columns:
                    if col not in excel_headers:
                        need_mapping = True
                        break
                
                if need_mapping or self.force_show_mapping:
                    column_mapping = show_column_mapping_dialog(excel_headers, required_columns + optional_columns)
                    if column_mapping:
                        save_column_mapping(column_mapping)  # Save for future use
                    else:
                        return None, None  # User cancelled mapping
                else:
                    # Use direct column names since they match
                    column_mapping = {col: col for col in required_columns + optional_columns if col in excel_headers}
                
                # Read data using the mapping
                print(f"Reading data from {cable_sheet} with column mapping...")
                # Specify numeric columns as integers
                cable_list = pd.read_excel(
                    xls, 
                    sheet_name=cable_sheet,
                    dtype={
                        'NUMBER': 'Int64',  # Using Int64 to handle potential NaN values
                        'NUMC': 'Int64',
                        'Length': 'Int64'
                    }
                )
                
                # Rename columns according to mapping
                if column_mapping:
                    cable_list = cable_list.rename(columns=column_mapping)
                
                # Add missing optional columns with None values
                all_columns = required_columns + optional_columns
                for col in all_columns:
                    if col not in cable_list.columns:
                        cable_list[col] = None
                
                # Select only the columns we want
                cable_list = cable_list[all_columns]
                
                # Ensure numeric columns are integers (in case of column renaming)
                numeric_columns = ['NUMBER', 'Length']
                for col in numeric_columns:
                    if col in cable_list.columns:
                        cable_list[col] = cable_list[col].astype('Int64')
                
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

    def apply_filters(self, values):
        """Apply filters to the dataframe"""
        print("Starting filter application...")
        filtered_df = self.df.copy()
        
        # Number range filter
        if values['-NUM-START-']:
            print(f"Applying number start filter: {values['-NUM-START-']}")
            try:
                num_start = int(values['-NUM-START-'])
                filtered_df = filtered_df[filtered_df['NUMBER'] >= num_start]
            except ValueError:
                print(f"Invalid number format for start: {values['-NUM-START-']}")
        
        if values['-NUM-END-']:
            print(f"Applying number end filter: {values['-NUM-END-']}")
            try:
                num_end = int(values['-NUM-END-'])
                filtered_df = filtered_df[filtered_df['NUMBER'] <= num_end]
            except ValueError:
                print(f"Invalid number format for end: {values['-NUM-END-']}")
        
        # Text field filters
        for field, col in [
            ('-DWG-', 'DWG'),
            ('-ORIGIN-', 'ORIGIN'),
            ('-DEST-', 'DEST'),
            ('-ALT-DWG-', 'Alternate Dwg'),
            ('-WIRE-', 'Wire Type'),
            ('-LENGTH-', 'Length'),
            ('-NOTE-', 'Note'),
            ('-PROJECT-ID-', 'Project ID')
        ]:
            if values[field]:
                print(f"Applying filter for {field}: {values[field]}")
                if values[f'{field}EXACT-']:
                    filtered_df = filtered_df[filtered_df[col].astype(str) == str(values[field])]
                else:
                    filtered_df = filtered_df[filtered_df[col].astype(str).str.contains(str(values[field]), case=False, na=False)]
        
        print(f"Filter complete. Rows remaining: {len(filtered_df)}")
        return filtered_df

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
        self.filtered_df = data_manager.df.copy()
    
    def handle_event(self, event, values):
        """Central event handling"""
        try:
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
        except Exception as e:
            print(f"Error handling event {event}: {str(e)}")
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

class UIBuilder:
    def __init__(self):
        self.constants = UI_CONSTANTS
    
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
            ['&File', ['&Load Different File', '&Save Formatted Excel', 
                      'Save Changes to Source', '---', 'Settings', 'Exit']],
            ['&Actions', ['Export Options', 'Print Preview', 'Print']],
            ['&Colors', ['Configure Colors', 'Reset Colors']],
            ['&Help', ['About', 'Documentation']]
        ]
        return sg.Menu(menu_def, key='-MENU-', tearoff=False,
                      background_color='white', text_color='black',
                      font=('Arial', 10), pad=(0,0))
    
    def create_main_layout(self):
        """Create the main application layout"""
        menu_bar = self.create_menu()
        
        headers = self.constants['TABLE_COLS']
        
        # Left panel - Filters with special handling for NUMBER range
        filters_frame = sg.Frame('Filters', [
            # Special handling for NUMBER range
            [sg.Text('NUMBER:', size=self.constants['LABEL_SIZE']),
             sg.Input(size=(8,1), key='-NUM-START-', enable_events=True),
             sg.Text('to', size=(2,1)),  # Fixed width for 'to'
             sg.Input(size=(8,1), key='-NUM-END-', enable_events=True),
             sg.Checkbox('Exact', key='-NUM-EXACT-')],
            # Regular filter inputs
            self.create_filter_input('-DWG-', 'DWG:'),
            self.create_filter_input('-ORIGIN-', 'ORIGIN:'),
            self.create_filter_input('-DEST-', 'DEST:'),
            self.create_filter_input('-ALT-DWG-', 'Alternate D:'),
            self.create_filter_input('-WIRE-', 'Wire Type:'),
            self.create_filter_input('-LENGTH-', 'Length:'),
            self.create_filter_input('-NOTE-', 'Note:'),
            self.create_filter_input('-PROJECT-ID-', 'Project ID:')
        ])

        # Sort and Group frame
        sort_group_frame = sg.Frame('Sort and Group', [
            [sg.Text('Sort By:'), 
             sg.Combo(headers, key='-SORT-', size=(15,1))],
            [sg.Radio('Ascending', 'SORT', key='-SORT-ASC-', default=True),
             sg.Radio('Descending', 'SORT', key='-SORT-DESC-')],
            [sg.Button('Sort'), sg.Button('Reset Sort')],
            [sg.Text('Group By:'),
             sg.Combo(headers, key='-GROUP-', size=(15,1))],
            [sg.Button('Apply Grouping'), sg.Button('Reset Grouping')]
        ])

        # Filter buttons frame
        filter_buttons_frame = sg.Frame('Filter Controls', [
            [sg.Button('Filter'), sg.Button('Clear Filter')]
        ])

        # Table frame with updated colors
        table_frame = [[sg.Table(
            values=[],
            headings=headers,
            auto_size_columns=True,
            col_widths=self.constants['COL_WIDTHS'],
            justification='left',
            num_rows=25,
            background_color='#232323',           # Default background
            alternating_row_color='#191919',      # Alternating row color
            text_color='#d6d6dd',                 # Text color
            selected_row_colors=('white', 'navy'), # Selected row colors
            key='-TABLE-',
            enable_events=True,
            expand_x=True,
            expand_y=True,
            bind_return_key=True
        )]]

        # Complete layout
        layout = [
            [menu_bar],
            [sg.Column([[filters_frame]], size=(300, None), pad=(0,0)), 
             sg.Column([[sort_group_frame], [filter_buttons_frame]], size=(300, None), pad=(0,0))],
            [sg.Column(table_frame, expand_x=True, expand_y=True, pad=(0,0))]
        ]
        
        return layout, headers

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
            
        layout, headers = self.ui_builder.create_main_layout()
        self.window = sg.Window(
            'Cable Database Interface', 
            layout,
            resizable=True,
            finalize=True,
            size=(1200, 800),
            location=(100, 100),
            return_keyboard_events=True,  # Enable keyboard events
            use_default_focus=False
        )
        
        # Bind return key to all input fields
        input_keys = ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', '-DEST-', 
                     '-ALT-DWG-', '-WIRE-', '-LENGTH-', '-NOTE-', '-PROJECT-ID-']
        for key in input_keys:
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
    
    
























































































