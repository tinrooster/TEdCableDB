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
        self.df = None
        self.length_matrix = None
    
    def load_excel_file(self, settings, show_dialog=False):
        """Load data from Excel file"""
        try:
            if show_dialog:
                file_path = sg.popup_get_file('Select Excel file', 
                                            file_types=(("Excel Files", "*.xls*"),),
                                            initial_folder=settings.settings['last_directory'])
                if not file_path:
                    return None, None
                
                settings.settings['last_file_path'] = file_path
                settings.settings['last_directory'] = str(Path(file_path).parent)
                settings.save_settings()
            else:
                file_path = settings.settings['default_file_path']
                if not file_path or not Path(file_path).exists():
                    return None, None

            print(f"\nLoading data from {file_path}")
            with pd.ExcelFile(file_path) as xls:
                required_sheets = ["CableList", "LengthMatrix"]
                missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
                
                if missing_sheets:
                    sg.popup_error(f"Missing required sheets: {', '.join(missing_sheets)}")
                    return None, None

                # Specify the exact columns we want
                desired_columns = ['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 
                                 'Wire Type', 'Length', 'Note', 'Project ID']
                
                # Read only the columns we want
                cable_list = pd.read_excel(xls, sheet_name="CableList", usecols=desired_columns)
                length_matrix = pd.read_excel(xls, sheet_name="LengthMatrix", index_col=0)
                
                # Ensure column order matches our headers
                cable_list = cable_list[desired_columns]
                
                return cable_list, length_matrix
                
        except Exception as e:
            print(f"Error loading data: {str(e)}")
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
    """Show settings dialog with default file configuration"""
    layout = [
        [sg.Text("Default Excel File:")],
        [sg.Input(settings.settings['default_file_path'], key='-DEFAULT-FILE-', size=(50, 1)),
         sg.FileBrowse(file_types=(("Excel Files", "*.xlsx;*.xlsm"),))],
        [sg.Checkbox("Auto-load default file on startup", 
                    key='-AUTO-LOAD-',
                    default=settings.settings['auto_load_default'])],
        [sg.Button("Save", bind_return_key=True), sg.Button("Cancel")]
    ]
    
    window = sg.Window("Settings", layout, modal=True, finalize=True)
    window["Save"].set_focus()
    
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Cancel"):
            break
        if event == "Save":
            settings.settings['default_file_path'] = values['-DEFAULT-FILE-']
            settings.settings['auto_load_default'] = values['-AUTO-LOAD-']
            settings.save_settings()
            break
    
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
    
    def initialize(self):
        """Initialize the application"""
        if not self.data_manager.load_data(
            self.settings.settings['default_file_path'],
            self.settings,
            show_dialog=not self.settings.settings['auto_load_default']
        ):
            sg.popup_error("Failed to load data. Please check the file path and try again.")
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
    
    
























































































