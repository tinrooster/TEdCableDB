import pandas as pd
import PySimpleGUI as sg
import openpyxl
import os
import json
import traceback
from datetime import datetime
from pathlib import Path

# Constants
DEFAULT_SETTINGS = {
    'last_file_path': '',
    'default_file_path': '',
    'auto_load_default': True,
    'last_directory': '',
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

def load_data(file_path):
    print(f"\nLoading data from {file_path}")
    try:
        with pd.ExcelFile(file_path) as xls:
            required_sheets = ["CableList", "LengthMatrix"]
            missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
            
            if missing_sheets:
                sg.popup_error(f"Missing required sheets: {', '.join(missing_sheets)}")
                return None, None

            cable_list = pd.read_excel(xls, sheet_name="CableList")
            length_matrix = pd.read_excel(xls, sheet_name="LengthMatrix", index_col=0)
            return cable_list, length_matrix
            
    except Exception as e:
        print(f"Error loading data: {str(e)}")
        sg.popup_error(f"Error loading file: {str(e)}")
        return None, None

def create_menu_definition():
    """Create the menu bar definition"""
    return [
        ['File', [
            'Load Different File',
            'Save Formatted Excel',
            'Save Changes to Source',
            '---',
            'Settings',
            'Exit'
        ]],
        ['Actions', [
            'Load/Matrix Lookup',
            'Add New Record',
            'Import CSV',
            'Reload Data'
        ]],
        ['Help', [
            'About',
        ]]
    ]

def create_main_layout():
    """Create the main application layout"""
    menu_bar = sg.Menu(create_menu_definition(), key='-MENU-', pad=(0,0))
    
    # Left panel - Filters with Enter key binding
    filters_frame = sg.Frame('Filters', [
        [sg.Text('NUMBER:', size=(8,1)), 
         sg.Input(size=(8,1), key='-NUM-START-', enable_events=True), 
         sg.Text('to'), 
         sg.Input(size=(8,1), key='-NUM-END-', enable_events=True)],
        [sg.Text('DWG:', size=(8,1)), 
         sg.Input(size=(15,1), key='-DWG-', enable_events=True), 
         sg.Checkbox('Exact', key='-DWG-EXACT-')],
        [sg.Text('ORIGIN:', size=(8,1)), 
         sg.Input(size=(15,1), key='-ORIGIN-', enable_events=True), 
         sg.Checkbox('Exact', key='-ORIGIN-EXACT-')],
        [sg.Text('DEST:', size=(8,1)), 
         sg.Input(size=(15,1), key='-DEST-', enable_events=True), 
         sg.Checkbox('Exact', key='-DEST-EXACT-')],
        [sg.Text('Alternate Dwg:', size=(8,1)), 
         sg.Input(size=(15,1), key='-ALT-DWG-', enable_events=True), 
         sg.Checkbox('Exact', key='-ALT-DWG-EXACT-')],
        [sg.Text('Wire Type:', size=(8,1)), 
         sg.Input(size=(15,1), key='-WIRE-', enable_events=True), 
         sg.Checkbox('Exact', key='-WIRE-EXACT-')],
        [sg.Text('Length:', size=(8,1)), 
         sg.Input(size=(15,1), key='-LENGTH-', enable_events=True), 
         sg.Checkbox('Exact', key='-LENGTH-EXACT-')],
        [sg.Text('Note:', size=(8,1)), 
         sg.Input(size=(15,1), key='-NOTE-', enable_events=True), 
         sg.Checkbox('Exact', key='-NOTE-EXACT-')],
        [sg.Text('Project ID:', size=(8,1)), 
         sg.Input(size=(15,1), key='-PROJECT-', enable_events=True), 
         sg.Checkbox('Exact', key='-PROJECT-EXACT-')],
        [sg.Button('Apply Filter', bind_return_key=True), 
         sg.Button('Clear Filter')]
    ])
    
    # Middle panel - Color Categories
    color_frame = sg.Frame('Color Categories', [
        [sg.Text('Yellow', size=(10,1)), sg.Input(size=(8,1), key='-COLOR1-'), sg.Button('Pick', key='-PICK1-'), sg.Input(size=(20,1), key='-KEYWORDS1-')],
        [sg.Text('Lavender', size=(10,1)), sg.Input(size=(8,1), key='-COLOR2-'), sg.Button('Pick', key='-PICK2-'), sg.Input(size=(20,1), key='-KEYWORDS2-')],
        [sg.Text('Light Green', size=(10,1)), sg.Input(size=(8,1), key='-COLOR3-'), sg.Button('Pick', key='-PICK3-'), sg.Input(size=(20,1), key='-KEYWORDS3-')],
        [sg.Text('Light Blue', size=(10,1)), sg.Input(size=(8,1), key='-COLOR4-'), sg.Button('Pick', key='-PICK4-'), sg.Input(size=(20,1), key='-KEYWORDS4-')],
        [sg.Text('Light Pink', size=(10,1)), sg.Input(size=(8,1), key='-COLOR5-'), sg.Button('Pick', key='-PICK5-'), sg.Input(size=(20,1), key='-KEYWORDS5-')],
        [sg.Button('Add Category')]
    ])
    
    # Sort and Group options
    sort_group = [
        [sg.Text('Sort By:'), 
         sg.Combo(values=['NUMBER', 'DWG', 'ORIGIN', 'DEST'], size=(15,1), key='-SORT-'),
         sg.Radio('Ascending', 'SORT_DIR', key='-SORT-ASC-', default=True),
         sg.Radio('Descending', 'SORT_DIR', key='-SORT-DESC-'),
         sg.Button('Sort')],
        [sg.Text('Group By:'),
         sg.Radio('Origin', 'GROUP_BY', key='-GROUP-ORIGIN-'),
         sg.Radio('Destination', 'GROUP_BY', key='-GROUP-DEST-'),
         sg.Button('Apply Grouping'),
         sg.Button('Reset Grouping')]
    ]
    
    # Main table with fixed NUMBER column
    table_frame = [
        [sg.Table(
            values=[[]],
            headings=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 'Wire Type', 'Length', 'Note', 'Project ID'],
            display_row_numbers=True,
            justification='left',
            num_rows=25,
            alternating_row_color='lightblue',
            key='-TABLE-',
            enable_events=True,
            expand_x=True,
            expand_y=True,
            auto_size_columns=False,
            def_col_width=12,  # Default width for all columns
            col_widths=[8, 12, 25, 25, 12, 12, 8, 25, 25],  # Fixed width for NUMBER and other columns
            selected_row_colors=('white', 'blue')
        )]
    ]
    
    # Complete layout including menu
    layout = [
        [menu_bar],  # Add menu at the top
        [sg.Column([[filters_frame]], size=(300, None)), 
         sg.Column([[color_frame]], size=(400, None))],
        [sg.Column(sort_group)],
        [sg.Column(table_frame, expand_x=True, expand_y=True, pad=(0,0))]
    ]
    
    return layout

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

def load_excel_file(settings, show_dialog=False):
    """Handle Excel file loading with various scenarios"""
    try:
        if show_dialog:
            # Show file picker
            initial_folder = (Path(settings.settings['last_directory']) 
                            if settings.settings['last_directory'] 
                            else Path.home())
            
            layout = [
                [sg.Text("Select Excel File:")],
                [sg.Input(key='-FILE-', size=(50, 1)),
                 sg.FileBrowse(initial_folder=initial_folder,
                             file_types=(("Excel Files", "*.xlsx;*.xlsm"),))],
                [sg.Button("Load", bind_return_key=True), sg.Button("Cancel")]
            ]
            
            window = sg.Window("Load File", layout, modal=True, finalize=True)
            window["Load"].set_focus()
            
            event, values = window.read()
            window.close()
            
            if event == "Load" and values['-FILE-']:
                file_path = values['-FILE-']
            else:
                return None, None
                
        else:
            # Try to load default file
            file_path = settings.settings['default_file_path']
            if not file_path or not Path(file_path).exists():
                return load_excel_file(settings, show_dialog=True)
        
        # Update settings with last used directory
        settings.settings['last_directory'] = str(Path(file_path).parent)
        settings.settings['last_file_path'] = file_path
        settings.save_settings()
        
        # Load the Excel file
        return load_data(file_path)
        
    except Exception as e:
        sg.popup_error(f"Error loading file: {str(e)}")
        return None, None

def apply_filters(df, values):
    """Apply filters to the dataframe"""
    filtered_df = df.copy()
    
    # Number range filter
    if values['-NUM-START-'] or values['-NUM-END-']:
        try:
            start = int(values['-NUM-START-']) if values['-NUM-START-'] else filtered_df['NUMBER'].min()
            end = int(values['-NUM-END-']) if values['-NUM-END-'] else filtered_df['NUMBER'].max()
            filtered_df = filtered_df[filtered_df['NUMBER'].between(start, end)]
        except ValueError:
            sg.popup_error("Please enter valid numbers for the number range filter")
            return df

    # Function to apply exact or contains filter
    def apply_column_filter(df, column, value, exact):
        if not value:
            return df
        if exact:
            return df[df[column].astype(str).str.lower() == value.lower()]
        return df[df[column].astype(str).str.lower().str.contains(value.lower(), na=False)]

    # Apply filters for each column
    filter_configs = [
        ('DWG', '-DWG-', '-DWG-EXACT-'),
        ('ORIGIN', '-ORIGIN-', '-ORIGIN-EXACT-'),
        ('DEST', '-DEST-', '-DEST-EXACT-'),
        ('Alternate Dwg', '-ALT-DWG-', '-ALT-DWG-EXACT-'),
        ('Wire Type', '-WIRE-', '-WIRE-EXACT-'),
        ('Length', '-LENGTH-', '-LENGTH-EXACT-'),
        ('Note', '-NOTE-', '-NOTE-EXACT-'),
        ('Project ID', '-PROJECT-', '-PROJECT-EXACT-')
    ]

    for column, value_key, exact_key in filter_configs:
        if values[value_key]:
            filtered_df = apply_column_filter(filtered_df, column, values[value_key], values[exact_key])

    return filtered_df

def apply_sorting(df, sort_column, ascending=True):
    """Apply sorting to the dataframe"""
    if sort_column:
        return df.sort_values(by=sort_column, ascending=ascending)
    return df

def apply_grouping(df, group_by):
    """Apply grouping to the dataframe"""
    if group_by:
        return df.sort_values(by=group_by)
    return df

def main():
    print("Starting application...")
    sg.theme('SystemDefault')
    settings = Settings()
    window = None
    
    try:
        # Initial file load
        df = None
        length_matrix = None
        
        if settings.settings['auto_load_default']:
            print("Attempting to load default file...")
            df, length_matrix = load_excel_file(settings)
            if df is None:
                print("Default file load failed, showing file picker...")
                df, length_matrix = load_excel_file(settings, show_dialog=True)
        else:
            print("Auto-load disabled, showing file picker...")
            df, length_matrix = load_excel_file(settings, show_dialog=True)
        
        if df is None:
            print("No file loaded, exiting...")
            return
        
        print(f"Successfully loaded data with {len(df)} records")
        
        # Create the main window
        layout = create_main_layout()
        window = sg.Window(
            'Cable Database Interface', 
            layout,
            resizable=True,
            finalize=True,
            size=(1200, 800),
            enable_close_attempted_event=True
        )
        
        # Update table with initial data
        window['-TABLE-'].update(values=df.values.tolist())
        
        # Main event loop
        while True:
            event, values = window.read()
            print(f"Event: {event}")
            
            # Handle window closing
            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            
            try:
                # Handle Enter key in any filter input
                if isinstance(event, str) and event in ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', '-DEST-', 
                            '-ALT-DWG-', '-WIRE-', '-LENGTH-', '-NOTE-', '-PROJECT-']:
                    if values[event].endswith('\n'):  # Enter was pressed
                        filtered_df = apply_filters(df, values)
                        window['-TABLE-'].update(values=filtered_df.values.tolist())
                        window[event].update(values[event].rstrip())
                
                # Menu events
                if event == 'Load Different File':
                    df, length_matrix = load_excel_file(settings, show_dialog=True)
                    if df is not None:
                        window['-TABLE-'].update(values=df.values.tolist())
                
                elif event == 'Settings':
                    show_settings_window(settings)
                
                elif event == 'Save Formatted Excel':
                    sg.popup_notify("Save Formatted Excel - Not implemented yet")
                
                elif event == 'Save Changes to Source':
                    sg.popup_notify("Save Changes to Source - Not implemented yet")
                
                elif event == 'Load/Matrix Lookup':
                    sg.popup_notify("Matrix Lookup - Not implemented yet")
                
                elif event == 'Add New Record':
                    sg.popup_notify("Add New Record - Not implemented yet")
                
                elif event == 'Import CSV':
                    sg.popup_notify("Import CSV - Not implemented yet")
                
                elif event == 'Reload Data':
                    if settings.settings['last_file_path']:
                        df, length_matrix = load_data(settings.settings['last_file_path'])
                        if df is not None:
                            window['-TABLE-'].update(values=df.values.tolist())
                
                elif event == 'About':
                    sg.popup('KGO Cable Database Interface',
                            'Version 1.0',
                            'AC Hay',
                            title='About')
                
                # Filter handling
                elif event == 'Apply Filter' or event.endswith('\r'):
                    filtered_df = apply_filters(df, values)
                    window['-TABLE-'].update(values=filtered_df.values.tolist())
                
                elif event == 'Clear Filter':
                    for key in ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', '-DEST-', 
                               '-ALT-DWG-', '-WIRE-', '-LENGTH-', '-NOTE-', '-PROJECT-']:
                        window[key].update('')
                        if f'{key}-EXACT-' in values:
                            window[f'{key}-EXACT-'].update(False)
                    window['-TABLE-'].update(values=df.values.tolist())
                
                # Handle window resize without using WINDOW_RESIZED_EVENT
                elif isinstance(event, tuple) and len(event) == 2:  # Window resize event
                    table = window['-TABLE-']
                    new_col_widths = [8]  # Fixed width for NUMBER
                    remaining_cols = [12, 25, 25, 12, 12, 8, 25, 25]
                    table.update(col_widths=new_col_widths + remaining_cols)
            
            except Exception as e:
                print(f"Error handling event {event}: {str(e)}")
                traceback.print_exc()
                continue  # Continue running even if an event handler fails
            
    except Exception as e:
        print(f"Error in main loop: {str(e)}")
        traceback.print_exc()
        sg.popup_error(f"An error occurred: {str(e)}")
    finally:
        if window is not None:
            window.close()
        print("Application closing...")

if __name__ == "__main__":
    main()
    
    























































































