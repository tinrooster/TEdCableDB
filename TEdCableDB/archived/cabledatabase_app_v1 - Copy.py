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
        ['Colors', [
            'Yellow Settings',
            'Lavender Settings',
            'Light Green Settings',
            'Light Blue Settings',
            'Light Pink Settings',
            '---',
            'Add New Category'
        ]],
        ['Help', [
            'About',
        ]]
    ]

def create_main_layout():
    """Create the main application layout"""
    menu_bar = sg.Menu(create_menu_definition(), key='-MENU-', pad=(0,0))
    
    # Left panel - Filters with VISIBLE buttons in their own row
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
        [sg.Text('Alternate D:', size=(8,1)), 
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
         sg.Input(size=(15,1), key='-PROJECT-ID-', enable_events=True), 
         sg.Checkbox('Exact', key='-PROJECT-ID-EXACT-')],
        [sg.Column([[  # Put buttons in their own column to ensure visibility
            sg.Button('Filter', size=(10,1), button_color=('white', 'navy')), 
            sg.Button('Clear Filter', size=(10,1), button_color=('white', 'navy'))
        ]], pad=(5,15), justification='center')]  # Add padding and center justify
    ], pad=(0,0))  # Remove padding from frame

    # Sort and Group frame with improved layout
    sort_group_frame = sg.Frame('Sort and Group', [
        [sg.Text('Sort By:', size=(8,1)), 
         sg.Combo(
            values=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 'Wire Type', 'Length', 'Note', 'Project ID'],
            size=(20,1), 
            key='-SORT-',
            enable_events=True
        )],
        [sg.Radio('Ascending', 'SORT_DIR', key='-SORT-ASC-', default=True),
         sg.Radio('Descending', 'SORT_DIR', key='-SORT-DESC-'),
         sg.Button('Sort', size=(8,1), button_color=('white', 'navy'))],
        [sg.Text('Group By:', size=(8,1))],
        [sg.Combo(
            values=['DWG', 'ORIGIN', 'DEST', 'Project ID'],
            size=(20,1),
            key='-GROUP-BY-',
            enable_events=True
        )],
        [sg.Button('Apply Grouping', size=(12,1), button_color=('white', 'navy')),
         sg.Button('Reset Grouping', size=(12,1), button_color=('white', 'navy'))]
    ], pad=(10,10))

    # Table definition with corrected right-click menu
    right_click_menu = ['&Right', [
        'Hide Column',
        'Show All Columns',
        '---',
        'Sort Ascending',
        'Sort Descending'
    ]]
    
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
            enable_click_events=True,
            expand_x=True,
            expand_y=True,
            auto_size_columns=False,
            def_col_width=12,
            col_widths=[8, 12, 25, 25, 12, 12, 8, 25, 25],
            selected_row_colors=('white', 'blue'),
            right_click_menu=right_click_menu,
            right_click_selects=True
        )]
    ]

    # Complete layout with adjusted column sizes
    layout = [
        [menu_bar],
        [sg.Column([[filters_frame]], size=(300, None), pad=(0,0)), 
         sg.Column([[sort_group_frame]], size=(300, None), pad=(0,0))],
        [sg.Column(table_frame, expand_x=True, expand_y=True, pad=(0,0))]
    ]
    
    return layout

def create_color_settings_window():
    """Create color settings popup window"""
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
    if values['-NUM-START-']:
        try:
            filtered_df = filtered_df[filtered_df['NUMBER'] >= float(values['-NUM-START-'])]
        except ValueError:
            pass
    if values['-NUM-END-']:
        try:
            filtered_df = filtered_df[filtered_df['NUMBER'] <= float(values['-NUM-END-'])]
        except ValueError:
            pass

    # Text field filters
    filter_map = {
        'DWG': ('-DWG-', '-DWG-EXACT-'),
        'ORIGIN': ('-ORIGIN-', '-ORIGIN-EXACT-'),
        'DEST': ('-DEST-', '-DEST-EXACT-'),
        'Alternate Dwg': ('-ALT-DWG-', '-ALT-DWG-EXACT-'),
        'Wire Type': ('-WIRE-', '-WIRE-EXACT-'),
        'Length': ('-LENGTH-', '-LENGTH-EXACT-'),
        'Note': ('-NOTE-', '-NOTE-EXACT-'),
        'Project ID': ('-PROJECT-ID-', '-PROJECT-ID-EXACT-')
    }

    for col, (value_key, exact_key) in filter_map.items():
        if values[value_key]:
            if values[exact_key]:  # Exact match
                filtered_df = filtered_df[filtered_df[col].astype(str).str.lower() == values[value_key].lower()]
            else:  # Contains
                filtered_df = filtered_df[filtered_df[col].astype(str).str.lower().str.contains(values[value_key].lower(), na=False)]
    
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
        [sg.Text('Color Configuration')],
        *[
            [
                sg.Text(f'Color {i+1}:'),
                sg.Input(color_settings[f'Color {i+1}']['color'], size=(10,1), key=f'-COLOR{i+1}-'),
                sg.ColorChooserButton('Pick', target=f'-COLOR{i+1}-'),
                sg.Text('Keywords:'),
                sg.Input(','.join(color_settings[f'Color {i+1}']['keywords']), 
                        size=(30,1), key=f'-KEYWORDS{i+1}-',
                        tooltip='Comma-separated keywords')
            ] for i in range(5)
        ],
        [sg.Button('Add Category')],
        [sg.Button('Save'), sg.Button('Cancel')]
    ]
    
    window = sg.Window('Color Configuration', layout, modal=True, finalize=True)
    
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, 'Cancel'):
            break
            
        if event == 'Save':
            # Save color settings
            new_settings = {}
            for i in range(5):
                color_key = f'Color {i+1}'
                new_settings[color_key] = {
                    'color': values[f'-COLOR{i+1}-'],
                    'keywords': [k.strip() for k in values[f'-KEYWORDS{i+1}-'].split(',') if k.strip()]
                }
            
            settings.settings['colors'] = new_settings
            settings.save()
            break
            
        if event == 'Add Category':
            # Add new color category
            i = len(color_settings)
            color_key = f'Color {i+1}'
            color_settings[color_key] = {'color': '#FFFFFF', 'keywords': []}
            # Refresh window with new category
            window.close()
            return show_color_config_window(settings)
    
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

def main():
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
            location=(100, 100),
            force_toplevel=True
        )
        
        # Update table with initial data
        window['-TABLE-'].update(values=df.values.tolist())
        
        while True:
            event, values = window.read()
            
            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            
            try:
                # Handle column header clicks
                if isinstance(event, tuple):
                    if event[0] == '-TABLE-' and event[2][1] is not None:  # Check if it's a column header click
                        col_num = event[2][1]
                        col_name = df.columns[col_num]
                        # Toggle sorting (ascending/descending)
                        if hasattr(main, 'last_sort_col') and main.last_sort_col == col_name:
                            main.sort_ascending = not main.sort_ascending
                        else:
                            main.sort_ascending = True
                        main.last_sort_col = col_name
                        
                        sorted_df = df.sort_values(by=col_name, ascending=main.sort_ascending)
                        window['-TABLE-'].update(values=sorted_df.values.tolist())
                        continue

                # Remove the export options window creation since it's not needed
                if event == 'Export Options':
                    continue  # Skip this event entirely

                # Handle right-click menu events
                if event == 'Hide Column' and len(values['-TABLE-']) > 0:
                    col_num = values['-TABLE-'][1]  # Get selected column number
                    visible_columns = list(window['-TABLE-'].VisibleColumns)
                    visible_columns[col_num] = False
                    window['-TABLE-'].update(visible_column_map=visible_columns)
                
                elif event == 'Show All Columns':
                    window['-TABLE-'].update(visible_column_map=[True] * len(df.columns))
                
                elif event == 'Sort Ascending' and len(values['-TABLE-']) > 0:
                    col_num = values['-TABLE-'][1]
                    col_name = df.columns[col_num]
                    sorted_df = df.sort_values(by=col_name, ascending=True)
                    window['-TABLE-'].update(values=sorted_df.values.tolist())
                
                elif event == 'Sort Descending' and len(values['-TABLE-']) > 0:
                    col_num = values['-TABLE-'][1]
                    col_name = df.columns[col_num]
                    sorted_df = df.sort_values(by=col_name, ascending=False)
                    window['-TABLE-'].update(values=sorted_df.values.tolist())
                
                # Filter handling
                elif event == 'Filter' or event.endswith('\r'):  # Handle both button and Enter key
                    filtered_df = apply_filters(df, values)
                    window['-TABLE-'].update(values=filtered_df.values.tolist())
                
                elif event == 'Clear Filter':
                    # Clear all filter inputs
                    for key in ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', '-DEST-', 
                               '-ALT-DWG-', '-WIRE-', '-LENGTH-', '-NOTE-', '-PROJECT-ID-']:
                        window[key].update('')
                        if f'{key}-EXACT-' in values:
                            window[f'{key}-EXACT-'].update(False)
                    # Reset table to show all data
                    window['-TABLE-'].update(values=df.values.tolist())
                
                # Sort handling
                elif event == 'Sort':
                    if values['-SORT-']:
                        col_idx = df.columns.get_loc(values['-SORT-'])
                        ascending = values['-SORT-ASC-']
                        sorted_df = df.sort_values(by=values['-SORT-'], ascending=ascending)
                        window['-TABLE-'].update(values=sorted_df.values.tolist())
                
                elif event == 'Apply Grouping':
                    if values['-GROUP-BY-']:
                        grouped_df = df.sort_values(by=values['-GROUP-BY-'])
                        window['-TABLE-'].update(values=grouped_df.values.tolist())
                
                # Handle popup windows
                if event == 'Color Settings':
                    color_window = create_color_settings_window()
                    while True:
                        color_event, color_values = color_window.read()
                        if color_event in (sg.WIN_CLOSED, 'Cancel'):
                            break
                        if color_event == 'Save':
                            settings.save_color_settings(color_values)
                            break
                    color_window.close()
                
                if event == 'Export Options':
                    export_window = create_export_options_window()
                    while True:
                        export_event, export_values = export_window.read()
                        if export_event in (sg.WIN_CLOSED, 'Cancel'):
                            break
                        if export_event == 'Export':
                            handle_export(df, export_values)
                            break
                    export_window.close()
                
            except Exception as e:
                print(f"Error handling event {event}: {str(e)}")
                traceback.print_exc()
                continue

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
    
    























































































