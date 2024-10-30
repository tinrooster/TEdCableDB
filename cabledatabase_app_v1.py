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
import tkinter.ttk as ttk

# Constants
DEFAULT_SETTINGS = {
    'last_file_path': '',
    'default_file_path': '',
    'auto_load_default': True,
    'last_directory': '',
    'table_config': {
        'columns': [
            'NUMBER',
            'DWG',
            'ORIGIN',
            'DEST',
            'Alternate Dwg',
            'Wire Type',
            'Length',
            'Note',
            'Project ID'
        ],
        'column_widths': {
            'NUMBER': 10,
            'DWG': 15,
            'ORIGIN': 60,
            'DEST': 60,
            'Alternate Dwg': 15,
            'Wire Type': 15,
            'Length': 10,
            'Note': 20,
            'Project ID': 10
        },
        'required_columns': [
            'NUMBER',
            'DWG',
            'ORIGIN',
            'DEST'
        ],
        'filter_keys': {
            'NUMBER': '-NUM-START-',
            'DWG': '-DWG-',
            'ORIGIN': '-ORIGIN-',
            'DEST': '-DEST-',
            'Wire Type': '-WIRE-TYPE-',
            'Length': '-LENGTH-',
            'Project ID': '-PROJECT-'
        }
    }
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
        self.settings_file = Path('config/settings.json')
        self.settings = self.load_settings()

    def create_default_settings(self) -> Dict:
        """Create default settings with proper paths"""
        return DEFAULT_SETTINGS

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

    def get_table_config(self) -> Dict:
        """Get table configuration from settings"""
        return self.settings.get('table_config', DEFAULT_SETTINGS['table_config'])

    def update_table_config(self, new_config: Dict):
        """Update table configuration"""
        self.settings['table_config'] = new_config
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
    def __init__(self, settings):
        self.settings = settings
        self.df = None
        self.original_df = None
        self.filtered_df = None
        self.current_filters = None
        self.current_group = None
        self.current_sort = None
        self.base_filtered_df = None  # Add this to store the filter-only result

    def get_current_data(self):
        """Get the current working dataset respecting filters"""
        if self.base_filtered_df is not None:
            print(f"Returning base filtered data: {len(self.base_filtered_df)} records")
            return self.base_filtered_df
        print(f"Returning original data: {len(self.df)} records")
        return self.df

    def load_file(self, file_path):
        """Load data from file"""
        try:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Attempting to load file: {file_path}")
            
            # Load Excel file
            df = pd.read_excel(file_path)
            
            # Clean up column names and data
            df = df.fillna('') # Replace NaN with empty string
            
            # Define expected columns and their order
            expected_columns = [
                'NUMBER',
                'DWG',
                'ORIGIN',
                'DEST',
                'Alternate Dwg',
                'Wire Type',
                'Length',
                'Note'
            ]
            
            # Ensure all expected columns exist
            for col in expected_columns:
                if col not in df.columns:
                    df[col] = ''  # Add missing columns with empty values
            
            # Reorder columns
            self.df = df[expected_columns]
            self.original_df = self.df.copy()
            self.filtered_df = None
            
            print(f"Successfully processed {len(self.df)} records")
            return True
            
        except Exception as e:
            print(f"Error loading file: {str(e)}")
            traceback.print_exc()
            return False

    def get_display_data(self):
        """Get current data for display"""
        df_to_display = self.filtered_df if self.filtered_df is not None else self.df
        if df_to_display is not None:
            # Convert to list and ensure empty strings instead of 'nan'
            return df_to_display.fillna('').values.tolist()
        return []

class ThemeManager:
    """Manage table colors"""
    
    @classmethod
    def apply_theme(cls, window):
        """Apply default table colors"""
        colors = {
            'even_row': '#181818',
            'odd_row': '#232323',
            'header': '#303030',
            'text': 'white',
            'selected': ('white', '#0078D7')
        }
        
        # Get table element
        table = window['-TABLE-']
        
        # Create row colors list for current data
        num_rows = len(table.Values) if table.Values else 1000
        row_colors = []
        for i in range(num_rows):
            color = colors['even_row'] if i % 2 == 0 else colors['odd_row']
            row_colors.append((i, color))
        
        # Update table with only supported parameters
        table.update(
            values=table.Values,  # Preserve current values
            row_colors=row_colors
        )

class EventHandler:
    """Handles all window events"""
    def __init__(self, window, data_manager, settings):
        self.window = window
        self.data_manager = data_manager
        self.settings = settings
        self.table_config = settings.get_table_config()
        self.bind_keyboard_shortcuts()
        self.update_status_counts()

    def bind_keyboard_shortcuts(self):
        """Bind keyboard shortcuts"""
        self.window.bind('<Control-o>', 'Open::open_key')
        self.window.bind('<Control-s>', 'Save::save_key')
        self.window.bind('<Control-comma>', 'Settings::settings_key')  # Ctrl+, for settings
        self.window.bind('<F1>', 'Help::help_key')

    def update_status_counts(self):
        """Update record and selection counts in status bar"""
        try:
            total_records = len(self.data_manager.df) if self.data_manager.df is not None else 0
            selected_rows = len(self.window['-TABLE-'].SelectedRows) if self.window['-TABLE-'].SelectedRows else 0
            
            self.window['-RECORDS-COUNT-'].update(f'{total_records:,}')
            self.window['-SELECTED-COUNT-'].update(f'{selected_rows:,}')
            
            # Update filter status if filtered
            if self.data_manager.filtered_df is not None:
                filtered_count = len(self.data_manager.filtered_df)
                if filtered_count != total_records:
                    self.window['-FILTER-STATUS-'].update(f'Filtered: {filtered_count:,} of {total_records:,}')
                else:
                    self.window['-FILTER-STATUS-'].update('')
                    
        except Exception as e:
            print(f"Error updating counts: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error: {str(e)}')

    def handle_event(self, event, values):
        """Handle window events"""
        try:
            print(f"Handling event: {event}")

            # Handle table click events properly
            if isinstance(event, tuple) and event[0] == '-TABLE-':
                if event[1] == '+CLICKED+':
                    self.update_status_counts()
                return True
            
            # Regular table selection events
            if event == '-TABLE-':
                self.update_status_counts()
                return True

            # Handle menu events
            if event in ('Settings', 'Settings::settings_key'):
                self.handle_settings_event()
                return True
            elif event == 'Open::open_key':
                self.handle_open_event(event, values)
                return True
            elif event == 'Save::save_key':
                self.handle_save_event()
                return True
            elif event == 'Help::help_key':
                self.handle_help_event()
                return True

            # Filter events
            if event == '-APPLY-FILTER-':
                self.handle_filter_event(values)
                return True
            elif event == '-CLEAR-FILTER-':
                self.handle_clear_filters()
                return True

            # Sort and Group events
            if event == '-APPLY-GROUP-':
                self.handle_group_event(values)
                return True
            elif event == '-CLEAR-GROUP-':
                self.handle_clear_group()
                return True
            elif event == '-SORT-BY-':
                self.handle_sort_event(values)
                return True

            return True  # Keep window open for unhandled events

        except Exception as e:
            print(f"Error handling event: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error: {str(e)}')
            return True  # Keep window open even if there's an error

    def handle_help_event(self):
        """Show help dialog"""
        help_text = """
        Keyboard Shortcuts:
        Ctrl+O - Open file
        Ctrl+S - Save file
        Ctrl+, - Open settings
        F1     - Show help

        Settings can be accessed through:
        - File menu -> Settings
        - Keyboard shortcut (Ctrl+,)
        """
        sg.popup_scrolled(help_text, title='Help', size=(50, 20))

    def handle_settings_event(self):
        """Handle settings dialog"""
        try:
            dialog = TableConfigurationDialog(self.settings)
            new_config = dialog.show()
            if new_config:
                self.settings.update_table_config(new_config)
                # Refresh table with new settings
                self.window['-TABLE-'].update(
                    values=self.data_manager.get_display_data(),
                    num_rows=new_config.get('rows_per_page', 25)
                )
        except Exception as e:
            print(f"Error in settings dialog: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error: {str(e)}')

    def update_table_data(self):
        """Update the table with current data"""
        try:
            if self.data_manager.filtered_df is not None:
                df = self.data_manager.filtered_df
            else:
                df = self.data_manager.df

            if df is not None:
                # Format NUMBER column as integer
                if 'NUMBER' in df.columns:
                    df['NUMBER'] = pd.to_numeric(df['NUMBER'], errors='coerce').fillna(0).astype('int64')
                
                # Convert DataFrame to list of lists for table
                data = df.values.tolist()
                self.window['-TABLE-'].update(values=data)
                self.update_status_counts()
        except Exception as e:
            print(f"Error updating table data: {str(e)}")
            traceback.print_exc()

    def handle_filter_event(self, values):
        """Handle filter application"""
        try:
            filters = {}
            
            # Number range filter
            if values['-NUM-START-'] or values['-NUM-END-']:
                try:
                    start = float(values['-NUM-START-']) if values['-NUM-START-'] else None
                    end = float(values['-NUM-END-']) if values['-NUM-END-'] else None
                    filters['NUMBER'] = (start, end)
                except ValueError:
                    sg.popup_error('Invalid number range')
                    return

            # Text field filters with fallback options
            text_fields = {
                'DWG': '-DWG-',
                'ORIGIN': '-ORIGIN-',
                'DEST': '-DEST-',
                'Wire Type': '-WIRE-TYPE-',
                'Length': '-LENGTH-',
                'Project ID': '-PROJECT-'  # Primary field name
            }
            
            # Field name variations to try
            field_variations = {
                'Project ID': ['ProjectID', 'Project_ID', 'Project'],  # Add any other variations here
            }
            
            for field, key in text_fields.items():
                if values[key]:
                    search_text = values[key].strip()
                    if search_text:
                        # Try primary field name first
                        if field in self.data_manager.df.columns:
                            filters[field] = search_text
                        else:
                            # Try variations if they exist
                            variations = field_variations.get(field, [])
                            for variant in variations:
                                if variant in self.data_manager.df.columns:
                                    filters[variant] = search_text
                                    break
                                else:
                                    print(f"Warning: Neither {field} nor its variations found in columns")

            # Apply search mode
            search_mode = 'standard'
            if values['-EXACT-']:
                search_mode = 'exact'
            elif values['-FUZZY-SEARCH-']:
                search_mode = 'fuzzy'

            # Apply filters to data
            self.apply_filters(filters, search_mode)
            
        except Exception as e:
            print(f"Error in handle_filter_event: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error applying filters: {str(e)}')

    def apply_filters(self, filters, search_mode='standard'):
        """Apply filters to the data"""
        try:
            print(f"Applying filters: {filters}")
            df = self.data_manager.df.copy()
            print(f"Initial data count: {len(df)}")
            
            for field, value in filters.items():
                if field not in df.columns:
                    print(f"Warning: Column '{field}' not found in DataFrame. Available columns: {df.columns.tolist()}")
                    continue
                    
                if field == 'NUMBER':
                    if isinstance(value, tuple):
                        start, end = value
                        print(f"Applying number range filter: {start} to {end}")
                        # Convert to integer, handling any non-numeric values
                        numeric_col = pd.to_numeric(df['NUMBER'], errors='coerce').astype('Int64')
                        
                        if start is not None:
                            try:
                                start = int(float(start))
                                df = df[numeric_col >= start]
                            except ValueError:
                                print(f"Invalid start number: {start}")
                        
                        if end is not None:
                            try:
                                end = int(float(end))
                                df = df[numeric_col <= end]
                            except ValueError:
                                print(f"Invalid end number: {end}")
                        
                        print(f"After number filter: {len(df)} records")
                else:
                    if search_mode == 'exact':
                        df = df[df[field].str.lower() == value.lower()]
                    elif search_mode == 'fuzzy':
                        df = df[df[field].str.contains(value, case=False, na=False)]
                    else:  # standard
                        df = df[df[field].str.contains(value, case=False, na=False)]
                    print(f"After {field} filter: {len(df)} records")

            self.data_manager.base_filtered_df = df.copy()
            self.data_manager.filtered_df = df.copy()
            self.data_manager.current_filters = (filters, search_mode)
            print(f"Final filtered count: {len(df)}")
            
            self.update_table_data()
            
        except Exception as e:
            print(f"Error in apply_filters: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error applying filters: {str(e)}')

    def _apply_grouping(self, group_by):
        """Internal method to apply grouping"""
        try:
            df = self.data_manager.filtered_df
            grouped = df.groupby(group_by, dropna=False)
            
            summary = []
            for name, group in grouped:
                row = {col: '' for col in df.columns}
                row[group_by] = str(name) if pd.notna(name) else '(Empty)'
                row['Count'] = len(group)
                
                for col in df.columns:
                    if col != group_by and col != 'Count':
                        first_val = group[col].iloc[0] if not group[col].empty else ''
                        row[col] = str(first_val) if pd.notna(first_val) else ''
                
                summary.append(row)
                
            self.data_manager.filtered_df = pd.DataFrame(summary)
            return True
        except Exception as e:
            print(f"Error in _apply_grouping: {str(e)}")
            return False

    def _apply_sorting(self, sort_by, ascending=True):
        """Internal method to apply sorting"""
        try:
            df = self.data_manager.filtered_df
            self.data_manager.filtered_df = df.sort_values(by=sort_by, ascending=ascending)
            return True
        except Exception as e:
            print(f"Error in _apply_sorting: {str(e)}")
            return False

    def handle_clear_filters(self):
        """Clear all filters"""
        try:
            # Clear filter inputs
            filter_keys = [
                '-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', 
                '-DEST-', '-WIRE-TYPE-', '-LENGTH-', '-PROJECT-'
            ]
            for key in filter_keys:
                self.window[key].update('')
            
            # Reset search mode to standard
            self.window['-STANDARD-SEARCH-'].update(True)
            
            # Clear filter state
            self.data_manager.current_filters = None
            
            # Reapply any active grouping or sorting
            if self.data_manager.current_group:
                self._apply_grouping(self.data_manager.current_group)
            if self.data_manager.current_sort:
                self._apply_sorting(*self.data_manager.current_sort)
            
            # Update table and status
            self.update_table_data()
            self.window['-FILTER-STATUS-'].update('')
            self.window['-STATUS-'].update('Filters cleared')
            
        except Exception as e:
            print(f"Error clearing filters: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error clearing filters: {str(e)}')

    def handle_group_event(self, values):
        """Handle grouping of data"""
        try:
            group_by = values['-GROUP-BY-']
            print(f"Handling group by: {group_by}")
            
            if not group_by or group_by == '':
                print("No group selected, clearing grouping")
                self.handle_clear_group()
                return
            
            # Use filtered data if exists
            df = self.data_manager.get_current_data()
            print(f"Data count before grouping: {len(df)}")
            
            if df is None or len(df) == 0:
                print("No data to group")
                return
            
            # Group the data
            grouped = df.groupby(group_by, dropna=False)
            print(f"Number of groups: {len(grouped)}")
            
            summary = []
            for name, group in grouped:
                row = {col: '' for col in df.columns}
                row[group_by] = str(name) if pd.notna(name) else '(Empty)'
                row['Count'] = len(group)
                
                for col in df.columns:
                    if col != group_by and col != 'Count':
                        first_val = group[col].iloc[0] if not group[col].empty else ''
                        row[col] = str(first_val) if pd.notna(first_val) else ''
                
                summary.append(row)
            
            # Store the grouped data
            summary_df = pd.DataFrame(summary)
            print(f"Summary data count: {len(summary_df)}")
            self.data_manager.filtered_df = summary_df
            self.data_manager.current_group = group_by
            
            # Update table
            self.update_table_data()
            
            # Update status
            self.window['-STATUS-'].update(f'Grouped by {group_by}')
            self.window['-FILTER-STATUS-'].update(f'{len(grouped)} groups')
            
        except Exception as e:
            print(f"Error in group operation: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error in group operation: {str(e)}')

    def handle_clear_group(self):
        """Clear grouping and restore filtered/original data"""
        try:
            print("Clearing group")
            # Clear group selection
            self.window['-GROUP-BY-'].update('')
            
            # Restore the base filtered data if it exists
            if self.data_manager.base_filtered_df is not None:
                print("Restoring base filtered data")
                self.data_manager.filtered_df = self.data_manager.base_filtered_df.copy()
            else:
                print("Restoring original data")
                self.data_manager.filtered_df = None
            
            self.data_manager.current_group = None
            
            # Update table
            self.update_table_data()
            
            # Update status
            self.window['-STATUS-'].update('Grouping cleared')
            
            # Maintain filter status if filtered
            if self.data_manager.base_filtered_df is not None:
                filtered_count = len(self.data_manager.base_filtered_df)
                total_count = len(self.data_manager.df)
                self.window['-FILTER-STATUS-'].update(
                    f'Filtered: {filtered_count:,} of {total_count:,} records'
                )
            else:
                self.window['-FILTER-STATUS-'].update('')
            
        except Exception as e:
            print(f"Error clearing group: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error clearing group: {str(e)}')

    def handle_sort_event(self, values):
        """Handle sorting of data"""
        try:
            sort_by = values['-SORT-BY-']
            if not sort_by:
                return
                
            ascending = values['-SORT-ASC-']
            
            # Use current filtered/grouped data
            df = self.data_manager.filtered_df if self.data_manager.filtered_df is not None else self.data_manager.get_current_data()
            
            # Apply sort
            self.data_manager.filtered_df = df.sort_values(by=sort_by, ascending=ascending)
            self.data_manager.current_sort = (sort_by, ascending)
            
            # Update table
            self.update_table_data()
            
            # Update status
            direction = "ascending" if ascending else "descending"
            self.window['-STATUS-'].update(f'Sorted by {sort_by} ({direction})')
            
        except Exception as e:
            print(f"Error in sort operation: {str(e)}")
            traceback.print_exc()
            self.window['-STATUS-'].update(f'Error in sort operation: {str(e)}')

class TableConfigurationDialog:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.table_config = settings.get_table_config()
        
    def create_column_config_layout(self):
        """Create layout for column configuration"""
        layout = [
            [sg.Text('Column Configuration', font=('Any', 12, 'bold'))],
            [sg.Text('_' * 80)],
            [
                sg.Column([
                    [sg.Text('Columns')],
                    [sg.Listbox(
                        values=self.table_config['columns'],
                        size=(20, 10),
                        key='-COLUMNS-LIST-',
                        enable_events=True
                    )],
                    [
                        sg.Button('Add', key='-ADD-COL-'),
                        sg.Button('Remove', key='-REMOVE-COL-'),
                        sg.Button('Move Up', key='-MOVE-UP-'),
                        sg.Button('Move Down', key='-MOVE-DOWN-')
                    ]
                ]),
                sg.Column([
                    [sg.Text('Column Properties')],
                    [sg.Text('Name:'), sg.Input(key='-COL-NAME-', size=(20, 1))],
                    [sg.Text('Width:'), sg.Input(key='-COL-WIDTH-', size=(10, 1))],
                    [sg.Checkbox('Required Column', key='-COL-REQUIRED-')],
                    [sg.Checkbox('Include in Filters', key='-COL-FILTER-')],
                    [sg.Button('Apply Changes', key='-APPLY-COL-')]
                ])
            ]
        ]
        return layout

    def create_layout(self):
        """Create the main dialog layout"""
        layout = [
            [sg.TabGroup([[
                sg.Tab('Columns', self.create_column_config_layout()),
                sg.Tab('General', [
                    [sg.Text('Default Settings', font=('Any', 12, 'bold'))],
                    [sg.Text('_' * 80)],
                    [sg.Checkbox('Auto-size columns', key='-AUTO-SIZE-',
                               default=self.table_config.get('auto_size', False))],
                    [sg.Checkbox('Remember column widths', key='-REMEMBER-WIDTHS-',
                               default=self.table_config.get('remember_widths', True))],
                    [sg.Text('Rows per page:'),
                     sg.Input(self.table_config.get('rows_per_page', 25),
                             key='-ROWS-PER-PAGE-', size=(5, 1))]
                ])
            ]])],
            [sg.Button('Save Configuration'), sg.Button('Cancel')]
        ]
        return layout

    def handle_events(self, window, event, values):
        """Handle dialog events"""
        if event == '-COLUMNS-LIST-':
            selected = values['-COLUMNS-LIST-']
            if selected:
                col_name = selected[0]
                window['-COL-NAME-'].update(col_name)
                window['-COL-WIDTH-'].update(self.table_config['column_widths'].get(col_name, 15))
                window['-COL-REQUIRED-'].update(col_name in self.table_config['required_columns'])
                window['-COL-FILTER-'].update(col_name in self.table_config['filter_keys'])
                
        elif event == '-APPLY-COL-':
            selected = values['-COLUMNS-LIST-']
            if selected:
                old_name = selected[0]
                new_name = values['-COL-NAME-']
                
                # Update column name and properties
                if old_name != new_name:
                    self.update_column_name(old_name, new_name)
                
                # Update column width
                try:
                    width = int(values['-COL-WIDTH-'])
                    self.table_config['column_widths'][new_name] = width
                except ValueError:
                    sg.popup_error('Column width must be a number')
                    return
                
                # Update required status
                if values['-COL-REQUIRED-']:
                    if new_name not in self.table_config['required_columns']:
                        self.table_config['required_columns'].append(new_name)
                else:
                    if new_name in self.table_config['required_columns']:
                        self.table_config['required_columns'].remove(new_name)
                
                # Update filter status
                if values['-COL-FILTER-']:
                    if new_name not in self.table_config['filter_keys']:
                        self.table_config['filter_keys'][new_name] = f'-{new_name.upper().replace(" ", "-")}-'
                else:
                    if new_name in self.table_config['filter_keys']:
                        del self.table_config['filter_keys'][new_name]
                
                # Update listbox
                window['-COLUMNS-LIST-'].update(self.table_config['columns'])
                
        elif event in ('-MOVE-UP-', '-MOVE-DOWN-'):
            selected = values['-COLUMNS-LIST-']
            if selected:
                idx = self.table_config['columns'].index(selected[0])
                if event == '-MOVE-UP-' and idx > 0:
                    self.table_config['columns'][idx], self.table_config['columns'][idx-1] = \
                        self.table_config['columns'][idx-1], self.table_config['columns'][idx]
                elif event == '-MOVE-DOWN-' and idx < len(self.table_config['columns']) - 1:
                    self.table_config['columns'][idx], self.table_config['columns'][idx+1] = \
                        self.table_config['columns'][idx+1], self.table_config['columns'][idx]
                window['-COLUMNS-LIST-'].update(self.table_config['columns'])

    def update_column_name(self, old_name: str, new_name: str):
        """Update column name and all related configurations"""
        # Update columns list
        idx = self.table_config['columns'].index(old_name)
        self.table_config['columns'][idx] = new_name
        
        # Update column widths
        if old_name in self.table_config['column_widths']:
            self.table_config['column_widths'][new_name] = self.table_config['column_widths'].pop(old_name)
            
        # Update required columns
        if old_name in self.table_config['required_columns']:
            self.table_config['required_columns'].remove(old_name)
            self.table_config['required_columns'].append(new_name)
            
        # Update filter keys
        if old_name in self.table_config['filter_keys']:
            self.table_config['filter_keys'][new_name] = self.table_config['filter_keys'].pop(old_name)

    def show(self):
        """Show the configuration dialog"""
        window = sg.Window('Table Configuration',
                          self.create_layout(),
                          modal=True,
                          finalize=True)
        
        while True:
            event, values = window.read()
            
            if event in (None, 'Cancel'):
                window.close()
                return None
                
            if event == 'Save Configuration':
                # Update general settings
                self.table_config.update({
                    'auto_size': values['-AUTO-SIZE-'],
                    'remember_widths': values['-REMEMBER-WIDTHS-'],
                    'rows_per_page': int(values['-ROWS-PER-PAGE-'])
                })
                
                window.close()
                return self.table_config
                
            self.handle_events(window, event, values)

class UIBuilder:
    def __init__(self, settings):
        self.settings = settings
        self.table_config = settings.get_table_config()
        # Define menu once
        self.menu_def = [
            ['&File', [
                '&Open::open_key',
                '&Save::save_key',
                '---',
                'Se&ttings::settings_key',
                '---',
                'E&xit'
            ]],
            ['&Help', [
                '&Quick Guide',
                '&Keyboard Shortcuts',
                '---',
                '&About'
            ]]
        ]

    def create_window(self):
        """Create the main application window"""
        layout = self.create_main_layout()
        return sg.Window('Cable Database Interface',
                        layout,
                        resizable=True,
                        finalize=True,
                        size=(800, 600))  # Set initial size only

    def create_filter_frame(self):
        """Create filter section"""
        filter_layout = [
            # Search Options
            [
                sg.Frame('Search Options', [
                    [
                        sg.Radio('Standard Search', 'SEARCH_MODE', default=True, key='-STANDARD-SEARCH-'),
                        sg.Radio('Exact Match', 'SEARCH_MODE', key='-EXACT-'),
                        sg.Radio('Fuzzy Search', 'SEARCH_MODE', key='-FUZZY-SEARCH-')
                    ]
                ])
            ],
            
            # Filter Fields - organized in columns
            [
                sg.Column([
                    [   # Fixed: NUMBER row properly nested
                        sg.Text('NUMBER:', size=(8, 1)), 
                        sg.Input(key='-NUM-START-', size=(10, 1)),
                        sg.Text('to'),
                        sg.Input(key='-NUM-END-', size=(10, 1))
                    ],
                    [   # DWG row
                        sg.Text('DWG:', size=(8, 1)), 
                        sg.Input(key='-DWG-', size=(25, 1))
                    ]
                ]),
                sg.Column([
                    [sg.Text('ORIGIN:', size=(8, 1)), 
                     sg.Input(key='-ORIGIN-', size=(25, 1))],
                    [sg.Text('DEST:', size=(8, 1)), 
                     sg.Input(key='-DEST-', size=(25, 1))]
                ])
            ],
            
            # Additional Filters
            [
                sg.Column([
                    [sg.Text('Wire Type:', size=(8, 1)), 
                     sg.Input(key='-WIRE-TYPE-', size=(15, 1))],
                    [sg.Text('Length:', size=(8, 1)), 
                     sg.Input(key='-LENGTH-', size=(15, 1))]
                ]),
                sg.Column([
                    [sg.Text('Project:', size=(8, 1)), 
                     sg.Input(key='-PROJECT-', size=(15, 1))]
                ])
            ],
            
            # Filter Actions
            [
                sg.Button('Apply Filters', key='-APPLY-FILTER-', bind_return_key=True),
                sg.Button('Clear Filters', key='-CLEAR-FILTER-'),
                sg.Push(),
                sg.Text('', key='-FILTER-STATUS-', size=(30, 1), text_color='yellow')
            ]
        ]
        return filter_layout

    def create_sort_group_frame(self):
        """Create sort and group controls"""
        return [
            [
                sg.Column([
                    [
                        sg.Text('Sort by:', size=(8, 1)),
                        sg.Combo(self.table_config['columns'], key='-SORT-BY-', size=(15, 1)),
                        sg.Radio('Ascending', 'SORT_DIR', default=True, key='-SORT-ASC-'),
                        sg.Radio('Descending', 'SORT_DIR', key='-SORT-DESC-')
                    ]
                ]),
                sg.VerticalSeparator(),
                sg.Column([
                    [
                        sg.Text('Group by:', size=(8, 1)),
                        sg.Combo(self.table_config['columns'], key='-GROUP-BY-', size=(15, 1)),
                        sg.Button('Apply', key='-APPLY-GROUP-'),
                        sg.Button('Clear', key='-CLEAR-GROUP-')
                    ]
                ])
            ]
        ]

    def create_main_layout(self):
        """Create the main application layout"""
        # Define valid columns for grouping
        groupable_columns = [''] + [
            'NUMBER',
            'DWG',
            'ORIGIN',
            'DEST',
            'Wire Type',
            'Length'
        ]

        layout = [
            # Menu
            [sg.Menu(self.menu_def, key='-MENU-', tearoff=False)],
            
            # Controls Row
            [
                # Left side - Filters
                sg.Frame('Filters', [
                    [sg.Frame('Search Options', [
                        [
                            sg.Radio('Standard Search', 'SEARCH_MODE', default=True, key='-STANDARD-SEARCH-'),
                            sg.Radio('Exact Match', 'SEARCH_MODE', key='-EXACT-'),
                            sg.Radio('Fuzzy Search', 'SEARCH_MODE', key='-FUZZY-SEARCH-')
                        ]
                    ])],
                    [
                        sg.Column([
                            [sg.Text('NUMBER:', size=(8, 1)), 
                             sg.Input(key='-NUM-START-', size=(10, 1)),
                             sg.Text('to'),
                             sg.Input(key='-NUM-END-', size=(10, 1))],
                            [sg.Text('DWG:', size=(8, 1)), 
                             sg.Input(key='-DWG-', size=(25, 1))],
                            [sg.Text('ORIGIN:', size=(8, 1)), 
                             sg.Input(key='-ORIGIN-', size=(25, 1))],
                            [sg.Text('DEST:', size=(8, 1)), 
                             sg.Input(key='-DEST-', size=(25, 1))]
                        ]),
                        sg.Column([
                            [sg.Text('Wire Type:', size=(8, 1)), 
                             sg.Input(key='-WIRE-TYPE-', size=(15, 1))],
                            [sg.Text('Length:', size=(8, 1)), 
                             sg.Input(key='-LENGTH-', size=(15, 1))],
                            [sg.Text('Project:', size=(8, 1)), 
                             sg.Input(key='-PROJECT-', size=(15, 1))]
                        ])
                    ],
                    [
                        sg.Button('Apply Filters', key='-APPLY-FILTER-', bind_return_key=True),
                        sg.Button('Clear Filters', key='-CLEAR-FILTER-')
                    ]
                ]),
                
                # Right side - Sort and Group
                sg.Frame('Sort and Group', [
                    [
                        sg.Text('Sort by:', size=(8, 1)),
                        sg.Combo(self.table_config['columns'], key='-SORT-BY-', size=(15, 1)),
                        sg.Radio('Ascending', 'SORT_DIR', default=True, key='-SORT-ASC-'),
                        sg.Radio('Descending', 'SORT_DIR', key='-SORT-DESC-')
                    ],
                    [
                        sg.Text('Group by:', size=(8, 1)),
                        sg.Combo(groupable_columns, key='-GROUP-BY-', size=(15, 1)),
                        sg.Button('Apply', key='-APPLY-GROUP-'),
                        sg.Button('Clear', key='-CLEAR-GROUP-')
                    ]
                ])
            ],
            
            # Table
            [sg.Table(
                values=[],
                headings=self.table_config['columns'],
                auto_size_columns=False,
                col_widths=[self.table_config['column_widths'][col] for col in self.table_config['columns']],
                justification='left',
                num_rows=25,
                key='-TABLE-',
                enable_events=True,
                expand_x=True,
                expand_y=True,
                vertical_scroll_only=False,
                enable_click_events=True,
                right_click_menu=['&Right', ['Copy', 'Export Selection', '---', 'Settings']],
                selected_row_colors=('white', 'blue'),
                header_background_color=sg.theme_input_background_color(),
                row_height=25
            )],
            
            # Status Bar
            [sg.HorizontalSeparator()],
            [
                sg.Text('Ready', key='-STATUS-', size=(30, 1)),
                sg.Push(),
                sg.Text('', key='-FILTER-STATUS-', size=(30, 1), text_color='yellow'),
                sg.VerticalSeparator(),
                sg.Text('Records:', pad=(5, 0)),
                sg.Text('0', size=(8, 1), key='-RECORDS-COUNT-', justification='right'),
                sg.VerticalSeparator(),
                sg.Text('Selected:', pad=(5, 0)),
                sg.Text('0', size=(8, 1), key='-SELECTED-COUNT-', justification='right')
            ]
        ]
        return layout

class CableDatabaseApp:
    def __init__(self):
        print("Application starting...")
        self.settings = Settings()
        self.data_manager = DataManager(self.settings)
        self.ui_builder = UIBuilder(self.settings)
        self.window = self.ui_builder.create_window()
        self.event_handler = EventHandler(self.window, self.data_manager, self.settings)
        # Note: Don't load file here

    def update_status(self, message):
        """Update status bar message"""
        try:
            if self.window and not self.window.was_closed():
                self.window['-STATUS-'].update(message)
                print(f"Status: {message}")
        except Exception as e:
            print(f"Error updating status: {str(e)}")

    def load_initial_file(self):
        """Load initial file if configured"""
        try:
            default_file = self.settings.settings.get('default_file_path', '')
            if default_file and os.path.exists(default_file):
                print(f"Loading default file: {default_file}")
                
                if self.data_manager.load_file(default_file):
                    # Update the table with the loaded data
                    self.event_handler.update_table_data()
                    self.update_status("File loaded successfully")
                    print("File loaded and table updated")
                else:
                    print("Error loading default file")
                    self.update_status('Error loading default file')
                    
        except Exception as e:
            print(f"Error in load_initial_file: {str(e)}")
            self.update_status(f'Error: {str(e)}')
            traceback.print_exc()

    def run(self):
        """Main application loop"""
        try:
            # Load file once at start of run
            self.load_initial_file()
            
            while True:
                event, values = self.window.read(timeout=100)
                
                if event in (None, 'Exit', sg.WIN_CLOSED, sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
                    break
                
                if event != sg.TIMEOUT_KEY:
                    if not self.event_handler.handle_event(event, values):
                        break
                    
            self.window.close()
            
        except Exception as e:
            print(f"Critical error in run: {str(e)}")
            traceback.print_exc()
            if self.window:
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
   
 