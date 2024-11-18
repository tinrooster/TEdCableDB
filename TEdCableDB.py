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
import random
from collections import defaultdict

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
        """Initialize settings with default values"""
        self.table_config = {
            'columns': [
                'NUMBER', 'DWG', 'ORIGIN', 'DEST',
                'Wire Type', 'Length', 'Project ID'
            ],
            'required_columns': [
                'NUMBER', 'DWG', 'ORIGIN', 'DEST',
                'Wire Type', 'Length', 'Project ID'
            ],
            'column_widths': {
                'NUMBER': 10,
                'DWG': 15,
                'ORIGIN': 20,
                'DEST': 20,
                'Wire Type': 15,
                'Length': 10,
                'Project ID': 15
            },
            'rows_per_page': 25
        }
        self.config_file = 'config/settings.json'
        self.settings = {}  # For additional runtime settings
        self.load_settings()

    def get_table_config(self):
        """Return the table configuration"""
        return self.table_config

    def __setitem__(self, key, value):
        """Support dictionary-style item assignment"""
        self.settings[key] = value
        self.save_settings()

    def __getitem__(self, key):
        """Support dictionary-style item access"""
        return self.settings.get(key)

    def save_settings(self):
        """Save settings to file"""
        try:
            # Combine table_config and runtime settings
            save_data = {
                'table_config': self.table_config,
                'runtime_settings': self.settings
            }
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w') as f:
                json.dump(save_data, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")
            traceback.print_exc()

    def load_settings(self):
        """Load settings from config file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    loaded_data = json.load(f)
                    # Update table_config with loaded values, keeping defaults if not present
                    if 'table_config' in loaded_data:
                        self.table_config.update(loaded_data['table_config'])
                    # Load runtime settings
                    if 'runtime_settings' in loaded_data:
                        self.settings = loaded_data['runtime_settings']
        except Exception as e:
            print(f"Error loading settings: {e}")
            traceback.print_exc()

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
        self.df = None                # Original dataset
        self.filtered_df = None       # Filtered dataset
        self.display_df = None        # Currently displayed data (filtered or grouped)
        self.is_grouped = False       # Track if we're in a grouped state
        self.column_aliases = {
            'ProjectID': 'Project ID',
            'Project ID': 'Project ID',
            'Project': 'Project ID',
            'PROJECTID': 'Project ID',
            'PROJECT_ID': 'Project ID',
            'PROJECT ID': 'Project ID'
        }
        
    def load_file(self, filename):
        try:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Attempting to load file: {filename}")
            self.df = pd.read_excel(filename)
            
            # Initial sort by NUMBER ascending
            if 'NUMBER' in self.df.columns:
                self.df = self.df.sort_values(by='NUMBER', ascending=True)
                self.current_sort = ('NUMBER', True)
            
            print(f"Successfully processed {len(self.df):,} records")
            return True
        except Exception as e:
            print(f"Error loading file: {str(e)}")
            traceback.print_exc()
            return False

    def sort_data(self, sort_col, ascending=True):
        """Sort the current data by column"""
        try:
            # Use filtered data if it exists, otherwise use main data
            df = self.filtered_df if self.filtered_df is not None else self.df
            
            if df is None:
                print("No data to sort")
                return None
                
            if sort_col not in df.columns:
                print(f"Column {sort_col} not found in data")
                return None
                
            # Sort the data
            self.filtered_df = df.sort_values(by=sort_col, ascending=ascending)
            return self.filtered_df
            
        except Exception as e:
            print(f"Error sorting data: {str(e)}")
            traceback.print_exc()
            return None

    def apply_filters(self, filters, search_mode='standard'):
        """Apply filters to the data"""
        try:
            if self.df is None:
                return None

            df = self.df.copy()
            print(f"Initial data count: {len(df)}")

            # Number range filter
            if 'num_start' in filters or 'num_end' in filters:
                numeric_col = pd.to_numeric(df['NUMBER'], errors='coerce')
                if 'num_start' in filters and filters['num_start']:
                    df = df[numeric_col >= filters['num_start']]
                if 'num_end' in filters and filters['num_end']:
                    df = df[numeric_col <= filters['num_end']]

            # Text filters
            text_fields = {
                'DWG': '-DWG-',
                'ORIGIN': '-ORIGIN-',
                'DEST': '-DEST-',
                'Wire Type': '-WIRE-TYPE-',
                'Length': '-LENGTH-',
                'Project': '-PROJECT-'
            }

            for field, key in text_fields.items():
                if key in filters and filters[key]:
                    value = str(filters[key]).strip().lower()
                    if search_mode == 'exact':
                        df = df[df[field].astype(str).str.lower() == value]
                    elif search_mode == 'fuzzy':
                        df = df[df[field].astype(str).str.lower().str.contains(value, na=False)]
                    else:  # standard
                        df = df[df[field].astype(str).str.lower().str.startswith(value, na=False)]

            self.filtered_df = df
            print(f"Filtered to {len(df)} records")
            return df

        except Exception as e:
            print(f"Error applying filters: {str(e)}")
            traceback.print_exc()
            return None

    def normalize_column_name(self, column_name):
        """Normalize column names to handle variations"""
        try:
            # Check if it's in aliases
            if column_name in self.column_aliases:
                return self.column_aliases[column_name]
            
            # Check if it exists exactly in dataframe
            if column_name in self.df.columns:
                return column_name
                
            # Try case-insensitive match
            for col in self.df.columns:
                if col.lower() == column_name.lower():
                    return col
                    
            print(f"Warning: Column '{column_name}' not found. Available columns: {list(self.df.columns)}")
            return column_name
            
        except Exception as e:
            print(f"Error normalizing column name '{column_name}': {str(e)}")
            return column_name

    def apply_grouping(self, group_by: str) -> bool:
        """Apply grouping while maintaining filtered state"""
        working_df = self.get_current_data()
        
        if working_df is None or group_by not in working_df.columns:
            print(f"Cannot group: invalid column {group_by}")
            return False
        
        try:
            print(f"Grouping by: {group_by}")
            
            # Create summary DataFrame
            grouped = working_df.groupby(group_by, dropna=False)
            summary = []
            
            for name, group in grouped:
                row = {col: '' for col in working_df.columns}
                row[group_by] = str(name) if pd.notna(name) else '(Empty)'
                row['Count'] = len(group)
                
                # Keep first value for other columns
                for col in working_df.columns:
                    if col != group_by and col != 'Count':
                        first_val = group[col].iloc[0] if not group[col].empty else ''
                        row[col] = str(first_val) if pd.notna(first_val) else ''
                
                summary.append(row)
            
            # Convert summary to DataFrame
            summary_df = pd.DataFrame(summary)
            
            # Update the appropriate dataframe
            self.filtered_df = summary_df
            self.current_group = group_by
            
            print(f"Grouped data has {len(summary_df)} rows")
            return True
            
        except Exception as e:
            print(f"Error in grouping: {str(e)}")
            traceback.print_exc()
            return False

    def clear_grouping(self):
        """Clear grouping and return to filtered view"""
        try:
            # Ensure we return to filtered view, not original dataset
            if self.filtered_df is not None:
                self.display_df = self.filtered_df
            else:
                self.display_df = self.df
            self.is_grouped = False
            print(f"Cleared grouping, returned to {len(self.display_df)} records")
            return self.display_df
        except Exception as e:
            print(f"Error clearing grouping: {str(e)}")
            traceback.print_exc()
            return None

    def get_display_data(self):
        """Get current display data"""
        if self.display_df is not None:
            return self.display_df
        if self.filtered_df is not None:
            return self.filtered_df
        return self.df

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
        self.poker_game = None  # Initialize as None
        self.mad_panda_art = """                                                          
          ▒▒▒▒▒▒  ▒▒▒▒▒▒▒▒▒▒▒▒▒▒  ▒▒▒▒▒▒          
        ▒▒░░░░░░▒▒░░░░░░░░░░░░░░▒▒░░░░░░▒▒        
      ▒▒░░░░░░▒▒░░░░░░░░░░░░░░░░░░▒▒░░░░░░▒▒      
      ▒▒░░░░▒▒░░░░░░░░░░░░░░░░░░░░░░▒▒░░░░▒▒      
      ▒▒░░░░▒▒░░░░░░░░░░██░░░░░░░░░░▒▒░░░░▒▒      
      ▒▒▒▒▒▒░░░░██░░░░██████░░░░██░░░░▒▒▒▒        
      ▒▒▒▒▒▒░░░░██░░░░█████░░░░██░░░░▒▒          
      ▓▓▒▒▓▓▒▒░░░░░░░░██████░░░░░░░░░░▒▒          
      ▓▓▒▒▒▒▒▒▒▒░░░░░░░░░░░░░░░░░░░░░░▒▒          
        ▓▓▓▓▒▒▒▒▒▒░░░░░░░░░░░░░░░░░░▒▒            
        ▓▓▓▓▒▒▓▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒              
      ▒▒▒▒▓▓▒▒▒▒▓▓▒▒░░░░░░░░░░▒▒░░▒▒              
      ▒▒▒▒▓▓▒▒▒▒▒▒▒▒░░░░░░░░▒▒▒▒░░▒▒              
          ▒▒▓▓▒▒▒▒▓▓▒▒▒▒▒▒▒▒░░▒▒░░▒▒              
        ▒▒▒▒▓▓▒▒▒▒▓▓▒▒░░░░░░░░▒▒░░▒▒              
              ▓▓▒▒▒▒▒▒░░░░░░▒▒░░▒▒                
              ▓▓▒▒▒▒▓▓▒▒▒▒▒▒▒▒▒▒                  
              ▓▓▓▓▒▒▓▓▓▓▒▒                        
"""
        self.keyboard_bindings = {
            '<Control-o>': self.handle_open,
            '<Control-s>': self.handle_save,
            '<Control-e>': self.handle_export,
            '<Control-f>': self.handle_find,
            '<Control-r>': self.handle_refresh,
            '<Escape>': self.handle_clear_filter
        }
        self.bind_keyboard_shortcuts()
        self.update_status_counts()
        self.file_manager = FileManager()
        self.current_file = self.load_last_file_path()
        
        # If we have a last file, try to load it
        if self.current_file and os.path.exists(self.current_file):
            print(f"Loading last file: {self.current_file}")
            if self.data_manager.load_file(self.current_file):
                self.update_table_data()
                self.window['-STATUS-'].update(f'Loaded: {self.current_file}')

    def handle_event(self, event, values):
        """Main event handler"""
        try:
            print(f"Handling event: {event}")
            
            # Handle Enter key for filter inputs
            if isinstance(event, str) and ('Return' in event or event.endswith('\r')):
                # Check if any filter input has focus
                filter_inputs = ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', 
                               '-DEST-', '-WIRE-TYPE-', '-LENGTH-', '-PROJECT-']
                focused = self.window.find_element_with_focus()
                if focused and focused.Key in filter_inputs:
                    print("Enter key pressed in filter input - applying filters")
                    self.handle_filter_event(values)
                    return True

            # Handle window close
            if event in (None, 'Exit', sg.WIN_CLOSED):
                print("**** EXITING ****")
                return False
                
            # Ignore mouse wheel events
            if isinstance(event, str) and 'MouseWheel' in event:
                return True
                
            # Handle input focus events without additional processing
            input_fields = ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', 
                          '-DEST-', '-WIRE-TYPE-', '-LENGTH-', '-PROJECT-']
            if event in input_fields:
                return True

            # Menu events
            if event.startswith('About'):
                self.handle_help_event('About')
            elif event.startswith('Quick Guide'):
                self.handle_help_event('Quick Guide')
            elif event.startswith('Shortcuts'):
                self.handle_help_event('Shortcuts')
                
            # File operations
            elif event in ('Open::open_key', '-OPEN-'):
                self.handle_open()
            elif event in ('Save::save_key', '-SAVE-'):
                self.handle_save()
            elif event in ('Export::export_key', '-EXPORT-'):
                self.handle_export()
            elif event == 'Import::import_key':
                self.handle_import()
                
            # Filter and sort operations
            elif event == '-APPLY-FILTER-':
                self.handle_filter_event(values)
            elif event == '-CLEAR-FILTER-':
                self.handle_clear_filters()
            elif event == '-APPLY-SORT-':
                self.handle_sort_event(values)
            elif event == '-APPLY-GROUP-':
                self.handle_group_event(values)
            elif event == '-CLEAR-GROUP-':
                self.data_manager.filtered_df = None
                self.update_table_data()
                
            return True
            
        except Exception as e:
            print(f"Error handling event {event}: {str(e)}")
            traceback.print_exc()
            return True

    def handle_file_open(self):
        """Handle file open operation"""
        try:
            import tkinter as tk
            from tkinter import filedialog
            
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            
            initial_dir = os.path.dirname(self.current_file) if self.current_file else ''
            filename = filedialog.askopenfilename(
                title='Open Excel File',
                initialdir=initial_dir,
                filetypes=[
                    ('Excel Files', '*.xlsx;*.xlsm'),
                    ('All Files', '*.*')
                ]
            )
            
            if filename:
                if self.data_manager.load_file(filename):
                    self.current_file = filename
                    self.save_last_file_path(filename)
                    self.update_table_data()
                    self.window['-STATUS-'].update(f'Loaded: {filename}')
                    print(f"Successfully loaded: {filename}")
            
            root.destroy()
            
        except Exception as e:
            print(f"Error in file open: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f"Error opening file: {str(e)}")

    def handle_file_save(self, save_as=False):
        """Handle file save operation"""
        try:
            if self.data_manager.df is None:
                sg.popup_error('No data to save')
                return

            import tkinter as tk
            from tkinter import filedialog
            
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            
            if save_as or not self.current_file:
                initial_dir = os.path.dirname(self.current_file) if self.current_file else ''
                filename = filedialog.asksaveasfilename(
                    title='Save As',
                    initialdir=initial_dir,
                    defaultextension='.xlsx',
                    filetypes=[('Excel Files', '*.xlsx')]
                )
                if not filename:
                    root.destroy()
                    return
            else:
                filename = self.current_file

            # Save the file
            df = self.data_manager.get_current_data()
            if df is not None:
                df.to_excel(filename, index=False)
                
                # Update current file and save to config
                self.current_file = filename
                self.save_last_file_path(filename)
                
                # Update UI with feedback
                self.window['-STATUS-'].update(f'Saved: {filename}')
                sg.popup_quick_message('File Saved Successfully', 
                                     background_color='green',
                                     text_color='white',
                                     auto_close_duration=2)
                print(f"Successfully saved: {filename}")
            else:
                print("Error: No data to save")
                sg.popup_error("No data to save")
            
            root.destroy()
            
        except Exception as e:
            print(f"Error in file save: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f"Error saving file: {str(e)}")

    def load_last_file_path(self):
        """Load the last file path from config"""
        try:
            if os.path.exists('config/last_file.json'):
                with open('config/last_file.json', 'r') as f:
                    config = json.load(f)
                    return config.get('last_file')
        except Exception as e:
            print(f"Error loading last file path: {e}")
        return None

    def save_last_file_path(self, path):
        """Save the last file path to config"""
        try:
            os.makedirs('config', exist_ok=True)
            with open('config/last_file.json', 'w') as f:
                json.dump({'last_file': path}, f, indent=4)
        except Exception as e:
            print(f"Error saving last file path: {e}")

    def handle_open(self, event=None):
        """Handle file open operation"""
        try:
            import tkinter as tk
            from tkinter import filedialog
            
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            
            initial_dir = os.path.dirname(self.current_file) if self.current_file else ''
            filename = filedialog.askopenfilename(
                title='Open File',
                initialdir=initial_dir,
                filetypes=[
                    ('Excel Files', '*.xlsx;*.xlsm'),
                    ('All Files', '*.*')
                ]
            )
            
            if filename:
                if self.data_manager.load_file(filename):
                    self.update_table_data()
                    self.window['-STATUS-'].update(f'Loaded: {filename}')
                    # Save last directory
                    self.settings['last_directory'] = os.path.dirname(filename)
            
            root.destroy()
        except Exception as e:
            print(f"Error in handle_open: {str(e)}")
            self.window['-STATUS-'].update('Error opening file')

    def handle_save(self, event=None):
        """Handle file save operation"""
        try:
            if self.data_manager.df is None:
                sg.popup_error('No data to save')
                return
                
            import tkinter as tk
            from tkinter import filedialog
            
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            
            initial_dir = os.path.dirname(self.current_file) if self.current_file else ''
            filename = filedialog.asksaveasfilename(
                title='Save As',
                initialdir=initial_dir,
                defaultextension='.xlsx',
                filetypes=[('Excel Files', '*.xlsx')]
            )
            
            if filename:
                self.data_manager.df.to_excel(filename, index=False)
                self.window['-STATUS-'].update(f'Saved: {filename}')
                self.settings['last_directory'] = os.path.dirname(filename)
            
            root.destroy()
        except Exception as e:
            print(f"Error in handle_save: {str(e)}")
            self.window['-STATUS-'].update('Error saving file')

    def handle_export(self, event=None):
        """Handle export operation"""
        try:
            if self.data_manager.df is None:
                sg.popup_error('No data to export')
                return
                
            import tkinter as tk
            from tkinter import filedialog
            
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            
            initial_dir = os.path.dirname(self.current_file) if self.current_file else ''
            filename = filedialog.asksaveasfilename(
                title='Export As',
                initialdir=initial_dir,
                defaultextension='.xlsx',
                filetypes=[('Excel Files', '*.xlsx')]
            )
            
            if filename:
                if filename.endswith('.csv'):
                    self.data_manager.df.to_csv(filename, index=False)
                else:
                    self.data_manager.df.to_excel(filename, index=False)
                self.window['-STATUS-'].update(f'Exported: {filename}')
                self.settings['last_directory'] = os.path.dirname(filename)
            
            root.destroy()
        except Exception as e:
            print(f"Error in handle_export: {str(e)}")
            self.window['-STATUS-'].update('Error exporting file')

    def handle_find(self, event=None):
        """Handle find operation - focus on filter input"""
        try:
            self.window['-NUM-START-'].set_focus()
        except Exception as e:
            print(f"Error in handle_find: {str(e)}")

    def handle_refresh(self, event=None):
        """Handle refresh operation"""
        try:
            self.update_table_data()
            self.window['-STATUS-'].update('Display refreshed')
        except Exception as e:
            print(f"Error in handle_refresh: {str(e)}")
            self.window['-STATUS-'].update('Error refreshing display')

    def handle_clear_filter(self, event=None):
        """Handle clear filter operation"""
        try:
            self.handle_clear_filters()
            self.window['-STATUS-'].update('Filters cleared')
        except Exception as e:
            print(f"Error in handle_clear_filter: {str(e)}")
            self.window['-STATUS-'].update('Error clearing filters')

    def create_about_window(self):
        """Create the About window"""
        layout = [
            [sg.Text("TEd Cable DB", font=("Impact", 20), justification='center', pad=(0,20))],
            [sg.Text("Version 1.0", font=("Helvetica", 10))],
            [sg.HorizontalSeparator()],
            
            [sg.Text("KGO-TV Engineering Department", 
                    font=("Helvetica", 12, "bold"), 
                    justification='center',
                    pad=(0,10))],
            [sg.HorizontalSeparator()],
            
            [sg.Text("Engineering Staff:", font=("Helvetica", 10, "bold"), pad=(0,10))],
            [sg.Column([
                [sg.Text("David Fortin", key='-DAVE-', enable_events=True, pad=(0,5))],
                [sg.Text("Marcus Saxton", pad=(0,5))],
                [sg.Text("Dave Figura", pad=(0,5))],
                [sg.Text("Jack Frasier", pad=(0,5))],
                [sg.Text("Rosendo Pena", pad=(0,5))],
                [sg.Text("Felice Gandolfo", key='-FELICE-', enable_events=True, pad=(0,5))],
            ], pad=(20, 0))],
            
            [sg.HorizontalSeparator()],
            [sg.Text("Developed by:", font=("Helvetica", 10, "bold"), pad=(0,10))],
            [sg.Text("AC Hay", pad=(20, 0))],
            
            [sg.HorizontalSeparator()],
            [sg.Text("Special Thanks:", font=("Helvetica", 10, "bold"), pad=(0,10))],
            [sg.Text("Claude 3.5 Sonnet (Anthropic)", pad=(20, 0))],
            
            [sg.HorizontalSeparator()],
            [sg.Button("OK", key="-HELP-OK-", pad=(0,20))],
            
            # Hidden frame for Easter egg
            [sg.Frame('', [[sg.Text(self.mad_panda_art, font='Courier 8', 
                    key='-MAD-SCIENTIST-')]], key='-MAD-FRAME-', visible=False)]
        ]
        
        return sg.Window(
            "About TEd Cable DB",
            layout,
            modal=True,
            finalize=True,
            element_justification='center',
            font=("Helvetica", 10),
            keep_on_top=True
        )

    def handle_help_event(self, event):
        """Handle help menu events"""
        try:
            print(f"Processing help event: {event}")
            if event == "About":
                about_window = self.create_about_window()
                while True:
                    try:
                        event, values = about_window.read()
                        if event in (sg.WIN_CLOSED, "-HELP-OK-"):
                            break
                    except Exception as e:
                        print(f"Error in About window event loop: {str(e)}")
                        traceback.print_exc()
                        break
                about_window.close()
            elif event == "Quick Guide":
                print("Quick Guide not implemented yet")
                sg.popup_error("Quick Guide not implemented yet")
            elif event == "Shortcuts":
                print("Shortcuts not implemented yet")
                sg.popup_error("Shortcuts not implemented yet")
                
        except Exception as e:
            print(f"Error in help event: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f'Error displaying help: {str(e)}')

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
        """Update table with current display data"""
        try:
            display_data = self.data_manager.get_display_data()
            if display_data is not None:
                # Convert DataFrame to list of lists for display
                table_data = display_data.fillna('').values.tolist()
                self.window['-TABLE-'].update(values=table_data)
                
                # Update record count - using correct element key
                total_records = len(self.data_manager.df)
                current_records = len(display_data)
                self.window['-RECORDS-COUNT-'].update(
                    f'{current_records:,}'
                )
                self.window['-FILTER-STATUS-'].update(
                    f'of {total_records:,} total' if current_records != total_records else ''
                )
                print(f"Table updated with {current_records:,} records")
                
        except Exception as e:
            print(f"Error updating table: {str(e)}")
            traceback.print_exc()

    def handle_filter_event(self, values):
        """Handle filter application"""
        try:
            print("Processing filter request...")
            if self.data_manager.df is None:
                sg.popup_error("No data loaded to filter")
                return

            filters = {}
            
            # Number range filter
            try:
                if values['-NUM-START-']:
                    filters['num_start'] = float(values['-NUM-START-'])
                if values['-NUM-END-']:
                    filters['num_end'] = float(values['-NUM-END-'])
            except ValueError:
                sg.popup_error('Invalid number in filter range')
                return

            # Text filters
            text_fields = {
                '-DWG-': 'DWG',
                '-ORIGIN-': 'ORIGIN',
                '-DEST-': 'DEST',
                '-WIRE-TYPE-': 'Wire Type',
                '-LENGTH-': 'Length',
                '-PROJECT-': 'Project ID'
            }

            for key, field in text_fields.items():
                if values[key]:
                    filters[field] = values[key].strip()

            # Apply filters
            filtered_df = self.data_manager.apply_filters(filters)
            if filtered_df is not None:
                self.update_table_data()
                count = len(filtered_df)
                total = len(self.data_manager.df)
                self.window['-STATUS-'].update(f'Filtered: {count:,} of {total:,} records')
            
        except Exception as e:
            print(f"Error in filter operation: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f'Error applying filters: {str(e)}')

    def handle_clear_filters(self):
        """Handle clear filters event"""
        try:
            # Clear the filter inputs
            filter_keys = ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', 
                      '-DEST-', '-WIRE-TYPE-', '-LENGTH-', '-PROJECT-']
            for key in filter_keys:
                self.window[key].update('')
            
            # Clear the filters in data manager
            df = self.data_manager.clear_filters()
            if df is not None:
                self.update_table(df)
            
        except Exception as e:
            print(f"Error clearing filters: {str(e)}")
            traceback.print_exc()

    def handle_group_event(self, values):
        """Handle grouping events"""
        try:
            group_by = values['-GROUP-BY-']
            if not group_by:
                return
                
            # Show processing indicator for large datasets
            source_df = self.data_manager.filtered_df if self.data_manager.filtered_df is not None else self.data_manager.df
            if len(source_df) > 1000:
                sg.popup_quick_message(
                    f"Grouping {len(source_df):,} records...\nThis may take a moment.",
                    auto_close=True,
                    auto_close_duration=2,
                    non_blocking=True
                )
                
            grouped_df = self.data_manager.apply_grouping(group_by)
            if grouped_df is not None:
                self.data_manager.display_df = grouped_df  # Set the display DataFrame
                self.update_table_data()  # Update the table
                print(f"Grouped by {group_by}: {len(grouped_df)} groups")
                
        except Exception as e:
            print(f"Error in group event: {str(e)}")
            traceback.print_exc()

    def handle_clear_group(self):
        """Handle clear group operation"""
        try:
            # Restore the filtered view, not the original dataset
            self.data_manager.display_df = self.data_manager.filtered_df if self.data_manager.filtered_df is not None else self.data_manager.df
            self.update_table_data()
            self.window['-GROUP-BY-'].update('')  # Clear the group by selection
            
        except Exception as e:
            print(f"Error clearing group: {str(e)}")
            traceback.print_exc()

    def handle_sort_event(self, values):
        """Handle sorting of data"""
        try:
            print("Processing sort request...")
            sort_col = values['-SORT-BY-']
            if not sort_col:
                print("No sort column selected")
                return

            ascending = values['-SORT-ASC-']
            
            # Apply sort
            sorted_df = self.data_manager.sort_data(sort_col, ascending)
            if sorted_df is not None:
                self.update_table_data()
                direction = "ascending" if ascending else "descending"
                self.window['-STATUS-'].update(f'Sorted by {sort_col} ({direction})')
            
        except Exception as e:
            print(f"Error in sort operation: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f'Error sorting data: {str(e)}')

    def handle_copy_selection(self):
        """Copy selected rows to clipboard"""
        try:
            selected_rows = self.window['-TABLE-'].SelectedRows
            if not selected_rows:
                return
            
            df = self.data_manager.get_current_data()
            if df is None:
                return
                
            # Get selected data
            selected_data = df.iloc[selected_rows]
            
            # Copy to clipboard
            selected_data.to_clipboard(index=False)
            self.window['-STATUS-'].update('Selection copied to clipboard')
            
        except Exception as e:
            print(f"Error copying selection: {str(e)}")
            self.window['-STATUS-'].update('Error copying selection')

    def handle_export_selection(self):
        """Export selected rows to Excel"""
        try:
            selected_rows = self.window['-TABLE-'].SelectedRows
            if not selected_rows:
                sg.popup_error('No rows selected')
                return
            
            df = self.data_manager.get_current_data()
            if df is None:
                return
                
            # Get selected data
            selected_data = df.iloc[selected_rows]
            
            # Get save path
            save_path = sg.popup_get_file(
                'Save As',
                save_as=True,
                file_types=(('Excel Files', '*.xlsx'),),
                default_extension='xlsx'
            )
            
            if save_path:
                selected_data.to_excel(save_path, index=False)
                self.window['-STATUS-'].update(f'Selection exported to {save_path}')
                
        except Exception as e:
            print(f"Error exporting selection: {str(e)}")
            self.window['-STATUS-'].update('Error exporting selection')

    def bind_keyboard_shortcuts(self):
        """Bind keyboard shortcuts for common actions"""
        try:
            # Define keyboard shortcuts
            self.window.bind('<Control-o>', '-OPEN-')  # Ctrl+O for open
            self.window.bind('<Control-s>', '-SAVE-')  # Ctrl+S for save
            self.window.bind('<Control-e>', '-EXPORT-')  # Ctrl+E for export
            self.window.bind('<Control-f>', '-FIND-')  # Ctrl+F for find/filter
            self.window.bind('<Control-r>', '-REFRESH-')  # Ctrl+R for refresh
            self.window.bind('<Escape>', '-CLEAR-FILTER-')  # Esc to clear filters
            
        except Exception as e:
            print(f"Error binding keyboard shortcuts: {str(e)}")
            traceback.print_exc()

    def update_status_counts(self):
        """Update record counts in status bar"""
        try:
            if self.data_manager.df is not None:
                total_records = len(self.data_manager.df)
                filtered_records = len(self.data_manager.filtered_df) if self.data_manager.filtered_df is not None else total_records
                
                # Update records count
                self.window['-RECORDS-COUNT-'].update(f'{filtered_records:,}')
                
                # Update filter status if filtered
                if self.data_manager.filtered_df is not None:
                    self.window['-FILTER-STATUS-'].update(f'of {total_records:,} total')
                else:
                    self.window['-FILTER-STATUS-'].update('')
                    
        except Exception as e:
            print(f"Error updating status counts: {str(e)}")
            traceback.print_exc()

    def handle_import(self):
        """Handle import operation"""
        try:
            print("Processing import request...")
            import tkinter as tk
            from tkinter import filedialog
            
            root = tk.Tk()
            root.withdraw()
            
            filename = filedialog.askopenfilename(
                title='Import File',
                filetypes=[
                    ('Excel Files', '*.xlsx;*.xlsm'),
                    ('CSV Files', '*.csv'),
                    ('All Files', '*.*')
                ]
            )
            
            if filename:
                # TODO: Implement actual import logic
                print(f"Import from {filename} not implemented yet")
                sg.popup_error("Import functionality not implemented yet")
            
            root.destroy()
            
        except Exception as e:
            print(f"Error in import: {str(e)}")
            traceback.print_exc()
            sg.popup_error(f"Error importing: {str(e)}")

class UIBuilder:
    def __init__(self):
        """Initialize settings with proper file paths"""
        self.window_title = "TE/d Cable DB v1.0"
        # Add table configuration
        self.table_config = {
            'columns': ['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Wire Type', 'Length', 'Project'],
            'column_widths': {
                'NUMBER': 10,
                'DWG': 15,
                'ORIGIN': 20,
                'DEST': 20,
                'Wire Type': 15,
                'Length': 10,
                'Project': 15
            }
        }
        self.menu_def = [
            ['&File', [
                '&Open::open_key',
                '&Save::save_key',
                'Save &As::saveas_key',
                '---',
                '&Import::import_key',
                '&Export::export_key',
                '---',
                'E&xit'
            ]],
            ['&Help', ['&Quick Guide', '&Shortcuts', '&About']]
        ]

    def create_window(self):
        """Create the main window"""
        # Create menu
        menu_def = [
            ['&File', ['&Open::open_key', '&Save::save_key', 'Save &As::saveas_key', '---', 
                      '&Import::import_key', '&Export::export_key', '---', 'E&xit']],
            ['&Help', ['&Quick Guide', '&Shortcuts', '&About']]
        ]

        # Main layout
        layout = [
            [sg.Menu(menu_def)],
            
            # Search and Filter Section
            [
                sg.Column([
                    [self.create_filter_frame()],  # Left column with filters
                ], vertical_alignment='top'),
                
                sg.Column([
                    [self.create_sort_frame()],    # Right column with sort/group
                ], vertical_alignment='top')
            ],
            
            # Table Section
            [sg.Table(
                values=[],
                headings=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Wire Type', 'Length', 'Project'],
                auto_size_columns=True,
                justification='left',
                key='-TABLE-',
                enable_events=True,
                expand_x=True,
                expand_y=True,
                enable_click_events=True
            )],
            
            # Status Bar with Records Count
            [
                sg.Text('', key='-STATUS-', size=(40, 1)),
                sg.Text('Records:', pad=(10,0)),
                sg.Text('0', key='-RECORDS-COUNT-', size=(10, 1)),
                sg.Text('', key='-FILTER-STATUS-', size=(30, 1))
            ]
        ]

        return sg.Window(
            'TEd Cable DB v1.0',
            layout,
            resizable=True,
            finalize=True,
            return_keyboard_events=True
        )

    def create_filter_frame(self):
        """Create the filter frame with all filter options"""
        return sg.Frame('Filters', [
            [sg.Text('Search Options')],
            [
                sg.Radio('Standard Search', 'SEARCH', key='-STANDARD-SEARCH-', default=True),
                sg.Radio('Exact Match', 'SEARCH', key='-EXACT-'),
                sg.Radio('Fuzzy Search', 'SEARCH', key='-FUZZY-SEARCH-')
            ],
            [sg.Text('NUMBER:', size=(8,1)), 
             sg.Input(size=(10, 1), key='-NUM-START-', enable_events=True),
            [sg.Text('DWG:', size=(8,1)), 
             sg.Input(size=(30, 1), key='-DWG-', enable_events=True)],
            [sg.Text('ORIGIN:', size=(8,1)), 
             sg.Input(size=(30, 1), key='-ORIGIN-', enable_events=True)],
            [sg.Text('DEST:', size=(8,1)), 
             sg.Input(size=(30, 1), key='-DEST-', enable_events=True)],
            [sg.Text('Wire Type:', size=(8,1)), 
             sg.Input(size=(30, 1), key='-WIRE-TYPE-', enable_events=True)],
            [sg.Text('Length:', size=(8,1)), 
             sg.Input(size=(30, 1), key='-LENGTH-', enable_events=True)],
            [sg.Text('Project:', size=(8,1)), 
             sg.Input(size=(30, 1), key='-PROJECT-', enable_events=True)],
            [sg.Button('Apply Filters', key='-APPLY-FILTER-', bind_return_key=True),
             sg.Button('Clear Filters', key='-CLEAR-FILTER-')]
        ])

    def create_sort_frame(self):
        """Create the sort and group frame"""
        return sg.Frame('Sort and Group', [
            [sg.Text('Sort by:'),
             sg.Combo(['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Wire Type', 'Length', 'Project'],
                     default_value='NUMBER',
                     key='-SORT-BY-',
                     size=(15, 1)),
             sg.Radio('Ascending', 'SORT', key='-SORT-ASC-', default=True),
             sg.Radio('Descending', 'SORT', key='-SORT-DESC-'),
             sg.Button('Sort', key='-APPLY-SORT-')],
            [sg.Text('Group by:'),
             sg.Combo(['DWG', 'ORIGIN', 'DEST', 'Wire Type', 'Length', 'ProjectID'],
                     key='-GROUP-BY-',
                     size=(15, 1)),
             sg.Button('Apply', key='-APPLY-GROUP-'),
             sg.Button('Clear', key='-CLEAR-GROUP-')]
        ])

class FileManager:
    def __init__(self):
        self.config_file = "config.json"
        self.default_config = {
            "last_file": None,
            "save_directory": None,
            "settings": {
                "window_size": [800, 600],
                "window_location": None,
                "last_directory": None
            }
        }
        self.config = self.load_config()
        
    def load_config(self):
        """Load configuration from JSON file or create default"""
        try:
            if os.path.exists(self.config_file):
                try:
                    with open(self.config_file, 'r') as f:
                        config = json.load(f)
                    # Validate config structure
                    if not isinstance(config, dict):
                        raise ValueError("Invalid config format")
                    return config
                except (json.JSONDecodeError, ValueError):
                    print("Invalid config file, creating new one")
                    os.remove(self.config_file)
                    return self.create_default_config()
            else:
                return self.create_default_config()
        except Exception as e:
            print(f"Error loading config: {e}")
            return self.default_config.copy()
            
    def create_default_config(self):
        """Create and save default configuration"""
        config = self.default_config.copy()
        self.save_config(config)
        return config

    def update_status_counts(self):
        """Update record counts in status bar"""
        try:
            table = self.window['-TABLE-']
            if table and hasattr(table, 'Values'):
                total_records = len(table.Values) if table.Values else 0
                selected_records = len(table.SelectedRows) if hasattr(table, 'SelectedRows') else 0
                
                self.window['-RECORDS-COUNT-'].update(f'{total_records:,}')
                self.window['-SELECTED-COUNT-'].update(f'{selected_records:,}')
        except Exception as e:
            print(f"Error updating status counts: {str(e)}")

class CableDatabaseApp:
    def __init__(self):
        print("Application starting...")
        sg.theme('DarkBlue3')
        
        # Initialize components
        self.settings = Settings()
        self.data_manager = DataManager(self.settings)
        self.ui_builder = UIBuilder()
        self.window = self.ui_builder.create_window()
        
        # Initialize event handler
        self.event_handler = EventHandler(
            window=self.window,
            data_manager=self.data_manager,
            settings=self.settings
        )
        
        # Remove the duplicate load - EventHandler now handles initial load

    def run(self):
        """Main application loop"""
        try:
            print("App run started...")
            
            while True:
                event, values = self.window.read(timeout=100)
                
                if event in (None, 'Exit', sg.WIN_CLOSED):
                    break
                
                if event != sg.TIMEOUT_KEY:
                    if not self.event_handler.handle_event(event, values):
                        break
            
            self.window.close()
            print("App run completed")
            
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
    except Exception as e:
        print(f"Critical error: {str(e)}")
        traceback.print_exc()
        sg.popup_error(f"Critical Error: {str(e)}")
   
 