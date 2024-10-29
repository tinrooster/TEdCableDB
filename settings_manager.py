import json
import os

def load_settings():
    settings_file = 'settings.json'
    default_settings = {
        'window_size': (800, 600),
        'window_location': (None, None),
        'last_file_path': '',
        'background_color': '#F0F0F0',
        'function_background_color': '#E0E0E0'
    }
    
    if os.path.exists(settings_file):
        try:
            with open(settings_file, 'r') as f:
                settings = json.load(f)
            for key, value in default_settings.items():
                if key not in settings:
                    settings[key] = value
        except json.JSONDecodeError:
            print(f"Error decoding {settings_file}. Using default settings.")
            settings = default_settings
    else:
        settings = default_settings
    
    return settings

def save_settings(settings):
    settings_file = 'settings.json'
    try:
        with open(settings_file, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"Error saving settings: {str(e)}")

print("settings_manager.py loaded")
