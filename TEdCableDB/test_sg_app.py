import PySimpleGUI as sg
import os
import traceback

print("Starting test app...")

class TestApp:
    def __init__(self):
        print("Initializing...")
        sg.theme('DarkGrey13')
        
        # Basic layout with a table and some buttons
        layout = [
            [sg.Text("Cable Database Test")],
            [sg.Table(values=[], headings=['Col 1', 'Col 2'], 
                     key='-TABLE-', expand_x=True, expand_y=True)],
            [sg.Button("Load File"), sg.Button("Exit")]
        ]
        
        self.window = sg.Window("Test App", layout, resizable=True, finalize=True)
        print("Window created")

    def run(self):
        print("Starting main loop...")
        try:
            while True:
                event, values = self.window.read()
                print(f"Event: {event}")
                
                if event in (sg.WIN_CLOSED, 'Exit'):
                    break
                    
                if event == 'Load File':
                    filename = sg.popup_get_file(
                        'Select a file',
                        no_window=True,
                        file_types=(('Excel Files', '*.xlsx'),)
                    )
                    if filename:
                        print(f"Selected file: {filename}")
                        
        except Exception as e:
            print(f"Error in main loop: {str(e)}")
            traceback.print_exc()
        finally:
            print("Closing window...") 