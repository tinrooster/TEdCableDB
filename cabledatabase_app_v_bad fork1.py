import PySimpleGUI as sg
import pandas as pd
import traceback

def create_main_layout():
    # Left panel - Filters
    filter_column = [
        [sg.Frame('Filters', [
            [sg.Text('NUMBER:', size=(8,1)), 
             sg.Input(size=(8,1), key='-NUM-START-'), 
             sg.Text('to'), 
             sg.Input(size=(8,1), key='-NUM-END-')],
            [sg.Text('DWG:', size=(8,1)), 
             sg.Input(size=(15,1), key='-DWG-'), 
             sg.Checkbox('Exact', key='-DWG-EXACT-')],
            [sg.Text('ORIGIN:', size=(8,1)), 
             sg.Input(size=(15,1), key='-ORIGIN-'), 
             sg.Checkbox('Exact', key='-ORIGIN-EXACT-')],
            [sg.Text('DEST:', size=(8,1)), 
             sg.Input(size=(15,1), key='-DEST-'), 
             sg.Checkbox('Exact', key='-DEST-EXACT-')],
            [sg.Text('Alt DWG:', size=(8,1)), 
             sg.Input(size=(15,1), key='-ALT-DWG-'), 
             sg.Checkbox('Exact', key='-ALT-DWG-EXACT-')],
            [sg.Text('Wire Type:', size=(8,1)), 
             sg.Input(size=(15,1), key='-WIRE-'), 
             sg.Checkbox('Exact', key='-WIRE-EXACT-')],
            [sg.Text('Length:', size=(8,1)), 
             sg.Input(size=(15,1), key='-LENGTH-'), 
             sg.Checkbox('Exact', key='-LENGTH-EXACT-')],
            [sg.Text('Note:', size=(8,1)), 
             sg.Input(size=(15,1), key='-NOTE-'), 
             sg.Checkbox('Exact', key='-NOTE-EXACT-')],
            [sg.Text('Project ID:', size=(8,1)), 
             sg.Input(size=(15,1), key='-PROJECT-ID-'), 
             sg.Checkbox('Exact', key='-PROJECT-ID-EXACT-')],
            [sg.Button('Filter', size=(10,1), button_color=('white', 'navy'))],
            [sg.Button('Clear Filters', size=(10,1), button_color=('white', 'maroon'))]
        ])]
    ]

    # Right panel - Table
    table = sg.Table(
        values=[[]],
        headings=['NUMBER', 'DWG', 'ORIGIN', 'DEST', 'Alternate Dwg', 'Wire Type', 'Length', 'Note', 'Project ID'],
        auto_size_columns=True,
        display_row_numbers=True,
        justification='left',
        num_rows=25,
        key='-TABLE-',
        selected_row_colors=('white', 'blue'),
        enable_events=True,
        expand_x=True,
        expand_y=True,
        enable_click_events=True,
        right_click_menu=['&Right', ['Hide Column', 'Show All Columns', 'Sort Ascending', 'Sort Descending']]
    )

    # Complete layout
    layout = [
        [sg.Column(filter_column, vertical_alignment='top'), 
         sg.VSeparator(),
         sg.Column([[table]], expand_x=True, expand_y=True)]
    ]
    
    return layout

def main():
    sg.theme('SystemDefault')
    
    try:
        # Create window
        window = sg.Window(
            'Cable Database Interface',
            create_main_layout(),
            resizable=True,
            finalize=True,
            size=(1200, 800)
        )
        
        # Load data
        df = pd.read_excel("D:/code/Cable DB 3 _ test.xlsm")
        window['-TABLE-'].update(values=df.values.tolist())
        
        while True:
            event, values = window.read()
            
            if event == sg.WIN_CLOSED:
                break
                
            try:
                # Handle table column operations
                if event == 'Hide Column':
                    col_num = window['-TABLE-'].SelectedColumns[0] if hasattr(window['-TABLE-'], 'SelectedColumns') else None
                    if col_num is not None:
                        window['-TABLE-'].hide_column(col_num)
                
                elif event == 'Show All Columns':
                    for i in range(len(df.columns)):
                        window['-TABLE-'].unhide_column(i)
                
                elif event == 'Sort Ascending':
                    col_num = window['-TABLE-'].SelectedColumns[0] if hasattr(window['-TABLE-'], 'SelectedColumns') else None
                    if col_num is not None:
                        col_name = df.columns[col_num]
                        df = df.sort_values(by=col_name, ascending=True)
                        window['-TABLE-'].update(values=df.values.tolist())
                
                elif event == 'Sort Descending':
                    col_num = window['-TABLE-'].SelectedColumns[0] if hasattr(window['-TABLE-'], 'SelectedColumns') else None
                    if col_num is not None:
                        col_name = df.columns[col_num]
                        df = df.sort_values(by=col_name, ascending=False)
                        window['-TABLE-'].update(values=df.values.tolist())
                
                elif event == 'Clear Filters':
                    for key in ['-NUM-START-', '-NUM-END-', '-DWG-', '-ORIGIN-', '-DEST-', 
                               '-ALT-DWG-', '-WIRE-', '-LENGTH-', '-NOTE-', '-PROJECT-ID-']:
                        window[key].update('')
                        if f'{key}-EXACT-' in values:
                            window[f'{key}-EXACT-'].update(False)
                    window['-TABLE-'].update(values=df.values.tolist())
                
            except Exception as e:
                print(f"Error handling event {event}: {str(e)}")
                traceback.print_exc()
                continue
            
    except Exception as e:
        print(f"Error in main loop: {str(e)}")
        traceback.print_exc()
        sg.popup_error(f"An error occurred: {str(e)}")
    finally:
        if 'window' in locals():
            window.close()
        print("Application closing...")

if __name__ == "__main__":
    main()
    
    























































































