import PySimpleGUI as sg

print("Starting test...")

try:
    sg.theme('DarkGrey13')
    print("Theme set")
    
    layout = [[sg.Text("Hello World")], [sg.Button("OK")]]
    print("Layout created")
    
    window = sg.Window("Test Window", layout)
    print("Window created")
    
    while True:
        event, values = window.read()
        print(f"Event: {event}")
        if event == sg.WIN_CLOSED or event == "OK":
            break
    
    window.close()
    print("Test complete")

except Exception as e:
    print(f"Error: {str(e)}")
    import traceback
    traceback.print_exc() 