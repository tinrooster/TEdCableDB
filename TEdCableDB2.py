# Update the window title and about information
class UIBuilder:
    def __init__(self):
        self.window_title = "TEd Cable DB v1.0"  # Updated name
        self.menu_def = [
            ['File', ['Open::open_key', 'Save::save_key', 'Save As::saveas_key', '---', 'Exit']],
            ['Help', ['Quick Guide', 'Shortcuts', 'About']]
        ]

    def create_help_window(self, help_type):
        """Create help window based on type"""
        if help_type == "About":
            layout = [
                [sg.Text("TEd Cable DB", font=("Helvetica", 16))],
                [sg.Text("Version 1.0 - KGO Engineering", font=("Helvetica", 10))],
                [sg.Text("_" * 50)],
                [sg.Text("Developed by:", font=("Helvetica", 10, "bold"))],
                [sg.Text("AC Hay")],
                [sg.Text("_" * 50)],
                [sg.Text("Special Thanks:", font=("Helvetica", 10, "bold"))],
                [sg.Text("Anthropic Claude AI Assistant")],
                [sg.Text("_" * 50)],
                [sg.Text("KGO Engineering Department:", font=("Helvetica", 10, "bold"))],
                [sg.Text("Dave Fortin\nDavid Figura\nMarcus Saxton")],
                [sg.Text("Jack Fraiser\nRosendo Pena")],
                [sg.Text("and especially")],
                [sg.Text("Felice Gondolfo", font=("Helvetica", 10, "bold"))],
                [sg.Text("_" * 50)],
                [sg.Button("OK", key="-HELP-OK-")]
            ] 