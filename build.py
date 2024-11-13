import os
import PyInstaller.__main__

def build_exe():
    PyInstaller.__main__.run([
        'TEdCableDB.py',
        '--onefile',
        '--windowed',
        '--name=TEd Cable DB',
        '--icon=app_icon.ico',
        '--add-data=config.json;.',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=thefuzz',
        '--clean'
    ])

if __name__ == "__main__":
    build_exe() 