import cx_Freeze, sys

base = None

if sys.platform == 'win32' or sys.platform == 'win64':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("main.py", base=base, icon="main_icon.ico", targetName="Parser")]

cx_Freeze.setup(
    name="Parser",
    options={"build_exe": {"packages": ["tkinter", "requests", "pyppeteer", "PyQt5", "configparser"],
                           "include_files": ["downarrow.svg", "uparrow.svg", "config.ini", "style.css", "icon.ico",
                                             "main_icon.ico"
                                             ]}},
    version="1.0",
    description="kad.arbitr.ru parser with gui for simple use",
    executables=executables,
    encoding='utf-8'
)
