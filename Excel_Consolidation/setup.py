'''Setup script for cx_Freeze to create an executable from the Excel_Consolidation script'''

import sys
from cx_Freeze import setup, Executable

# Add any packages or modules that need to be included
build_exe_options = {
    "packages": ["os", "sys"],  # Include any additional packages here
    "excludes": [],  # Exclude unnecessary packages
    "include_files": ["Resources"],  # Include any additional files, like data files or icons
}

# Determine the base of the executable
BASE = None
if sys.platform == "win32":
    BASE = "Win32GUI"  # Use "Win32GUI" for GUI applications (no console window)

# Define the executables
executables = [
    Executable(
        "main.py",
        base=BASE,
        target_name="Excel Consolidator.exe",
        shortcut_name="Excel Consolidator",
        shortcut_dir="DesktopFolder"
        # icon="path_to_icon.ico"
    )
]

bdist_msi_options = {
    "upgrade_code": "{a1c4ad49-ea3f-4f5f-bc92-3ebc618a39fc}",
    "add_to_path": False,
    "initial_target_dir": r"C:\Genpact LDT\Excel Consolidator",
}

# Setup configuration
setup(
    name="Excel Consolidator",
    version="1.0",
    description="Program to consolidate Excel files",
    options={
        "build_exe": build_exe_options,
        "bdist_msi": bdist_msi_options
    },
    executables=executables,
)
