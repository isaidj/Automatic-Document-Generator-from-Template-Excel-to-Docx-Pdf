# import sys
# import os
from cx_Freeze import setup, Executable

# ADD FILES TO INCLUDE HERE
files = ["icon.ico"]

# TARGET
target = Executable(script="main.py", base="Win32GUI", icon="icon.ico")

# Setup cx_Freeze
setup(
    name="ExcelToDocxTemplate",
    version="1.0",
    description="Excel values to docx template",
    author="Isai hernandez",
    options={
        "build_exe": {
            "include_files": files,
        }
    },
    executables=[target],
    target_name="ExcelToDocxTemplate.exe",
)
