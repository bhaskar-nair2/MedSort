from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.

buildOptions = dict(packages= ['scrollFrame','searchSetup',"tkinter","threading","sqlite3","openpyxl","re"],
                                include_msvcr= True,
                                include_files=['C:\\Program Files\\Python36\\DLLs\\tcl86t.dll', 'C:\\Program Files\\Python36\\DLLs\\tk86t.dll','S:\\Documents\\Python\\Medical-Sorter\\trash.txt'])

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('MedSort.py', base=base)
]

setup(name='MedSort',
      version = '2.0',
      description = 'Medical Indent management',
      options = dict(build_exe = buildOptions),
      executables = executables)
