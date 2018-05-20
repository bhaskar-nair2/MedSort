import cx_Freeze
import sys
import os
import openpyxl
import itertools
import re


os.environ['TCL_LIBRARY'] = r'C:\Program Files (x86)\Python3\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Program Files (x86)\Python3\tcl\tk8.6'

base = None

if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [cx_Freeze.Executable("front.py", base=base)]

options = {'build_exe': {'packages':["tkinter","openpyxl","itertools",'re']}}

cx_Freeze.setup(
    name = "Medical Sorter Software",
    options = options,
    version = "0.1",
    description = 'Moms Software',
    executables = executables
)
