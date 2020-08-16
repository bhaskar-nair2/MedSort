# MedSort
A small python TKinter program to manage medical indents

## Components
MedSort.py -> Main window GUI
Searcher.py -> Algo for searcher

SearchSetup.py -> Refresh DB GUI
SearchFileMaker.py -> Algo for DB Maker


## Helpful Commands
### Build exe for windows
*Install PyInstaller*
pip3 install pyinstaller

*Create .spec file **!Required***
pyinstaller -F MedSort.py 

*Make the .exe file*
docker run -v "$(pwd):/src/" cdrx/pyinstaller-windows

https://github.com/pyinstaller/pyinstaller/issues/2613#issuecomment-302298224
