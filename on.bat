
@echo off
cd /d "%~dp0"
python downloadCSV.py
python createExcel.py
python sortNaber.py
python redmineDataTransfer.py
python copyingShapes.py
pause
