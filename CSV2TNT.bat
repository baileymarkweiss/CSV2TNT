@echo off
REM This batch file runs the CSV2TNT python code. Read the manual for more information.

REM installs required packages
pip install -r requirements.txt

REM minimises cmd window
if not "%Minimized%"=="" goto :Minimized
set Minimized=True
start /min cmd /C "%~dpnx0"
goto :EOF
:Minimized

REM Runs python code
pythonw.exe C:\Users\baile\OneDrive\Python\CSV2TNT.py
