py export_ui.py
if errorlevel 1	goto ERROR
pause

:ERROR
python export_ui.py
pause