@echo off
cd /d "%~dp0"

:: Avvia con pythonw (senza console) se disponibile, altrimenti con start /b
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw main.py
) else (
    start "" python main.py
)

:: Chiude immediatamente la finestra del batch
exit
