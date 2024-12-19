REM filepath: /g:/My Drive/Python Scripts/OIT Billing/run.bat
@echo off

REM Check if the virtual environment exists
if not exist "venv\Scripts\activate" (
    echo Virtual environment not found. Please run setup.bat first.
    pause
    exit /b 1
)

REM Activate the virtual environment
call venv\Scripts\activate

REM Run the Python script
python "main.py"
if %errorlevel% neq 0 (
    echo Error: Failed to execute the Python script.
    call venv\Scripts\deactivate
    pause
    exit /b 1
)

REM Deactivate the virtual environment
call venv\Scripts\deactivate

echo Script executed successfully.
pause