@echo off
REM Run the Python script to select the directory
echo Running Python script to select the directory...
for /f "delims=" %%i in ('python "%~dp0select_directory.py"') do set INSTALL_DIR=%%i

REM Check if the directory was selected
if "%INSTALL_DIR%"=="" (
    echo No directory selected!
    pause
    exit /b 1
)

REM Change to the selected directory
cd /d "%INSTALL_DIR%"

REM Copy template.xlsx and other necessary files to the selected directory
echo Copying files to the selected directory...
copy "%~dp0template.xlsx" "%INSTALL_DIR%"
copy "%~dp0main.py" "%INSTALL_DIR%"
copy "%~dp0run.bat" "%INSTALL_DIR%"
copy "%~dp0logging_setup.py" "%INSTALL_DIR%"
copy "%~dp0file_operations.py" "%INSTALL_DIR%"
copy "%~dp0extract_quota_usage.py" "%INSTALL_DIR%"
copy "%~dp0select_directory.py" "%INSTALL_DIR%"

REM Copy as_built_doc-v2.35.0 directory to the selected directory
xcopy "%~dp0as_built_doc-v2.35.0" "%INSTALL_DIR%\as_built_doc-v2.35.0" /E /I /Y

REM Create a virtual environment named 'venv'
echo Creating virtual environment...
python -m venv venv

REM Activate the virtual environment
echo Activating virtual environment...
call venv\Scripts\activate

REM Upgrade pip
python -m pip install --upgrade pip

REM Install the required modules
echo Installing required modules...
pip install openpyxl

REM Deactivate the virtual environment
echo Deactivating virtual environment...
deactivate

echo Virtual environment setup complete. Required modules installed.
echo Files copied to %INSTALL_DIR%.
pause