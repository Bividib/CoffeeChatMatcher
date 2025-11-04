@echo off
setlocal EnableDelayedExpansion

echo Starting the KBC Coffee Pair setup...
echo ----------------------------------------

:: --- 2. Find the script's own directory ---
:: This changes the directory to the batch file's location
cd /d "%~dp0"

echo Running in project directory: %cd%

:: --- 3. Check for Python 3 ---
echo Checking for 'py' launcher...
:: We use 'findstr' to check the output of '--version' for "Python 3"
:: 2>&1 redirects error output (like "'py' is not recognized") to standard output
:: | findstr ... searches that output
:: > nul silences the output of findstr
for /f "delims=" %%i in ('py -3 --version 2^>^&1') do (
    echo %%i | findstr /i "Python 3" >nul && (
        echo Python 3 found (py launcher)
        set "PYTHON_EXE=py -3"
        goto :setup
    )
)

for /f "delims=" %%i in ('python --version 2^>^&1') do (
    echo %%i | findstr /i "Python 3" >nul && (
        echo Python 3 found (python)
        set "PYTHON_EXE=python"
        goto :setup
    )
)

echo Neither 'py' nor 'python' was found, or neither is Python 3.
echo [ERROR] Python 3 is not installed or not in your PATH.
echo Please install Python 3 from python.org to continue.
pause
exit /b 1


:setup


:: --- 5. Create Virtual Environment and Install (silently) ---
echo Setting up environment and installing libraries... (This may take a moment)

:: Check if the venv activation script already exists
if not exist "venv\Scripts\activate.bat" (
echo Creating new virtual environment...
:: Create venv and redirect all output to nul
%PYTHON_EXE% -m venv venv > nul 2>&1
) else (
echo Found existing virtual environment.
)

:: Activate the virtual environment. Use 'call'
echo Activating venv...
call "venv\Scripts\activate.bat" > nul 2>&1

:: Install/verify dependencies. This is fast and ensures requirements are always met.
echo Installing requirements...
pip install -r requirements.txt > nul 2>&1

echo [OK] Environment is ready.

:: --- 6. Run the Python Script ---
echo ----------------------------------------
echo Running the coffee pair script...
echo.

set "PYTHON_SCRIPT=create_coffee_pairs.py"

:: Check if the user provided an argument (a file path)
:: %~1 removes quotes from the path, "%~1" re-adds them safely
if not "%~1"=="" (
echo Using provided Excel file: %1
:: After activation, 'python' should point to the venv
python "%PYTHON_SCRIPT%" "%~1"
) else (
echo No Excel file provided. Python script will search for one...
python "%PYTHON_SCRIPT%"
)

echo.
echo ----------------------------------------
echo Pairing complete!
echo Check the output above.

:: Pause at the end so the user can read the output
pause