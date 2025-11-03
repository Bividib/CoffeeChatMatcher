@echo off
setlocal

echo ☕ Starting the KBC Coffee Pair setup...
echo ----------------------------------------

:: --- 2. Find the script's own directory ---
:: This changes the directory to the batch file's location
cd /d "%~dp0"

echo Running in project directory: %cd%

:: --- 3. Check for Python 3 ---
:: First, check for the 'py' launcher (modern standard)
py -3 --version > nul 2>&1
if not errorlevel 1 (
    echo ✅ Python 3 found (using 'py' launcher).
    set "PYTHON_EXE=py -3"
    goto :setup
)

:: If 'py' not found, check for 'python'
python --version > nul 2>&1
if not errorlevel 1 (
    :: Found 'python', now check if it's version 3
    for /f "tokens=1,2 delims= ." %%a in ('python --version 2^>&1') do (
        if "%%b" equ "3" (
            echo ✅ Python 3 found (using 'python').
            set "PYTHON_EXE=python"
            goto :setup
        )
    )
    echo ❌ Error: 'python' was found, but it is not Python 3.
    echo Please install Python 3 from python.org.
    pause
    exit /b 1
)

:: If neither is found
echo ❌ Error: Python 3 is not installed or not in your PATH.
echo Please install Python 3 from python.org to continue.
pause
exit /b 1


:setup
:: --- 4. Create requirements.txt (silently) ---
(
    echo pandas
    echo openpyxl
) > requirements.txt

:: --- 5. Create Virtual Environment and Install (silently) ---
echo Setting up environment and installing libraries... (This may take a moment)

:: Create venv and redirect all output to nul
%PYTHON_EXE% -m venv venv > nul 2>&1

:: Activate the virtual environment. Use 'call'
call "venv\Scripts\activate.bat" > nul 2>&1

:: Install dependencies and redirect all output to nul
pip install -r requirements.txt > nul 2>&1

echo ✅ Environment is ready.

:: --- 6. Run the Python Script ---
echo ----------------------------------------
echo 🚀 Running the coffee pair script...
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
echo 🎉 Pairing complete!
echo Check the output above.

:: Pause at the end so the user can read the output
pause
