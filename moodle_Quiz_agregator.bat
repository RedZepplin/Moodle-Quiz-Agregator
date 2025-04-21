@echo off
REM Get the directory where the batch file is located
set SCRIPT_DIR=%~dp0
call conda activate moodle-quiz-agregator

REM Construct the full path to the Python script
set PYTHON_SCRIPT="%SCRIPT_DIR%moodle_quiz_agregator.py"

REM Check if a folder path was provided as the first argument (%1)
IF "%~1"=="" (
    echo No folder specified, using default 'Files' subfolder...
    python %PYTHON_SCRIPT%
) ELSE (
    echo Using specified folder: %1
    REM Pass the first argument (%1) and any subsequent arguments (%*) to the script
    python %PYTHON_SCRIPT% %*
)

echo.
pause
