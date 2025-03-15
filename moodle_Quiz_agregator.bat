@echo off
echo Activating the moodle-quiz-agregator environment...
call conda activate moodle-quiz-agregator
if %errorlevel% neq 0 (
    echo Failed to activate environment. Please check your conda installation.
    pause
    exit /b 1
)

echo Running the Moodle Quiz Aggregator script...
python moodle_quiz_agregator.py

if %errorlevel% neq 0 (
    echo An error occurred while running the script.
    pause
    exit /b 1
)

echo Script completed successfully.
pause