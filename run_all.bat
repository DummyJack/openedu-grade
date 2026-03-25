@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_CMD="
if exist "%~dp0venv\Scripts\python.exe" (
    set "PYTHON_CMD=%~dp0venv\Scripts\python.exe"
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=py -3"
    ) else (
        where python >nul 2>nul
        if %errorlevel%==0 (
            set "PYTHON_CMD=python"
        ) else (
            echo [ERROR] Python not found.
            exit /b 1
        )
    )
)

echo [INFO] Using Python: %PYTHON_CMD%
echo [INFO] Project Dir: %cd%
echo [INFO] Cleaning old output files...

if not exist "scores" mkdir "scores"
if not exist "errors" mkdir "errors"
if exist "openedu_all_videos.xlsx" del /f /q "openedu_all_videos.xlsx"
if exist "scores\*.xlsx" del /f /q "scores\*.xlsx"
if exist "errors\*.xlsx" del /f /q "errors\*.xlsx"

call :run_step "student.py" "Please rerun student.py. If some users failed, check errors/video_errors.xlsx."
if errorlevel 1 exit /b %errorlevel%

if not exist "openedu_all_videos.xlsx" (
    echo [FAIL] openedu_all_videos.xlsx was not generated.
    echo [RETRY] Please rerun student.py first.
    exit /b 1
)

call :run_step "video_grade.py" "Please rerun video_grade.py (it will ask chapter range again)."
if errorlevel 1 exit /b %errorlevel%

call :run_step "grade.py" "Please rerun grade.py. If partial failures exist, choose option 3 to reprocess errors/grade_errors.xlsx."
if errorlevel 1 exit /b %errorlevel%

echo [DONE] All scripts completed successfully.
exit /b 0

:run_step
set "SCRIPT_NAME=%~1"
set "RETRY_HINT=%~2"

echo.
echo ==================================================
echo [RUN] %SCRIPT_NAME%
echo ==================================================

%PYTHON_CMD% %SCRIPT_NAME%
if errorlevel 1 (
    echo.
    echo [FAIL] %SCRIPT_NAME% failed.
    echo [RETRY] %RETRY_HINT%
    exit /b 1
)
exit /b 0
