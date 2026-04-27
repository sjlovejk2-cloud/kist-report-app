@echo off

setlocal EnableExtensions



cd /d "%~dp0"

title KIST Report App Server



echo [1/5] Working folder: %CD%



set "VENV_DIR=.venv-win"

set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"

set "PY_LAUNCHER=py"

set "FALLBACK_PY=C:\Python314\python.exe"

set "PYTHONHOME="
set "PYTHONPATH="

if exist "%FALLBACK_PY%" (

    set "PY_LAUNCHER=%FALLBACK_PY%"

) else (

    where %PY_LAUNCHER% >nul 2>nul

    if errorlevel 1 (

        echo [ERROR] Python launcher not found.

        echo         Checked: py and %FALLBACK_PY%

        pause

        exit /b 1

    )

)



echo [2/5] Python launcher ready: %PY_LAUNCHER%



if not exist "%PYTHON_EXE%" (

    echo [3/5] Creating Windows venv: %VENV_DIR%

    %PY_LAUNCHER% -m venv "%VENV_DIR%"

    if errorlevel 1 (

        echo [ERROR] Failed to create virtual environment.

        pause

        exit /b 1

    )



    echo [4/5] Installing required packages. Please wait...

    "%PYTHON_EXE%" -m pip install --upgrade pip

    if errorlevel 1 goto :pip_fail

    "%PYTHON_EXE%" -m pip install -r requirements.txt chromadb pillow pyyaml

    if errorlevel 1 goto :pip_fail

) else (

    echo [3/5] Using existing Windows venv: %VENV_DIR%

)



echo [5/5] Starting Streamlit app...

echo         Open http://localhost:8501 if the browser does not open automatically.

echo.

"%PYTHON_EXE%" -m streamlit run app.py

set "EXIT_CODE=%ERRORLEVEL%"

echo.

echo [EXIT] Streamlit exit code: %EXIT_CODE%

pause

exit /b %EXIT_CODE%



:pip_fail

echo [ERROR] Package installation failed.

pause

exit /b 1

