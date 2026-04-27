@echo off
setlocal EnableExtensions

cd /d "%~dp0"
title KIST Report App Server

echo [1/5] 작업 폴더: %CD%

set "VENV_DIR=.venv-win"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "PY_LAUNCHER=py"

where %PY_LAUNCHER% >nul 2>nul
if errorlevel 1 (
    echo [오류] Windows Python launcher(py)를 찾지 못했습니다.
    echo        Python이 설치되어 있는지 확인해주세요.
    echo        예: C:\Python314\python.exe 또는 py 명령 사용 가능 상태
    pause
    exit /b 1
)

echo [2/5] Python launcher 확인 완료

if not exist "%PYTHON_EXE%" (
    echo [3/5] 가상환경이 없어 새로 만듭니다: %VENV_DIR%
    %PY_LAUNCHER% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [오류] 가상환경 생성에 실패했습니다.
        echo        python -m venv 가 가능한지 확인해주세요.
        pause
        exit /b 1
    )

    echo [4/5] 첫 실행용 패키지를 설치합니다. 잠시 기다려주세요.
    "%PYTHON_EXE%" -m pip install --upgrade pip
    if errorlevel 1 goto :pip_fail

    "%PYTHON_EXE%" -m pip install -r requirements.txt chromadb pillow pyyaml
    if errorlevel 1 goto :pip_fail
) else (
    echo [3/5] 기존 가상환경 사용: %VENV_DIR%
)

echo [5/5] Streamlit 서버를 시작합니다...
echo        브라우저가 자동으로 안 열리면 아래 주소를 직접 여세요.
echo        http://localhost:8501
echo.
"%PYTHON_EXE%" -m streamlit run app.py
set "EXIT_CODE=%ERRORLEVEL%"
echo.
echo [종료] Streamlit 종료 코드: %EXIT_CODE%
pause
exit /b %EXIT_CODE%

:pip_fail
echo [오류] 패키지 설치에 실패했습니다.
echo        네트워크 또는 Python/pip 상태를 확인해주세요.
pause
exit /b 1
