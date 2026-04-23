@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo 현재 변경사항을 확인합니다...
git status
echo.

set /p MSG=커밋 메시지를 입력하세요: 
if "%MSG%"=="" (
    echo [오류] 커밋 메시지가 비어 있습니다.
    pause
    exit /b 1
)

git add .
git commit -m "%MSG%"
if errorlevel 1 (
    echo.
    echo [안내] 커밋할 변경사항이 없거나 커밋에 실패했습니다.
    pause
    exit /b 1
)

git push
if errorlevel 1 (
    echo.
    echo [오류] git push에 실패했습니다. 인증 상태와 네트워크를 확인하세요.
    pause
    exit /b 1
)

echo.
echo 완료되었습니다.
pause
