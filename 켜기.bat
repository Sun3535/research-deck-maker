@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\ready.txt" (
  echo 처음이거나 설치가 덜 끝났어요. 준비하기.bat 을 먼저 눌러 주세요.
  pause & exit /b 1
)
".venv\Scripts\python.exe" 서버.py %*
pause
