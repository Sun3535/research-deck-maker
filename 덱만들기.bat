@echo off
chcp 65001 >nul
rem 덱설계.json(또는 프로젝트 폴더)을 이 파일 위에 끌어다 놓으면 결과 전부를 만든다
cd /d "%~dp0"
if not exist ".venv\ready.txt" (
  echo 처음이거나 설치가 덜 끝났어요. 준비하기.bat 을 먼저 눌러 주세요.
  pause & exit /b 1
)
set "TARGET=%~1"
if "%TARGET%"=="" set /p "TARGET=덱설계.json 경로(또는 프로젝트 이름)를 넣어 주세요: "
set "TARGET=%TARGET:"=%"
".venv\Scripts\python.exe" 덱만들기.py "%TARGET%" %2 %3 %4 %5 %6 %7 %8 %9
pause
