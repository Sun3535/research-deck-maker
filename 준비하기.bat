@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [덱편집기 준비] 관리자 권한 없이 이 폴더 안에 파이썬 환경(.venv)을 만들어요.

set "PY="
if exist "C:\VideoEnv\Scripts\python.exe" set "PY=C:\VideoEnv\Scripts\python.exe"
if not defined PY (where py >nul 2>nul && set "PY=py -3")
if not defined PY set "PY=python"
%PY% --version >nul 2>nul || (echo 파이썬을 찾지 못했어요. python.org 에서 3.10 이상을 설치해 주세요. & pause & exit /b 1)
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" || (echo 파이썬 3.10 이상이 필요해요. 지금 버전: & %PY% --version & pause & exit /b 1)

if exist ".venv\ready.txt" del ".venv\ready.txt"
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv || (echo 가상환경을 만들지 못했어요. & pause & exit /b 1)
)
set "VPY=.venv\Scripts\python.exe"

rem 꾸러미 폴더(미리 받아 둔 설치 파일 묶음)가 있으면 인터넷 없이 설치해요.
if exist "꾸러미\" (
  "%VPY%" -m pip install --no-index --find-links "꾸러미" -r requirements.txt
) else (
  rem 회사 인증서를 윈도우 인증서 저장소로 확인해요(인증서 검사를 끄지 않아요).
  rem 오래된 pip 에서도 되도록 truststore 를 앱에 들어 있는 파일(준비물)로 먼저 깔아요.
  "%VPY%" -m pip install --no-index --find-links "준비물" truststore >nul 2>nul
  "%VPY%" -m pip install --use-feature=truststore -r requirements.txt || "%VPY%" -m pip install -r requirements.txt
)
if errorlevel 1 (
  echo 꾸러미를 설치하지 못했어요. 회사망에서 막혔다면 IT 담당자에게 README 의 "회사 PC에서" 부분을 보여 주세요.
  pause & exit /b 1
)
"%VPY%" -c "import pptx, PIL, pypdfium2, openpyxl" || (echo 설치가 덜 끝났어요. 준비하기.bat 을 다시 눌러 주세요. & pause & exit /b 1)
echo ok> ".venv\ready.txt"
echo 준비 끝! 이제 켜기.bat 을 눌러요.
pause
