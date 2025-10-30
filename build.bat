@echo off
setlocal

REM ========================
REM  설정
REM ========================
set APP_NAME=근태 자동정리 프로그램

REM ========================
REM  기존 빌드 제거
REM ========================
echo ===== 기존 빌드 제거 =====
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist %APP_NAME%.spec del %APP_NAME%.spec
echo ========================

REM ========================
REM  Python 가상환경에서 의존성 설치
REM ========================
echo ===== Python 의존성 설치 =====
python -m pip install -r requirements.txt
echo ========================

REM ========================
REM  PyInstaller 빌드
REM ========================
echo ===== PyInstaller 빌드 =====
pyinstaller --noconfirm ^
 --onefile ^
 --windowed ^
 --add-data "template.xlsx;." ^
 --name "%APP_NAME%" ^
 attendance_gui.py
echo ========================

echo.
echo ✅ 빌드 완료: dist\%APP_NAME%.exe
pause
