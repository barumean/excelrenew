@echo off
REM ===========================================================
REM  Windows 단일 실행 파일(.exe) 빌드 스크립트
REM  사전 준비:  pip install pyinstaller
REM  실행:       build_exe.bat
REM  결과물:     dist\엑셀정리기.exe
REM ===========================================================
chcp 65001 >nul

pyinstaller --noconfirm --clean ^
    --onefile ^
    --windowed ^
    --name "엑셀정리기" ^
    gui.py

echo.
echo 빌드 완료. dist\엑셀정리기.exe 를 확인하세요.
pause
