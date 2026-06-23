@echo off
REM ===========================================================
REM  Windows 단일 실행 파일(.exe) 빌드 스크립트
REM  사전 준비:  pip install pyinstaller
REM             (드래그 앤 드롭까지 원하면)  pip install tkinterdnd2
REM  실행:       build_exe.bat
REM  결과물:     dist\엑셀정리기.exe
REM ===========================================================
chcp 65001 >nul

REM  tkinterdnd2 가 설치돼 있으면 그 바이너리(tkdnd)까지 포함시킨다.
REM  없으면 드래그 앤 드롭 없이(버튼 방식으로) 빌드한다.
set "DND_OPT="
python -c "import tkinterdnd2" >nul 2>&1 && set "DND_OPT=--collect-all tkinterdnd2"

pyinstaller --noconfirm --clean ^
    --onefile ^
    --windowed ^
    --name "엑셀정리기" ^
    %DND_OPT% ^
    gui.py

echo.
echo 빌드 완료. dist\엑셀정리기.exe 를 확인하세요.
pause
