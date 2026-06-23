@echo off
REM ===========================================================
REM  Windows 단일 실행 파일(.exe) 빌드 스크립트
REM  사전 준비:  pip install pyinstaller
REM             (드래그 앤 드롭까지 원하면)  pip install tkinterdnd2
REM  실행:       build_exe.bat  (더블클릭 또는 명령창에서)
REM  결과물:     dist\ExcelRenew.exe
REM ===========================================================
chcp 65001 >nul

REM  이 배치 파일이 있는 폴더로 이동(어디서 실행하든 gui.py 를 찾도록)
cd /d "%~dp0"

REM  PyInstaller 설치 여부 확인
where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [오류] PyInstaller 가 설치되어 있지 않습니다.
    echo        먼저 아래 명령으로 설치하세요:
    echo            pip install pyinstaller
    echo.
    pause
    exit /b 1
)

REM  tkinterdnd2 가 설치돼 있으면 그 바이너리(tkdnd)까지 포함(드래그 앤 드롭).
REM  없으면 드래그 앤 드롭 없이(버튼 방식으로) 빌드한다.
set "DND_OPT="
python -c "import tkinterdnd2" >nul 2>&1 && set "DND_OPT=--collect-all tkinterdnd2"
if defined DND_OPT (
    echo [정보] tkinterdnd2 감지됨 - 드래그 앤 드롭 포함하여 빌드합니다.
) else (
    echo [정보] tkinterdnd2 없음 - 버튼 방식으로만 빌드합니다.
    echo        드래그 앤 드롭을 원하면:  pip install tkinterdnd2
)
echo.

REM  exe 이름은 ASCII(ExcelRenew)로 둬서 한글 코드페이지 문제를 피한다.
REM  (만든 뒤 파일명을 한글로 자유롭게 바꿔도 동작에는 영향 없음)
pyinstaller --noconfirm --clean ^
    --onefile ^
    --windowed ^
    --name "ExcelRenew" ^
    %DND_OPT% ^
    gui.py

if errorlevel 1 (
    echo.
    echo [오류] 빌드에 실패했습니다. 위 메시지를 확인하세요.
    pause
    exit /b 1
)

echo.
echo 빌드 완료. dist\ExcelRenew.exe 를 확인하세요.
pause
