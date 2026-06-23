# 엑셀 정리기 (Excel Renew)

엑셀 파일에서 오류를 자주 일으키는 요소들을 **한 번에 정리**해 정상 파일로
되돌려 주는 프로그램입니다. 예전에 VBA로 "정의된 이름 전부 삭제"를 하던
원리를 확장해, 파일 자체를 직접 정리합니다.

## 정리하는 항목

1. **정의된 이름(Defined Names) 전부 삭제**
   - 숨겨진 이름, `#REF!` 로 깨진 이름, 전역/시트범위 이름까지 모두 제거.
   - 아래 VBA 코드와 동일한 효과를 파일 레벨에서 수행합니다.
     ```vba
     Sub Delete_Names()
         Dim n As Name
         On Error Resume Next
         For Each n In ThisWorkbook.Names
             n.Visible = True
             n.Delete
         Next n
     End Sub
     ```
2. **외부 링크/연결 제거** — 다른 통합문서를 참조하는 외부 링크와 연결을 제거.
3. **숨겨진 시트 다시 표시** — `hidden` / `veryHidden` 상태의 시트를 모두 복구.

## 동작 원리 (왜 안전한가)

`.xlsx`·`.xlsm` 파일은 사실 여러 XML 파일을 담은 **ZIP 압축 파일**입니다.
이 프로그램은 `openpyxl` 같은 라이브러리로 파일을 다시 저장하지 않고,
ZIP 안에서 **문제되는 부분만 외과적으로 수정**합니다.

- `xl/workbook.xml` → 정의된 이름·외부참조 제거, 시트 숨김 속성 제거
- `xl/externalLinks/` → 외부 링크 정의 파일 제거
- `xl/_rels/workbook.xml.rels`, `[Content_Types].xml` → 관련 관계 항목 정리

덕분에 **차트·피벗테이블·매크로·서식 등 나머지 내용은 원본 그대로 보존**되며,
**외부 라이브러리가 전혀 필요 없습니다(파이썬 표준 라이브러리만 사용).**

> ⚠️ 구형 `.xls`(이진 포맷)는 지원하지 않습니다. 엑셀에서 `.xlsx`로 먼저
> 저장한 뒤 사용해 주세요.

## 사용법 (GUI)

파이썬이 설치된 환경에서:

```bash
python gui.py
```

1. **파일 추가…** 로 정리할 엑셀 파일을 선택(여러 개 가능)
2. **정리 항목** 에서 원하는 작업 선택 (기본: 3가지 모두)
3. **저장 방식** 선택
   - *새 파일로 저장* (기본/권장): 원본은 그대로 두고 `<이름>_정리됨.xlsx` 생성
   - *원본 덮어쓰기*: 자동으로 `.bak` 백업을 만든 뒤 원본을 교체
4. **정리 실행** 클릭

## 사용법 (명령줄)

GUI 없이 터미널에서도 쓸 수 있고, 여러 파일 일괄 처리에 편리합니다.

```bash
# 새 파일로 저장 (원본 보존)
python excel_cleaner.py "파일1.xlsx" "파일2.xlsm"

# 원본 덮어쓰기(.bak 백업 생성)
python excel_cleaner.py --overwrite "파일.xlsx"

# 특정 작업만 끄기
python excel_cleaner.py --no-links --no-unhide "파일.xlsx"
```

## Windows 실행 파일(.exe) 만들기

파이썬이 없는 PC에서도 더블클릭으로 쓰도록 단일 `.exe` 로 빌드할 수 있습니다.

```bat
pip install pyinstaller
build_exe.bat
```

`dist\엑셀정리기.exe` 가 생성됩니다. (Windows 에서 실행해야 합니다.)

## 테스트

```bash
python -m unittest test_excel_cleaner -v
```

## 파일 구성

| 파일 | 설명 |
| --- | --- |
| `excel_cleaner.py` | 핵심 정리 로직 + 명령줄 인터페이스 |
| `gui.py` | tkinter 기반 GUI |
| `test_excel_cleaner.py` | 단위 테스트 |
| `build_exe.bat` | Windows `.exe` 빌드 스크립트 |
