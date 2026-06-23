# -*- coding: utf-8 -*-
"""
excel_cleaner.py
================
엑셀(.xlsx / .xlsm / .xltx / .xltm) 파일에서 오류를 유발하는 요소를
한 번에 정리해 "정상 파일"로 되돌려 주는 핵심 로직 모듈입니다.

정리 대상
---------
1. 정의된 이름(Defined Names) 전부 삭제
   - 숨겨진 이름, #REF! 로 깨진 이름, 전역/시트범위 이름 모두 포함.
   - (VBA 의 Names.Delete 와 동일한 효과를 파일 레벨에서 수행)
2. 외부 링크/연결(External Links) 제거
   - 다른 통합문서를 참조하는 외부 링크 정의와 관계(rels)를 제거.
3. 숨겨진 시트 다시 표시(Unhide)
   - hidden / veryHidden 상태의 시트를 모두 visible 로 복구.

설계상의 핵심
-------------
.xlsx 계열 파일은 사실 여러 XML 파일을 담은 ZIP 압축 파일입니다.
openpyxl 같은 라이브러리로 다시 저장하면 차트·피벗테이블·매크로 등이
손실될 수 있으므로, 여기서는 ZIP 안의 "문제되는 부분"만 외과적으로
수정하고 나머지는 원본 그대로 복사합니다. 따라서 외부 의존성이 전혀
없고(표준 라이브러리만 사용), 원본 콘텐츠 손실 위험이 최소화됩니다.
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field

# ZIP 기반(=XML 묶음)으로 처리 가능한 확장자. .xls(구형 이진 포맷)는 제외.
SUPPORTED_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

WORKBOOK_XML = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
EXTERNAL_LINKS_PREFIX = "xl/externalLinks/"
CALC_CHAIN = "xl/calcChain.xml"

# 엑셀이 내부적으로 "정의된 이름"으로 저장하는 인쇄 영역/제목.
# 이 이름들은 삭제 대상에서 제외할 수 있다(인쇄 영역 유지).
PRINT_BUILTIN_NAMES = ("_xlnm.Print_Area", "_xlnm.Print_Titles")


@dataclass
class CleanResult:
    """파일 한 개를 정리한 결과 보고서."""

    src_path: str
    dst_path: str = ""
    ok: bool = False
    names_removed: int = 0
    external_links_removed: int = 0
    sheets_unhidden: int = 0
    message: str = ""
    warnings: list = field(default_factory=list)

    def summary(self) -> str:
        if not self.ok:
            return f"[실패] {os.path.basename(self.src_path)} → {self.message}"
        parts = [
            f"정의된 이름 {self.names_removed}개 삭제",
            f"외부 링크 {self.external_links_removed}개 제거",
            f"숨겨진 시트 {self.sheets_unhidden}개 표시",
        ]
        return f"[완료] {os.path.basename(self.src_path)} → " + ", ".join(parts)


# ---------------------------------------------------------------------------
# workbook.xml 수정 헬퍼
# ---------------------------------------------------------------------------
def _strip_defined_names(xml: str, keep_print_areas: bool = True) -> tuple[str, int]:
    """
    <definedNames> 안의 개별 <definedName> 항목을 제거한다.

    keep_print_areas=True 이면 인쇄 영역/제목(_xlnm.Print_Area,
    _xlnm.Print_Titles)은 남겨 두고 나머지 이름만 삭제한다.
    삭제한 이름 개수를 함께 반환한다.
    """
    count = 0

    def _process_block(block_match: re.Match) -> str:
        nonlocal count
        block = block_match.group(0)

        kept_entries: list[str] = []

        def _each_name(entry_match: re.Match) -> str:
            nonlocal count
            entry = entry_match.group(0)
            name_attr = re.search(r'\bname="([^"]*)"', entry)
            name = name_attr.group(1) if name_attr else ""
            if keep_print_areas and name in PRINT_BUILTIN_NAMES:
                kept_entries.append(entry)  # 인쇄 영역은 보존
            else:
                count += 1  # 삭제
            return ""

        # 블록 내부의 개별 <definedName> 항목을 순회
        re.sub(
            r"<definedName\b[^>]*?(?:/>|>.*?</definedName>)",
            _each_name,
            block,
            flags=re.DOTALL,
        )

        if kept_entries:
            return "<definedNames>" + "".join(kept_entries) + "</definedNames>"
        return ""  # 남길 이름이 없으면 블록 자체 제거

    # <definedNames> ... </definedNames>
    xml = re.sub(
        r"<definedNames\b.*?</definedNames>",
        _process_block,
        xml,
        flags=re.DOTALL,
    )
    # 비어있는 self-closing 형태도 제거 (<definedNames/>)
    xml = re.sub(r"<definedNames\b[^>]*/>", "", xml)
    return xml, count


def _strip_external_references(xml: str) -> str:
    """workbook.xml 내 <externalReferences> 블록 제거."""
    xml = re.sub(
        r"<externalReferences\b.*?</externalReferences>",
        "",
        xml,
        flags=re.DOTALL,
    )
    xml = re.sub(r"<externalReferences\b[^>]*/>", "", xml)
    return xml


def _unhide_sheets(xml: str) -> tuple[str, int]:
    """<sheet> 요소의 state="hidden"/"veryHidden" 속성을 제거(=visible)."""
    matches = re.findall(r'\sstate="(?:hidden|veryHidden)"', xml)
    xml = re.sub(r'\sstate="(?:hidden|veryHidden)"', "", xml)
    return xml, len(matches)


# ---------------------------------------------------------------------------
# 관계(rels) / 콘텐츠 형식 수정 헬퍼
# ---------------------------------------------------------------------------
def _strip_external_link_rels(rels_xml: str) -> tuple[str, int]:
    """workbook.xml.rels 에서 externalLink 관계 항목을 제거."""
    pattern = r"<Relationship\b[^>]*externalLink[^>]*/>"
    count = len(re.findall(pattern, rels_xml))
    rels_xml = re.sub(pattern, "", rels_xml)
    return rels_xml, count


def _strip_external_link_content_types(ct_xml: str) -> str:
    """[Content_Types].xml 에서 externalLink Override 항목 제거."""
    return re.sub(r"<Override\b[^>]*externalLink[^>]*/>", "", ct_xml)


def _strip_calc_chain_rels(rels_xml: str) -> str:
    """workbook.xml.rels 에서 calcChain 관계 항목을 제거."""
    return re.sub(r"<Relationship\b[^>]*calcChain[^>]*/>", "", rels_xml)


def _strip_calc_chain_content_types(ct_xml: str) -> str:
    """[Content_Types].xml 에서 calcChain Override 항목 제거."""
    return re.sub(r"<Override\b[^>]*calcChain[^>]*/>", "", ct_xml)


# ---------------------------------------------------------------------------
# 메인 처리 함수
# ---------------------------------------------------------------------------
def clean_workbook(
    src_path: str,
    dst_path: str | None = None,
    *,
    delete_names: bool = True,
    remove_external_links: bool = True,
    unhide_sheets: bool = True,
    keep_print_areas: bool = True,
    drop_calc_chain: bool = True,
    overwrite: bool = False,
    backup: bool = True,
) -> CleanResult:
    """
    엑셀 파일 한 개를 정리한다.

    Parameters
    ----------
    src_path : 원본 파일 경로
    dst_path : 저장 경로. None 이면 자동 결정.
               - overwrite=False: 같은 폴더에 "<이름>_정리됨.<확장자>"
               - overwrite=True : 원본 경로 (backup=True 면 .bak 백업 생성)
    delete_names, remove_external_links, unhide_sheets : 수행할 작업 선택
    keep_print_areas : 인쇄 영역/제목(_xlnm.Print_Area/_xlnm.Print_Titles)은
                       삭제하지 않고 보존(기본값 True)
    drop_calc_chain : 계산 순서 캐시(xl/calcChain.xml)를 제거. 구조가 바뀌면
                      엑셀이 "제거된 레코드(계산 속성)" 경고를 띄우는데, 이
                      파일을 미리 지워 두면 경고가 사라지고 엑셀이 자동
                      재생성한다(기본값 True).
    overwrite : 원본을 덮어쓸지 여부
    backup : overwrite=True 일 때 원본 백업(.bak) 생성 여부
    """
    result = CleanResult(src_path=src_path)

    ext = os.path.splitext(src_path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        result.message = (
            f"지원하지 않는 형식({ext}). .xlsx/.xlsm/.xltx/.xltm 만 가능합니다. "
            "구형 .xls 파일은 먼저 .xlsx 로 저장해 주세요."
        )
        return result

    if not zipfile.is_zipfile(src_path):
        result.message = "올바른 엑셀 파일이 아니거나 손상되어 ZIP 으로 열 수 없습니다."
        return result

    # 저장 경로 결정
    if dst_path is None:
        if overwrite:
            dst_path = src_path
        else:
            root, e = os.path.splitext(src_path)
            dst_path = f"{root}_정리됨{e}"
    result.dst_path = dst_path

    # 원본을 메모리/임시로 처리하기 위해, 새 zip 은 임시 파일에 먼저 쓴다.
    tmp_path = dst_path + ".tmp_clean"

    try:
        with zipfile.ZipFile(src_path, "r") as zin:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    name = item.filename
                    data = zin.read(name)

                    # 1) 외부 링크 관련 파일은 통째로 제외(링크 파일 + 해당 _rels)
                    if remove_external_links and name.startswith(EXTERNAL_LINKS_PREFIX):
                        continue

                    # 1-2) 계산 순서 캐시는 제거(엑셀이 재생성 → 계산속성 경고 방지)
                    if drop_calc_chain and name == CALC_CHAIN:
                        continue

                    # 2) workbook.xml: 이름/외부참조 제거, 시트 숨김 해제
                    if name == WORKBOOK_XML:
                        xml = data.decode("utf-8")
                        if delete_names:
                            xml, n = _strip_defined_names(xml, keep_print_areas)
                            result.names_removed = n
                        if remove_external_links:
                            xml = _strip_external_references(xml)
                        if unhide_sheets:
                            xml, n = _unhide_sheets(xml)
                            result.sheets_unhidden = n
                        data = xml.encode("utf-8")

                    # 3) workbook.xml.rels: 외부 링크 + calcChain 관계 제거
                    #    (외부 링크는 연결 통합문서 1개당 관계 1개 = 정확한 링크 수)
                    elif name == WORKBOOK_RELS:
                        rels = data.decode("utf-8")
                        if remove_external_links:
                            rels, n = _strip_external_link_rels(rels)
                            result.external_links_removed = n
                        if drop_calc_chain:
                            rels = _strip_calc_chain_rels(rels)
                        data = rels.encode("utf-8")

                    # 4) [Content_Types].xml: 외부 링크 + calcChain Override 제거
                    elif name == CONTENT_TYPES:
                        ct = data.decode("utf-8")
                        if remove_external_links:
                            ct = _strip_external_link_content_types(ct)
                        if drop_calc_chain:
                            ct = _strip_calc_chain_content_types(ct)
                        data = ct.encode("utf-8")

                    # 원본 압축 방식/메타데이터를 유지하며 기록
                    zout.writestr(item, data)

    except Exception as exc:  # noqa: BLE001
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        result.message = f"처리 중 오류: {exc}"
        return result

    # 임시 파일을 최종 위치로 이동 (덮어쓰기 시 백업 처리)
    try:
        if overwrite and backup and os.path.abspath(dst_path) == os.path.abspath(src_path):
            bak = src_path + ".bak"
            if not os.path.exists(bak):
                shutil.copy2(src_path, bak)
            else:
                # 기존 .bak 이 있으면 번호를 붙여 보존
                i = 1
                while os.path.exists(f"{src_path}.bak{i}"):
                    i += 1
                shutil.copy2(src_path, f"{src_path}.bak{i}")
                result.warnings.append(f"백업: {os.path.basename(src_path)}.bak{i}")
        if os.path.exists(dst_path) and os.path.abspath(dst_path) == os.path.abspath(tmp_path):
            pass
        os.replace(tmp_path, dst_path)
    except Exception as exc:  # noqa: BLE001
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        result.message = f"파일 저장 오류: {exc}"
        return result

    result.ok = True
    result.message = "정상 처리되었습니다."
    return result


def clean_many(paths, **kwargs) -> list[CleanResult]:
    """여러 파일을 순서대로 정리하고 결과 리스트를 반환."""
    results = []
    for p in paths:
        results.append(clean_workbook(p, **kwargs))
    return results


# ---------------------------------------------------------------------------
# 명령줄 실행 지원: python excel_cleaner.py <파일들...>
# ---------------------------------------------------------------------------
def _main(argv) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="엑셀 파일의 정의된 이름/외부 링크/숨겨진 시트를 정리합니다."
    )
    parser.add_argument("files", nargs="+", help="정리할 엑셀 파일 경로(여러 개 가능)")
    parser.add_argument("--overwrite", action="store_true", help="원본 덮어쓰기(.bak 백업 생성)")
    parser.add_argument("--no-names", action="store_true", help="정의된 이름 삭제 안 함")
    parser.add_argument("--no-links", action="store_true", help="외부 링크 제거 안 함")
    parser.add_argument("--no-unhide", action="store_true", help="숨겨진 시트 표시 안 함")
    parser.add_argument(
        "--delete-print-areas",
        action="store_true",
        help="인쇄 영역(Print Area)도 함께 삭제 (기본은 보존)",
    )
    parser.add_argument(
        "--keep-calc-chain",
        action="store_true",
        help="계산 순서 캐시(calcChain.xml)를 남김 (기본은 제거)",
    )
    args = parser.parse_args(argv)

    results = clean_many(
        args.files,
        delete_names=not args.no_names,
        remove_external_links=not args.no_links,
        unhide_sheets=not args.no_unhide,
        keep_print_areas=not args.delete_print_areas,
        drop_calc_chain=not args.keep_calc_chain,
        overwrite=args.overwrite,
    )
    failed = 0
    for r in results:
        print(r.summary())
        if r.ok and r.dst_path != r.src_path:
            print(f"        저장 위치: {r.dst_path}")
        for w in r.warnings:
            print(f"        ⚠ {w}")
        if not r.ok:
            failed += 1
    return 1 if failed else 0


if __name__ == "__main__":
    import sys

    raise SystemExit(_main(sys.argv[1:]))
