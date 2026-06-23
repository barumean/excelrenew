# -*- coding: utf-8 -*-
"""
excel_cleaner 모듈의 동작을 검증하는 단위 테스트.

표준 라이브러리만 사용하므로 다음 명령으로 바로 실행할 수 있습니다.
    python -m unittest test_excel_cleaner -v
"""

import os
import tempfile
import unittest
import zipfile

from excel_cleaner import clean_workbook

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>'
    "</Types>"
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    "</Relationships>"
)
WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/>'
    '<sheet name="Sheet2" sheetId="2" state="hidden" r:id="rId2"/>'
    '<sheet name="Sheet3" sheetId="3" state="veryHidden" r:id="rId4"/></sheets>'
    '<externalReferences><externalReference r:id="rId3"/></externalReferences>'
    '<definedNames>'
    '<definedName name="MyName">Sheet1!$A$1</definedName>'
    '<definedName name="Hidden_Name" hidden="1">Sheet1!$B$1</definedName>'
    '<definedName name="Broken">#REF!</definedName></definedNames></workbook>'
)
WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
    "</Relationships>"
)
SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<sheetData><row r="1"><c r="A1" t="str"><v>hello</v></c></row></sheetData></worksheet>'
)
EXT_LINK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<externalBook/></externalLink>'
)


def _make_dirty_xlsx(path: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/worksheets/sheet1.xml", SHEET)
        z.writestr("xl/externalLinks/externalLink1.xml", EXT_LINK)
        z.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", ROOT_RELS)


class TestCleanWorkbook(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.src = os.path.join(self.tmp, "dirty.xlsx")
        _make_dirty_xlsx(self.src)

    def _read(self, path, member):
        with zipfile.ZipFile(path) as z:
            return z.read(member).decode("utf-8")

    def test_counts(self):
        r = clean_workbook(self.src)
        self.assertTrue(r.ok, r.message)
        self.assertEqual(r.names_removed, 3)
        self.assertEqual(r.external_links_removed, 1)
        self.assertEqual(r.sheets_unhidden, 2)  # hidden + veryHidden

    def test_workbook_cleaned(self):
        r = clean_workbook(self.src)
        wb = self._read(r.dst_path, "xl/workbook.xml")
        self.assertNotIn("definedName", wb)
        self.assertNotIn("externalReference", wb)
        self.assertNotIn('state="hidden"', wb)
        self.assertNotIn("veryHidden", wb)

    def test_external_files_and_rels_removed(self):
        r = clean_workbook(self.src)
        with zipfile.ZipFile(r.dst_path) as z:
            names = z.namelist()
        self.assertFalse(any(n.startswith("xl/externalLinks/") for n in names))
        rels = self._read(r.dst_path, "xl/_rels/workbook.xml.rels")
        self.assertNotIn("externalLink", rels)
        ct = self._read(r.dst_path, "[Content_Types].xml")
        self.assertNotIn("externalLink", ct)

    def test_worksheet_data_preserved(self):
        r = clean_workbook(self.src)
        sheet = self._read(r.dst_path, "xl/worksheets/sheet1.xml")
        self.assertIn("hello", sheet)

    def test_default_output_is_new_file(self):
        r = clean_workbook(self.src)
        self.assertTrue(r.dst_path.endswith("_정리됨.xlsx"))
        self.assertTrue(os.path.exists(self.src))  # 원본 보존

    def test_overwrite_creates_backup(self):
        r = clean_workbook(self.src, overwrite=True, backup=True)
        self.assertEqual(os.path.abspath(r.dst_path), os.path.abspath(self.src))
        self.assertTrue(os.path.exists(self.src + ".bak"))

    def test_selective_options(self):
        # 이름만 삭제, 링크/숨김은 유지
        r = clean_workbook(
            self.src, remove_external_links=False, unhide_sheets=False
        )
        wb = self._read(r.dst_path, "xl/workbook.xml")
        self.assertNotIn("definedName", wb)
        self.assertIn("externalReference", wb)
        self.assertIn('state="hidden"', wb)
        self.assertEqual(r.external_links_removed, 0)
        self.assertEqual(r.sheets_unhidden, 0)

    def test_unsupported_extension(self):
        bad = os.path.join(self.tmp, "old.xls")
        with open(bad, "wb") as f:
            f.write(b"not a zip")
        r = clean_workbook(bad)
        self.assertFalse(r.ok)
        self.assertIn("지원하지 않는", r.message)

    def test_corrupt_file(self):
        bad = os.path.join(self.tmp, "broken.xlsx")
        with open(bad, "wb") as f:
            f.write(b"not a zip at all")
        r = clean_workbook(bad)
        self.assertFalse(r.ok)


if __name__ == "__main__":
    unittest.main(verbosity=2)
