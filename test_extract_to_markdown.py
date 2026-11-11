"""Integration checks using the installed LibreOffice and Python UNO bridge."""
import json
import os
import socket
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import extract_to_markdown as converter


class WriterExtractionTests(unittest.TestCase):
    def test_docx_conversion_preserves_text_table_order(self):
        def paragraph(text, level=0):
            return Mock(**{
                "supportsService.side_effect": lambda service: service.endswith(".Paragraph"),
                "getString.return_value": text,
                "getPropertyValue.return_value": level,
            })

        table = Mock()
        table.supportsService.side_effect = lambda service: service.endswith(".TextTable")
        values = {
            "A1": "항목", "B1": "설명", "A2": "병합 셀",
            "A3.1.1": "분할 셀 1", "A3.1.2": "분할 셀 2",
            "B3": "파이프 | 줄바꿈\n다음 줄",
        }
        table.getCellNames.return_value = tuple(values)
        table.getCellByName.side_effect = lambda name: Mock(**{"getString.return_value": values[name]})
        elements = iter([
            paragraph("문서 제목", 2), paragraph("표 앞 문단"), table,
            paragraph(""), paragraph("표 뒤 문단"),
        ])
        enumeration = Mock()
        enumeration.hasMoreElements.side_effect = [True] * 5 + [False]
        enumeration.nextElement.side_effect = lambda: next(elements)
        document = Mock()
        document.getText.return_value.createEnumeration.return_value = enumeration
        document.supportsService.return_value = True
        with tempfile.TemporaryDirectory(dir=converter.BASE_DIR) as temp:
            root = Path(temp)
            source = root / "sample.DOCX"
            source.touch()
            output = root / "output"
            with patch.object(converter, "INPUT_DIR", root / "input"), \
                    patch.object(converter, "OUTPUT_DIR", output), \
                    patch.object(converter, "start_libreoffice", return_value=(None, None, None, None)), \
                    patch.object(converter, "open_document", return_value=document):
                converter.convert_file(source)
            md = (output / "sample.md").read_text(encoding="utf-8")
            self.assertIn("## 문서 제목", md)
            self.assertLess(md.index("표 앞 문단"), md.index("| 항목 | 설명 |"))
            self.assertLess(md.index("병합 셀"), md.index("표 뒤 문단"))
            self.assertIn("분할 셀 1<br>분할 셀 2", md)
            self.assertIn("파이프 \\| 줄바꿈<br>다음 줄", md)
            self.assertEqual({p.suffix for p in output.iterdir()}, {".md", ".json"})
            document.close.assert_called_once_with(True)

    def test_xlsx_text_and_table_serialization(self):
        blocks = converter.find_blocks([
            ["표 앞 텍스트", ""], ["", ""],
            ["항목", "값"], ["한글", "내용 | 값"], ["수량", "42"],
        ])
        md = converter.blocks_to_markdown("Sheet1", blocks)
        self.assertIn("표 앞 텍스트", md)
        self.assertIn("| 항목 | 값 |", md)
        self.assertIn("내용 \\| 값", md)
        self.assertIn("42", md)


@unittest.skipUnless(os.environ.get("RUN_LIBREOFFICE_TESTS") == "1",
                     "Set RUN_LIBREOFFICE_TESTS=1 to run live LibreOffice checks")
class ConversionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Isolate the test server from any existing converter on port 2002.
        with socket.socket() as listener:
            listener.bind(("127.0.0.1", 0))
            port = listener.getsockname()[1]
        cls.connection_patch = patch.multiple(
            converter, UNO_PORT=port,
            UNO_CONNECTION=f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
        )
        cls.connection_patch.start()
        cls.addClassCleanup(cls.connection_patch.stop)
        cls.temp = tempfile.TemporaryDirectory(prefix="markdown_test_")
        cls.root = Path(cls.temp.name)
        with patch.object(converter, "BASE_DIR", cls.root):
            process, desktop, profile, log = converter.start_libreoffice()
        try:
            doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
            text = doc.getText()
            text.setString("문서 제목\n표 앞 문단\n")
            cursor = text.createTextCursor()
            cursor.gotoEnd(False)
            table = doc.createInstance("com.sun.star.text.TextTable")
            table.initialize(3, 2)
            text.insertTextContent(cursor, table, False)
            table.getCellByName("A1").setString("항목")
            table.getCellByName("B1").setString("설명")
            table.getCellByName("A2").setString("한글")
            table.getCellByName("B2").setString("파이프 | 줄바꿈\n다음 줄")
            merged = table.createCursorByCellName("A3")
            merged.gotoCellByName("B3", True)
            merged.mergeRange()
            table.getCellByName("A3").setString("병합 셀")
            cursor.gotoEnd(False)
            text.insertString(cursor, "\n표 뒤 문단", False)
            cls.docx = cls.root / "sample.DOCX"
            doc.storeAsURL(cls.docx.as_uri(), (
                converter.make_property("FilterName", "Office Open XML Text"),
            ))
            doc.close(True)

            doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
            sheet = doc.Sheets.getByIndex(0)
            sheet.getCellByPosition(0, 0).setString("표 앞 텍스트")
            sheet.getCellRangeByName("A3:B5").setDataArray((
                ("항목", "값"), ("한글", "내용 | 값"), ("수량", 42),
            ))
            cls.xlsx = cls.root / "sample.XLSX"
            doc.storeAsURL(cls.xlsx.as_uri(), (
                converter.make_property("FilterName", "Calc MS Excel 2007 XML"),
            ))
            doc.close(True)
        finally:
            process.terminate()
            process.wait(timeout=10)
            log.close()
            # LibreOffice's temporary profile is managed by the OS temp area.
            converter.shutil.rmtree(profile, ignore_errors=True)

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def convert(self, source, folder):
        output = self.root / folder
        with patch.object(converter, "BASE_DIR", self.root), \
                patch.object(converter, "INPUT_DIR", self.root / "input"), \
                patch.object(converter, "OUTPUT_DIR", output):
            converter.convert_file(source)
        return output

    def test_docx_text_table_order_and_merged_cells(self):
        output = self.convert(self.docx, "docx_output")
        md = (output / "sample.md").read_text(encoding="utf-8")
        self.assertLess(md.index("표 앞 문단"), md.index("| 항목 | 설명 |"))
        self.assertLess(md.index("병합 셀"), md.index("표 뒤 문단"))
        self.assertIn("파이프 \\| 줄바꿈<br>다음 줄", md)
        manifest = json.loads((output / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["engine"], "LibreOffice Writer UNO")
        self.assertEqual([b["type"] for b in manifest["blocks"]].count("table"), 1)
        self.assertEqual({p.suffix for p in output.iterdir()}, {".md", ".json"})

    def test_xlsx_preserves_existing_output(self):
        output = self.convert(self.xlsx, "xlsx_output")
        manifest = json.loads((output / "manifest.json").read_text(encoding="utf-8"))
        md = (output / manifest["sheets"][0]["file"]).read_text(encoding="utf-8")
        self.assertIn("표 앞 텍스트", md)
        self.assertIn("| 항목 | 값 |", md)
        self.assertIn("내용 \\| 값", md)
        self.assertIn("42", md)
        self.assertTrue((output / "workbook.md").is_file())


if __name__ == "__main__":
    unittest.main()
