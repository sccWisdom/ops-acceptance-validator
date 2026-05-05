import subprocess
import sys
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from scripts.ops_acceptance_validator import format_output, validate_paths

TMP_ROOT = Path(__file__).resolve().parents[1] / ".tmp-tests"


def _xml_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def make_docx(path: Path, paragraphs=None, table_rows=None):
    paragraphs = paragraphs or []
    table_rows = table_rows or []

    p_xml = "".join(
        f"<w:p><w:r><w:t>{_xml_escape(text)}</w:t></w:r></w:p>"
        for text in paragraphs
    )
    table_xml = ""
    if table_rows:
        rows = []
        for row in table_rows:
            cells = "".join(
                "<w:tc><w:p><w:r><w:t>"
                f"{_xml_escape(cell)}"
                "</w:t></w:r></w:p></w:tc>"
                for cell in row
            )
            rows.append(f"<w:tr>{cells}</w:tr>")
        table_xml = f"<w:tbl>{''.join(rows)}</w:tbl>"

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{p_xml}{table_xml}</w:body>"
        "</w:document>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _sheet_xml(values, merges=None):
    rows_xml = []
    for row_idx, row in enumerate(values, 1):
        cells_xml = []
        for col_idx, value in enumerate(row, 1):
            if value is None:
                continue
            col = chr(ord("A") + col_idx - 1)
            cells_xml.append(
                f'<c r="{col}{row_idx}" t="inlineStr"><is><t>{_xml_escape(value)}</t></is></c>'
            )
        rows_xml.append(f'<row r="{row_idx}">{"".join(cells_xml)}</row>')
    merge_xml = ""
    if merges:
        refs = "".join(f'<mergeCell ref="{ref}"/>' for ref in merges)
        merge_xml = f'<mergeCells count="{len(merges)}">{refs}</mergeCells>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(rows_xml)}</sheetData>{merge_xml}'
        "</worksheet>"
    )


def make_xlsx(path: Path):
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/worksheets/sheet2.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        '<sheet name="巡检记录" sheetId="1" r:id="rId1"/>'
        '<sheet name="系统清单" sheetId="2" r:id="rId2"/>'
        "</sheets></workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet2.xml"/>'
        "</Relationships>"
    )
    sheet1 = _sheet_xml(
        [
            ["2025年7月巡检", None],
            ["上海城投（集团）有限公司", "835台服务器"],
        ],
        merges=["A1:B1"],
    )
    sheet2 = _sheet_xml(
        [
            ["子系统名称"],
            ["长三角共享交换平台"],
            ["数据共享子系统"],
            ["数据共享子系统"],
        ]
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet1)
        zf.writestr("xl/worksheets/sheet2.xml", sheet2)


class ValidatorTests(unittest.TestCase):
    def setUp(self):
        TMP_ROOT.mkdir(exist_ok=True)

    def temp_dir(self):
        return TemporaryDirectory(dir=TMP_ROOT)

    def test_docx_flags_all_rule_families_in_order(self):
        with self.temp_dir() as tmp:
            sample = Path(tmp) / "验收材料.docx"
            make_docx(
                sample,
                paragraphs=[
                    "本次服务周期覆盖2026年3月1日巡检。",
                    "本项目包含18个子系统和八百三十三台服务器。",
                    "数据治理子系统由厂商完成支持。",
                    "CERT参与安全复核。",
                ],
                table_rows=[["备注", "南洋负责问题跟进"]],
            )

            suggestions = validate_paths([sample])
            output = format_output(suggestions)

        self.assertIn("验收材料.docx", output)
        self.assertIn("2026年3月1日", output)
        self.assertIn("2026-02-28", output)
        self.assertIn("18个子系统", output)
        self.assertIn("19个子系统", output)
        self.assertIn("八百三十三台服务器", output)
        self.assertIn("834台服务器", output)
        self.assertIn("数据治理子系统", output)
        self.assertIn("数据治理分析系统", output)
        self.assertIn("厂商", output)
        self.assertIn("实施团队", output)
        self.assertIn("CERT", output)
        self.assertIn("实施13组", output)
        self.assertIn("南洋", output)
        self.assertIn("实施1组", output)
        self.assertNotIn("统计", output)
        self.assertNotIn("通过", output)
        self.assertTrue(output.startswith("1. "))

    def test_xlsx_reads_sheets_merged_cells_and_system_list_issues(self):
        with self.temp_dir() as tmp:
            sample = Path(tmp) / "服务器清单.xlsx"
            make_xlsx(sample)

            output = format_output(validate_paths([sample]))

        self.assertIn("服务器清单.xlsx", output)
        self.assertIn("巡检记录", output)
        self.assertIn("2025年7月", output)
        self.assertIn("上海城投（集团）有限公司", output)
        self.assertIn("实施57组", output)
        self.assertIn("835台服务器", output)
        self.assertIn("834台服务器", output)
        self.assertIn("系统清单", output)
        self.assertIn("重复出现“数据共享子系统”", output)
        self.assertIn("缺少“安全管理子系统”", output)

    def test_clean_file_returns_fixed_no_issue_message(self):
        with self.temp_dir() as tmp:
            sample = Path(tmp) / "合格材料.docx"
            make_docx(
                sample,
                paragraphs=[
                    "服务时间为2025年9月，范围覆盖19个子系统和834台服务器。",
                    "本次工作由实施团队完成。",
                    "安全管理子系统运行正常。",
                ],
            )

            output = format_output(validate_paths([sample]))

        self.assertEqual(output, "未发现需要修改的内容。")

    def test_unsupported_file_extension_is_reported_as_suggestion(self):
        with self.temp_dir() as tmp:
            sample = Path(tmp) / "验收材料.txt"
            sample.write_text("2026年3月1日", encoding="utf-8")

            output = format_output(validate_paths([sample]))

        self.assertIn("验收材料.txt", output)
        self.assertIn("文件格式", output)
        self.assertIn(".docx", output)
        self.assertIn(".xlsx", output)

    def test_cli_outputs_only_markdown_suggestions(self):
        with self.temp_dir() as tmp:
            sample = Path(tmp) / "验收材料.docx"
            make_docx(sample, paragraphs=["2025年7月出现20个系统。"])

            result = subprocess.run(
                [
                    sys.executable,
                    "scripts/ops_acceptance_validator.py",
                    str(sample),
                ],
                cwd=Path(__file__).resolve().parents[1],
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )

        output = result.stdout.strip()
        self.assertTrue(output.startswith("1. "))
        self.assertIn("2025年7月", output)
        self.assertIn("20个系统", output)
        self.assertNotIn("{", output)
        self.assertNotIn("debug", output.lower())


if __name__ == "__main__":
    unittest.main()
