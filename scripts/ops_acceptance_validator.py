#!/usr/bin/env python3
from __future__ import annotations

import argparse
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


SERVICE_START = date(2025, 8, 26)
SERVICE_END = date(2026, 2, 28)

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SHEET_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "office": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

STANDARD_SUBSYSTEMS = [
    "长三角共享交换平台",
    "运营管理子系统",
    "运维监控子系统",
    "数据资产子系统",
    "数据治理分析系统",
    "数据支撑子系统",
    "数据开发子系统",
    "数据驾驶舱",
    "数据归集子系统",
    "数据共享子系统",
    "数据标签子系统",
    "上海市数据分析平台",
    "上海市公共数据开放平台",
    "前置机管理子系统",
    "空间地理资产管理子系统",
    "国家资源平台级联系统",
    "公共数据管理门户",
    "电子证照管理系统",
    "安全管理子系统",
]
STANDARD_SET = set(STANDARD_SUBSYSTEMS)

VENDOR_GROUPS = {
    "南洋": "实施1组",
    "亚信": "实施2组",
    "普元": "实施3组",
    "双杨": "实施4组",
    "观安": "实施5组",
    "杰奕": "实施6组",
    "新点": "实施7组",
    "宽文": "实施8组",
    "浪潮": "实施9组",
    "微芯": "实施10组",
    "星环": "实施11组",
    "海致": "实施12组",
    "cert": "实施13组",
    "理想": "实施14组",
    "复尧": "实施15组",
    "公安三所": "实施16组",
    "华东院": "实施17组",
    "华宇": "实施18组",
    "联合征信": "实施19组",
    "谋乐": "实施20组",
    "赛博": "实施21组",
    "软中": "实施22组",
    "生腾": "实施23组",
    "数股": "实施24组",
    "万达": "实施25组",
    "仪电鑫森": "实施26组",
    "毓赢": "实施27组",
    "长亭": "实施28组",
    "智巡密码": "实施29组",
    "捷睿": "实施30组",
    "大数元": "实施31组",
    "浪擎": "实施32组",
    "通办": "实施33组",
    "众恒": "实施34组",
    "豌豆": "实施35组",
    "数发科": "实施36组",
    "炎黄": "实施37组",
    "至一": "实施38组",
    "测绘院": "实施39组",
    "东方通": "实施40组",
    "瑞数": "实施41组",
    "闪捷": "实施42组",
    "天融信": "实施43组",
    "上海南康科技有限公司": "实施44组",
    "派拉": "实施45组",
    "海量": "实施46组",
    "安华金和": "实施47组",
    "数据集团": "实施48组",
    "工创": "实施49组",
    "上咨": "实施50组",
    "地听": "实施51组",
    "泛微": "实施52组",
    "银江": "实施53组",
    "上数金科": "实施54组",
    "科林利康": "实施55组",
    "数喆": "实施56组",
    "上海城投(集团)有限公司": "实施57组",
    "上海燃气有限公司": "实施58组",
    "上海城投水务(集团)有限公司": "实施59组",
    "优刻得": "实施60组",
    "数梦": "实施61组",
    "仪电团队": "实施62组",
    "鸿冠": "实施63组",
    "万得信息技术股份有限公司": "实施64组",
    "上海京东智联信息技术有限公司": "实施65组",
    "深圳力维智联技术有限公司": "实施66组",
    "华润燃气": "实施67组",
    "上海德拓信息技术股份有限公司": "实施68组",
    "上海城市地理信息系统发展有限公司": "实施69组",
    "建信金科": "实施70组",
    "上海联合产权交易所有限公司": "实施71组",
    "上海机场(集团)有限公司": "实施72组",
    "上海亿通国际股份有限公司": "实施73组",
    "上海随申行智慧交通科技有限公司": "实施74组",
    "上海云赛创鑫企业管理有限公司": "实施75组",
    "上海市数字证书认证中心有限公司": "实施76组",
    "联通数字科技有限公司": "实施77组",
    "上海思亮信息技术股份有限公司": "实施78组",
    "上海交易集团有限公司": "实施79组",
    "上海领再科技有限公司": "实施80组",
    "上海杰狮信息技术有限公司": "实施81组",
    "上海临床创新转化研究院有限公司": "实施82组",
    "上海北码科技有限公司": "实施83组",
    "上海骊宵医疗技术有限公司": "实施84组",
    "上海仪电数联信息技术有限公司": "实施85组",
    "怀策": "实施86组",
    "爱数": "实施87组",
    "阿里": "实施88组",
    "微星": "实施89组",
    "鼎茂": "实施90组",
    "龙石": "实施91组",
}

GENERIC_VENDOR_TERMS = ["厂商", "服务商", "运营商"]

CHINESE_DIGITS = {
    "零": 0,
    "〇": 0,
    "一": 1,
    "二": 2,
    "两": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
}
CHINESE_UNITS = {"十": 10, "百": 100, "千": 1000, "万": 10000}
CN_NUM = "零〇一二两三四五六七八九十百千万"
FULL_DATE_PATTERNS = [
    re.compile(r"(?<!\d)(20\d{2})[./-](\d{1,2})[./-](\d{1,2})(?!\d)"),
    re.compile(r"(?<!\d)(20\d{2})年(\d{1,2})月(\d{1,2})日"),
    re.compile(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)"),
]
MONTH_PATTERN = re.compile(r"(?<!\d)(20\d{2})(?:年|[./-])(\d{1,2})月?(?![./-]?\d|日)")
COUNT_PATTERN = re.compile(
    rf"(?P<num>\d+|[{CN_NUM}]+)\s*(?P<classifier>个|套|台)?\s*(?P<target>子系统|系统|服务器)"
)
SUBSYSTEM_CANDIDATE_PATTERN = re.compile(
    r"[\u4e00-\u9fa5A-Za-z0-9（）()]{2,24}(?:子系统|系统|平台|门户|驾驶舱)"
)


@dataclass(frozen=True)
class TextItem:
    file_name: str
    text: str
    sheet_name: str | None = None


@dataclass(frozen=True)
class Suggestion:
    file_name: str
    original: str
    message: str
    sheet_name: str | None = None
    position: int = 0


def validate_paths(paths: Iterable[str | Path]) -> list[Suggestion]:
    suggestions: list[Suggestion] = []
    for raw_path in paths:
        path = Path(raw_path)
        file_suggestions: list[Suggestion] = []
        suffix = path.suffix.lower()
        if suffix not in {".docx", ".xlsx"}:
            suggestions.append(
                Suggestion(
                    path.name,
                    path.name,
                    "文件格式不在检查范围内，建议提供 .docx 或 .xlsx 文件后重新检查。",
                )
            )
            continue
        try:
            items = read_docx(path) if suffix == ".docx" else read_xlsx(path)
        except Exception:
            suggestions.append(
                Suggestion(
                    path.name,
                    path.name,
                    "文件无法读取，建议确认文件未损坏后重新检查。",
                )
            )
            continue
        for item in items:
            file_suggestions.extend(check_text_item(item))
        file_suggestions.extend(check_subsystem_lists(items))
        suggestions.extend(_dedupe(file_suggestions))
    return suggestions


def format_output(suggestions: Iterable[Suggestion]) -> str:
    lines = []
    for idx, suggestion in enumerate(suggestions, 1):
        location = f"“{suggestion.file_name}”"
        if suggestion.sheet_name:
            location += f"的“{suggestion.sheet_name}”工作表"
        lines.append(
            f"{idx}. {location}中“{suggestion.original}”{suggestion.message}"
        )
    return "\n".join(lines) if lines else "未发现需要修改的内容。"


def read_docx(path: Path) -> list[TextItem]:
    items: list[TextItem] = []
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", WORD_NS)
    if body is None:
        return items
    for child in list(body):
        tag = _local_name(child.tag)
        if tag == "p":
            text = _word_text(child)
            if text:
                items.append(TextItem(path.name, text))
        elif tag == "tbl":
            for row in child.findall(".//w:tr", WORD_NS):
                cells = []
                for cell in row.findall("./w:tc", WORD_NS):
                    cell_text = _word_text(cell)
                    if cell_text:
                        cells.append(cell_text)
                if cells:
                    items.append(TextItem(path.name, " | ".join(cells)))
    return items


def read_xlsx(path: Path) -> list[TextItem]:
    items: list[TextItem] = []
    with zipfile.ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = _workbook_relationships(zf)
        shared_strings = _shared_strings(zf)
        date_style_flags = _date_style_flags(zf)
        for sheet in workbook.findall(".//main:sheet", SHEET_NS):
            sheet_name = sheet.attrib.get("name", "工作表")
            rel_id = sheet.attrib.get(f"{{{SHEET_NS['office']}}}id")
            target = rels.get(rel_id or "")
            if not target:
                continue
            sheet_path = target.lstrip("/")
            if not sheet_path.startswith("xl/"):
                sheet_path = posixpath.normpath(posixpath.join("xl", sheet_path))
            if sheet_path not in zf.namelist():
                continue
            root = ET.fromstring(zf.read(sheet_path))
            for row in root.findall(".//main:sheetData/main:row", SHEET_NS):
                for cell in row.findall("main:c", SHEET_NS):
                    text = _cell_text(cell, shared_strings, date_style_flags)
                    if text:
                        items.append(TextItem(path.name, text, sheet_name))
    return items


def check_text_item(item: TextItem) -> list[Suggestion]:
    suggestions: list[Suggestion] = []
    suggestions.extend(_check_dates(item))
    suggestions.extend(_check_counts(item))
    suggestions.extend(_check_subsystem_names(item))
    suggestions.extend(_check_generic_vendor_terms(item))
    suggestions.extend(_check_specific_vendors(item))
    return sorted(suggestions, key=lambda suggestion: suggestion.position)


def check_subsystem_lists(items: list[TextItem]) -> list[Suggestion]:
    suggestions: list[Suggestion] = []
    groups: dict[tuple[str, str | None], list[str]] = {}
    has_list_hint: dict[tuple[str, str | None], bool] = {}
    for item in items:
        key = (item.file_name, item.sheet_name)
        groups.setdefault(key, [])
        has_list_hint.setdefault(key, False)
        if re.search(r"子系统名称|系统清单|子系统清单|系统名称", item.text):
            has_list_hint[key] = True
        if item.sheet_name and re.search(r"系统|子系统|清单", item.sheet_name):
            has_list_hint[key] = True
        for name in STANDARD_SUBSYSTEMS:
            if name in item.text:
                groups[key].append(name)

    for (file_name, sheet_name), names in groups.items():
        if not names:
            continue
        if not has_list_hint[(file_name, sheet_name)] and len(set(names)) < 5:
            continue
        seen: set[str] = set()
        duplicate_reported: set[str] = set()
        for name in names:
            if name in seen and name not in duplicate_reported:
                suggestions.append(
                    Suggestion(
                        file_name,
                        name,
                        f"重复出现“{name}”，建议删除重复项。",
                        sheet_name,
                    )
                )
                duplicate_reported.add(name)
            seen.add(name)
        for name in STANDARD_SUBSYSTEMS:
            if name not in seen:
                suggestions.append(
                    Suggestion(
                        file_name,
                        name,
                        f"缺少“{name}”，建议补充。",
                        sheet_name,
                    )
                )
    return suggestions


def _check_dates(item: TextItem) -> list[Suggestion]:
    suggestions: list[Suggestion] = []
    occupied_spans: list[tuple[int, int]] = []
    for pattern in FULL_DATE_PATTERNS:
        for match in pattern.finditer(item.text):
            original = match.group(0)
            parsed = _parse_date_groups(match.groups())
            if parsed is None:
                continue
            occupied_spans.append(match.span())
            if parsed < SERVICE_START:
                suggestions.append(
                    Suggestion(
                        item.file_name,
                        original,
                        "早于服务周期开始时间，建议核实实际发生时间并修改为 2025-08-26 及以后的日期。",
                        item.sheet_name,
                        match.start(),
                    )
                )
            elif parsed > SERVICE_END:
                suggestions.append(
                    Suggestion(
                        item.file_name,
                        original,
                        "超出服务周期，建议核实实际发生时间并修改为 2026-02-28 及以前的日期。",
                        item.sheet_name,
                        match.start(),
                    )
                )
    for match in MONTH_PATTERN.finditer(item.text):
        if _overlaps(match.span(), occupied_spans):
            continue
        original = match.group(0)
        year, month = int(match.group(1)), int(match.group(2))
        if not 1 <= month <= 12:
            continue
        if (year, month) < (SERVICE_START.year, SERVICE_START.month):
            suggestions.append(
                Suggestion(
                    item.file_name,
                    original,
                    "不在服务周期内，建议修改为服务周期内的实际月份，或删除该时间表述。",
                    item.sheet_name,
                    match.start(),
                )
            )
        elif (year, month) > (SERVICE_END.year, SERVICE_END.month):
            suggestions.append(
                Suggestion(
                    item.file_name,
                    original,
                    "超出服务周期，建议修改为服务周期内的实际月份，或删除该时间表述。",
                    item.sheet_name,
                    match.start(),
                )
            )
    return suggestions


def _check_counts(item: TextItem) -> list[Suggestion]:
    suggestions: list[Suggestion] = []
    for match in COUNT_PATTERN.finditer(item.text):
        original = match.group(0)
        target = match.group("target")
        count = _parse_count(match.group("num"))
        if count is None:
            continue
        if target in {"子系统", "系统"} and count != 19:
            suggestions.append(
                Suggestion(
                    item.file_name,
                    original,
                    "与标准数量不一致，建议修改为“19个子系统”。",
                    item.sheet_name,
                    match.start(),
                )
            )
        elif target == "服务器" and count != 834:
            suggestions.append(
                Suggestion(
                    item.file_name,
                    original,
                    "与标准数量不一致，建议修改为“834台服务器”。",
                    item.sheet_name,
                    match.start(),
                )
            )
    return suggestions


def _check_subsystem_names(item: TextItem) -> list[Suggestion]:
    suggestions: list[Suggestion] = []
    reported: set[str] = set()
    for match in SUBSYSTEM_CANDIDATE_PATTERN.finditer(item.text):
        candidate = _trim_candidate(match.group(0))
        if candidate in STANDARD_SET or candidate in reported:
            continue
        suggestion = _closest_subsystem(candidate)
        if suggestion:
            suggestions.append(
                Suggestion(
                    item.file_name,
                    candidate,
                    f"不是标准名称，建议核实是否应修改为“{suggestion}”。",
                    item.sheet_name,
                    match.start(),
                )
            )
            reported.add(candidate)
    return suggestions


def _check_generic_vendor_terms(item: TextItem) -> list[Suggestion]:
    suggestions = []
    for term in GENERIC_VENDOR_TERMS:
        for match in re.finditer(re.escape(term), item.text):
            suggestions.append(
                Suggestion(
                    item.file_name,
                    term,
                    "属于不建议出现的表述，建议删除该表述，或改为“实施团队”。",
                    item.sheet_name,
                    match.start(),
                )
            )
    return suggestions


def _check_specific_vendors(item: TextItem) -> list[Suggestion]:
    suggestions: list[Suggestion] = []
    for name, group in _ordered_vendor_items():
        pattern = _vendor_pattern(name)
        for match in pattern.finditer(item.text):
            original = match.group(0)
            suggestions.append(
                Suggestion(
                    item.file_name,
                    original,
                    f"涉及具体合作伙伴名称，建议替换为“{group}”。",
                    item.sheet_name,
                    match.start(),
                )
            )
    return suggestions


def _word_text(node: ET.Element) -> str:
    values = [part.text or "" for part in node.findall(".//w:t", WORD_NS)]
    return "".join(values).strip()


def _workbook_relationships(zf: zipfile.ZipFile) -> dict[str, str]:
    rels: dict[str, str] = {}
    root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in root.findall("rel:Relationship", SHEET_NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rel_id and target:
            rels[rel_id] = target
    return rels


def _shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values = []
    for si in root.findall("main:si", SHEET_NS):
        text = "".join(t.text or "" for t in si.findall(".//main:t", SHEET_NS))
        values.append(text)
    return values


def _date_style_flags(zf: zipfile.ZipFile) -> list[bool]:
    if "xl/styles.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/styles.xml"))
    custom_formats: dict[str, str] = {}
    for num_fmt in root.findall(".//main:numFmts/main:numFmt", SHEET_NS):
        fmt_id = num_fmt.attrib.get("numFmtId")
        fmt_code = num_fmt.attrib.get("formatCode", "")
        if fmt_id:
            custom_formats[fmt_id] = fmt_code
    builtin_date_ids = {
        "14",
        "15",
        "16",
        "17",
        "22",
        "27",
        "28",
        "29",
        "30",
        "31",
        "32",
        "33",
        "34",
        "35",
        "36",
        "45",
        "46",
        "47",
        "50",
        "51",
        "52",
        "53",
        "54",
        "55",
        "56",
        "57",
        "58",
    }
    flags = []
    cell_xfs = root.find("main:cellXfs", SHEET_NS)
    if cell_xfs is None:
        return flags
    for xf in cell_xfs.findall("main:xf", SHEET_NS):
        fmt_id = xf.attrib.get("numFmtId", "")
        fmt_code = custom_formats.get(fmt_id, "")
        is_date = fmt_id in builtin_date_ids or bool(
            fmt_code and re.search(r"(?i)(y|m|d|年|月|日)", fmt_code)
        )
        flags.append(is_date)
    return flags


def _cell_text(
    cell: ET.Element, shared_strings: list[str], date_style_flags: list[bool]
) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(".//main:t", SHEET_NS)).strip()
    value = cell.find("main:v", SHEET_NS)
    if value is None or value.text is None:
        text_node = cell.find("main:is/main:t", SHEET_NS)
        return (text_node.text or "").strip() if text_node is not None else ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)].strip()
        except (ValueError, IndexError):
            return ""
    if _cell_has_date_style(cell, date_style_flags):
        converted = _excel_serial_to_date(value.text)
        if converted:
            return converted.isoformat()
    return value.text.strip()


def _cell_has_date_style(cell: ET.Element, date_style_flags: list[bool]) -> bool:
    style = cell.attrib.get("s")
    if style is None:
        return False
    try:
        index = int(style)
    except ValueError:
        return False
    return 0 <= index < len(date_style_flags) and date_style_flags[index]


def _excel_serial_to_date(value: str) -> date | None:
    try:
        serial = float(value)
    except ValueError:
        return None
    if serial <= 0:
        return None
    return date(1899, 12, 30) + timedelta(days=int(serial))


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _parse_date_groups(groups: tuple[str, ...]) -> date | None:
    try:
        parsed = date(int(groups[0]), int(groups[1]), int(groups[2]))
    except ValueError:
        return None
    return parsed


def _overlaps(span: tuple[int, int], spans: list[tuple[int, int]]) -> bool:
    return any(span[0] < existing[1] and existing[0] < span[1] for existing in spans)


def _parse_count(value: str) -> int | None:
    if value.isdigit():
        return int(value)
    return _chinese_to_int(value)


def _chinese_to_int(value: str) -> int | None:
    if not value:
        return None
    total = 0
    section = 0
    number = 0
    for char in value:
        if char in CHINESE_DIGITS:
            number = CHINESE_DIGITS[char]
        elif char in CHINESE_UNITS:
            unit = CHINESE_UNITS[char]
            if unit == 10000:
                section = (section + number) * unit
                total += section
                section = 0
            else:
                if number == 0:
                    number = 1
                section += number * unit
            number = 0
        else:
            return None
    return total + section + number


def _trim_candidate(candidate: str) -> str:
    candidate = re.sub(r"^[0-9一二三四五六七八九十百千万两零〇个套项：:、，,\s]+", "", candidate)
    return candidate.strip(" ，,。；;：:")


def _closest_subsystem(candidate: str) -> str | None:
    from difflib import SequenceMatcher

    if len(candidate) < 4:
        return None
    best_name = None
    best_score = 0.0
    for name in STANDARD_SUBSYSTEMS:
        score = SequenceMatcher(None, candidate, name).ratio()
        if score > best_score:
            best_score = score
            best_name = name
    return best_name if best_name and best_score >= 0.68 else None


def _ordered_vendor_items() -> list[tuple[str, str]]:
    return sorted(VENDOR_GROUPS.items(), key=lambda item: len(item[0]), reverse=True)


def _vendor_pattern(name: str) -> re.Pattern[str]:
    if name.lower() == "cert":
        return re.compile(r"cert", re.IGNORECASE)
    escaped = re.escape(name)
    escaped = escaped.replace(r"\(", r"[\(（]").replace(r"\)", r"[\)）]")
    return re.compile(escaped)


def _dedupe(suggestions: list[Suggestion]) -> list[Suggestion]:
    seen: set[tuple[str, str | None, str, str]] = set()
    result = []
    for suggestion in suggestions:
        key = (
            suggestion.file_name,
            suggestion.sheet_name,
            suggestion.original,
            suggestion.message,
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(suggestion)
    return result


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(
        description="检查运维验收材料，并只输出修正建议。"
    )
    parser.add_argument("files", nargs="+", help="待检查的 .docx 或 .xlsx 文件")
    args = parser.parse_args(argv)
    output = format_output(validate_paths(args.files))
    sys.stdout.write(output + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
