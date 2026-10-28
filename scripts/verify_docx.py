#!/usr/bin/env python3
"""Independent format checks for the five editable patent DOCX files."""

from __future__ import annotations

import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
from patent_config import (  # noqa: E402
    captions_from_project,
    project_dir_from_args,
    title_from_project,
)

# 由 configure() 在运行时填入，本模块不写死任何具体专利的内容。
PROJECT_DIR: Path | None = None
WORD_DIR: Path | None = None
EXPECTED_TITLE = ""
EXPECTED_CAPTIONS: list[str] = []


def configure(project_dir: Path) -> None:
    global PROJECT_DIR, WORD_DIR, EXPECTED_TITLE, EXPECTED_CAPTIONS
    PROJECT_DIR = project_dir
    WORD_DIR = project_dir / "word"
    EXPECTED_TITLE = title_from_project(project_dir)
    EXPECTED_CAPTIONS = captions_from_project(project_dir)
EXPECTED_HEADINGS = ["技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式"]

OUTPUT_FILES = [
    "100001权利要求书.docx",
    "100002说明书.docx",
    "100003说明书附图.docx",
    "100004说明书摘要.docx",
    "摘要附图.docx",
]

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
}


@dataclass
class CheckResult:
    number: int
    passed: bool
    summary: str
    details: list[str]


def read_document_xml(path: Path) -> bytes:
    with zipfile.ZipFile(path) as package:
        return package.read("word/document.xml")


def document_root(path: Path) -> etree._Element:
    return etree.fromstring(read_document_xml(path))


def filing_body_children(root: etree._Element) -> list[etree._Element]:
    """Return direct body children before the filing-to-notice section break."""
    body_nodes = root.xpath("/w:document/w:body", namespaces=NS)
    if len(body_nodes) != 1:
        return []
    result: list[etree._Element] = []
    for child in body_nodes[0]:
        if child.xpath("./w:pPr/w:sectPr", namespaces=NS):
            break
        result.append(child)
    return result


def filing_paragraph_nodes(root: etree._Element) -> list[etree._Element]:
    return [
        child
        for child in filing_body_children(root)
        if child.tag == f"{{{NS['w']}}}p"
    ]


def media_members(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        return sorted(
            name
            for name in package.namelist()
            if name.startswith("word/media/") and not name.endswith("/")
        )


def all_xml_text(path: Path) -> str:
    chunks: list[str] = []
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                chunks.append(package.read(name).decode("utf-8", errors="replace"))
    return "\n".join(chunks)


class StyleResolver:
    """Resolve the effective OOXML formatting that Word displays after resave.

    Word commonly moves direct formatting to paragraph properties, the default
    paragraph style, or document defaults.  Looking only at ``w:rPr`` therefore
    creates false failures after an otherwise harmless manual Word edit.
    """

    def __init__(self, path: Path) -> None:
        with zipfile.ZipFile(path) as package:
            self.styles = etree.fromstring(package.read("word/styles.xml"))
        self.styles_by_type_and_id: dict[tuple[str, str], etree._Element] = {}
        self.default_style_ids: dict[str, str] = {}
        for style in self.styles.xpath("/w:styles/w:style", namespaces=NS):
            style_type = style.get(f"{{{NS['w']}}}type", "")
            style_id = style.get(f"{{{NS['w']}}}styleId", "")
            if style_type and style_id:
                self.styles_by_type_and_id[(style_type, style_id)] = style
                if style.get(f"{{{NS['w']}}}default") == "1":
                    self.default_style_ids[style_type] = style_id

    @staticmethod
    def _property_value(
        properties: etree._Element | None,
        property_name: str,
        attribute_name: str,
    ) -> str | None:
        if properties is None:
            return None
        nodes = properties.xpath(f"./w:{property_name}", namespaces=NS)
        if not nodes:
            return None
        return nodes[0].get(f"{{{NS['w']}}}{attribute_name}")

    @staticmethod
    def _boolean_value(
        properties: etree._Element | None, property_name: str
    ) -> bool | None:
        if properties is None:
            return None
        nodes = properties.xpath(f"./w:{property_name}", namespaces=NS)
        if not nodes:
            return None
        value = nodes[0].get(f"{{{NS['w']}}}val", "true").lower()
        return value not in {"0", "false", "off", "no"}

    def _style_chain(
        self, style_type: str, style_id: str | None
    ) -> list[etree._Element]:
        if not style_id:
            style_id = self.default_style_ids.get(style_type)
        chain: list[etree._Element] = []
        visited: set[str] = set()
        while style_id and style_id not in visited:
            visited.add(style_id)
            style = self.styles_by_type_and_id.get((style_type, style_id))
            if style is None:
                break
            chain.append(style)
            based_on = style.xpath("string(./w:basedOn/@w:val)", namespaces=NS)
            style_id = based_on or None
        return chain

    def _run_property_sources(
        self, run: etree._Element, paragraph: etree._Element
    ) -> list[etree._Element | None]:
        run_style_id = run.xpath("string(./w:rPr/w:rStyle/@w:val)", namespaces=NS)
        paragraph_style_id = paragraph.xpath(
            "string(./w:pPr/w:pStyle/@w:val)", namespaces=NS
        )
        sources: list[etree._Element | None] = [run.find(f"{{{NS['w']}}}rPr")]
        sources.extend(
            style.find(f"{{{NS['w']}}}rPr")
            for style in self._style_chain("character", run_style_id or None)
        )
        paragraph_properties = paragraph.find(f"{{{NS['w']}}}pPr")
        sources.append(
            paragraph_properties.find(f"{{{NS['w']}}}rPr")
            if paragraph_properties is not None
            else None
        )
        sources.extend(
            style.find(f"{{{NS['w']}}}rPr")
            for style in self._style_chain("paragraph", paragraph_style_id or None)
        )
        if not run_style_id:
            sources.extend(
                style.find(f"{{{NS['w']}}}rPr")
                for style in self._style_chain("character", None)
            )
        default_run_properties = self.styles.find(
            f"{{{NS['w']}}}docDefaults/{{{NS['w']}}}rPrDefault/"
            f"{{{NS['w']}}}rPr"
        )
        sources.append(default_run_properties)
        return sources

    def run_value(
        self,
        run: etree._Element,
        paragraph: etree._Element,
        property_name: str,
        attribute_name: str = "val",
    ) -> str | None:
        for properties in self._run_property_sources(run, paragraph):
            value = self._property_value(properties, property_name, attribute_name)
            if value is not None:
                return value
        return None

    def run_boolean(
        self, run: etree._Element, paragraph: etree._Element, property_name: str
    ) -> bool:
        for properties in self._run_property_sources(run, paragraph):
            value = self._boolean_value(properties, property_name)
            if value is not None:
                return value
        return False

    def paragraph_value(
        self,
        paragraph: etree._Element,
        property_name: str,
        attribute_name: str = "val",
    ) -> str | None:
        paragraph_properties = paragraph.find(f"{{{NS['w']}}}pPr")
        value = self._property_value(
            paragraph_properties, property_name, attribute_name
        )
        if value is not None:
            return value
        style_id = paragraph.xpath(
            "string(./w:pPr/w:pStyle/@w:val)", namespaces=NS
        )
        for style in self._style_chain("paragraph", style_id or None):
            value = self._property_value(
                style.find(f"{{{NS['w']}}}pPr"),
                property_name,
                attribute_name,
            )
            if value is not None:
                return value
        default_paragraph_properties = self.styles.find(
            f"{{{NS['w']}}}docDefaults/{{{NS['w']}}}pPrDefault/"
            f"{{{NS['w']}}}pPr"
        )
        return self._property_value(
            default_paragraph_properties, property_name, attribute_name
        )


def run_checks() -> list[CheckResult]:
    results: list[CheckResult] = []
    paths = {name: WORD_DIR / name for name in OUTPUT_FILES}
    documents = {}

    # 1. Existence and python-docx reopen.
    details: list[str] = []
    errors: list[str] = []
    for name, path in paths.items():
        if not path.is_file():
            errors.append(f"missing: {path}")
            continue
        try:
            documents[name] = Document(str(path))
            details.append(f"{name}: opened, {path.stat().st_size} bytes")
        except Exception as exc:
            errors.append(f"{name}: cannot open: {exc}")
    details.extend(errors)
    results.append(
        CheckResult(
            1,
            not errors and len(documents) == 5,
            "五个 DOCX 存在且可由 python-docx 打开",
            details,
        )
    )

    # 2. Description title, headings, and absence of body drawings.
    details = []
    errors = []
    name = "100002说明书.docx"
    if name not in documents:
        errors.append("说明书未成功打开")
    else:
        document = documents[name]
        nonempty = [paragraph for paragraph in document.paragraphs if paragraph.text.strip()]
        if not nonempty:
            errors.append("说明书没有非空段落")
        else:
            first = nonempty[0]
            details.append(f"首个非空段落={first.text!r}, alignment={first.alignment}")
            if first.text != EXPECTED_TITLE:
                errors.append("首个非空段落不是指定发明名称")
            if first.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("发明名称段落未居中")
        paragraph_texts = [paragraph.text.strip() for paragraph in document.paragraphs]
        missing_headings = [heading for heading in EXPECTED_HEADINGS if heading not in paragraph_texts]
        details.append(f"五段式标题命中={len(EXPECTED_HEADINGS) - len(missing_headings)}/5")
        if missing_headings:
            errors.append(f"缺少标题: {missing_headings}")
        drawing_count = len(document_root(paths[name]).xpath("//w:drawing", namespaces=NS))
        details.append(f"document.xml 中 w:drawing={drawing_count}")
        if drawing_count != 0:
            errors.append("说明书正文含图片")
    results.append(CheckResult(2, not errors, "说明书名称、五段式标题及无图检查", details + errors))

    # 3. Claims 1 through 11.
    details = []
    errors = []
    name = "100001权利要求书.docx"
    claim_numbers: list[int] = []
    if name not in documents:
        errors.append("权利要求书未成功打开")
    else:
        root = document_root(paths[name])
        for paragraph in filing_paragraph_nodes(root):
            text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
            match = re.match(r"^\s*(\d+)、\s*", text)
            if match:
                claim_numbers.append(int(match.group(1)))
        details.append(f"识别到的权利要求编号={claim_numbers}")
        if len(claim_numbers) < 2 or claim_numbers != list(range(1, len(claim_numbers) + 1)):
            errors.append("权利要求编号不是从 1 起的连续编号")
    results.append(CheckResult(3, not errors, "权利要求书从 1 起连续编号", details + errors))

    # 4. Report OMML counts in claims and description; reject empty formulas.
    details = []
    errors = []
    total_omath = 0
    total_omath_para = 0
    for name in ("100001权利要求书.docx", "100002说明书.docx"):
        if name not in documents:
            errors.append(f"无法统计 {name} 的 OMML")
            continue
        root = document_root(paths[name])
        omath_count = len(root.xpath("//m:oMath", namespaces=NS))
        omath_para_count = len(root.xpath("//m:oMathPara", namespaces=NS))
        empty_omath_count = len(
            [
                node
                for node in root.xpath("//m:oMath", namespaces=NS)
                if not "".join(node.itertext())
            ]
        )
        details.append(
            f"{name}: m:oMath={omath_count}, m:oMathPara={omath_para_count}, "
            f"空公式={empty_omath_count}"
        )
        total_omath += omath_count
        total_omath_para += omath_para_count
        if empty_omath_count:
            errors.append(f"{name} 含空 OMML 公式")
    details.append(
        f"合计 m:oMath={total_omath}, 独立公式 m:oMathPara={total_omath_para}；"
        "数量仅报告，不与外部源比对"
    )
    results.append(
        CheckResult(4, not errors, "DOCX 内 OMML 数量报告及空公式检查", details + errors)
    )

    # 5. DOCX table structure checks, independent of any external text source.
    details = []
    errors = []
    name = "100002说明书.docx"
    if name not in documents:
        errors.append("说明书未成功打开，无法检查表格")
    else:
        document = documents[name]
        tables = document.tables
        details.append(f"w:tbl 数量={len(tables)}")
        if len(tables) < 1:
            errors.append(f"说明书表格数应至少为 1，实际为 {len(tables)}")
        dimensions: list[tuple[int, int]] = []
        for index, table in enumerate(tables, start=1):
            row_count = len(table.rows)
            grid_columns = len(table._tbl.xpath("./w:tblGrid/w:gridCol"))
            row_widths = [len(row.cells) for row in table.rows]
            logical_columns = row_widths[0] if row_widths else 0
            header = (
                [cell.text.strip() for cell in table.rows[0].cells]
                if table.rows
                else []
            )
            empty_header_columns = [
                column
                for column, text in enumerate(header, start=1)
                if not text
            ]
            dimensions.append((row_count, logical_columns))
            details.append(
                f"表{index}: rows={row_count}, grid columns={grid_columns}, "
                f"各行逻辑列数={row_widths}, 空表头列={empty_header_columns}"
            )
            if row_count == 0 or logical_columns == 0:
                errors.append(f"表{index} 为空表")
            if len(set(row_widths)) > 1:
                errors.append(f"表{index} 各行逻辑列数不一致: {row_widths}")
            if logical_columns != grid_columns:
                errors.append(
                    f"表{index} 逻辑列数 {logical_columns} 与 tblGrid {grid_columns} 不一致"
                )
            if empty_header_columns:
                errors.append(f"表{index} 存在空表头列: {empty_header_columns}")
        if dimensions:
            details.append(f"各表行列数={dimensions}")
    results.append(
        CheckResult(5, not errors, "说明书表格数量、行列一致性及非空表头自检", details + errors)
    )

    # 6. Drawing and media counts.
    details = []
    errors = []
    for name, expected in (
        ("100003说明书附图.docx", len(EXPECTED_CAPTIONS)),
        ("摘要附图.docx", 1),
    ):
        if name not in documents:
            errors.append(f"{name} 未成功打开")
            continue
        root = document_root(paths[name])
        drawing_count = sum(
            len(child.xpath(".//w:drawing", namespaces=NS))
            for child in filing_body_children(root)
        )
        media_count = len(media_members(paths[name]))
        details.append(f"{name}: w:drawing={drawing_count}, media={media_count}")
        if drawing_count != expected or media_count != expected:
            errors.append(f"{name} 应嵌入 {expected} 张图")
    results.append(CheckResult(6, not errors, "说明书附图与摘要附图嵌入数量检查", details + errors))

    # 7. Internal blockquote/work-note markers absent from every package XML.
    details = []
    errors = []
    forbidden_markers = [
        "临时数值，提交前必须替换",
        "附图重画清单",
        "状态：草稿",
        "公式暂以可读形式",
        "本草" + "案补充说明",
        "摘要附图：S1–S5",
    ]
    for name, path in paths.items():
        if name not in documents:
            errors.append(f"无法扫描 {name}")
            continue
        package_text = all_xml_text(path)
        found = [marker for marker in forbidden_markers if marker in package_text]
        details.append(f"{name}: 命中禁入标记={found}")
        if found:
            errors.append(f"{name} 含内部工作注记: {found}")
    results.append(CheckResult(7, not errors, "所有产物均不含 blockquote 内部工作注记", details + errors))

    # 8. Strict Word body-child validity and OMML namespace/parent checks.
    details = []
    errors = []
    allowed_body_children = {
        f"{{{NS['w']}}}p",
        f"{{{NS['w']}}}tbl",
        f"{{{NS['w']}}}sectPr",
        f"{{{NS['w']}}}bookmarkStart",
        f"{{{NS['w']}}}bookmarkEnd",
    }

    def display_qname(element) -> str:
        qname = etree.QName(element)
        if qname.namespace == NS["w"]:
            return f"w:{qname.localname}"
        if qname.namespace == NS["m"]:
            return f"m:{qname.localname}"
        return f"{{{qname.namespace}}}{qname.localname}"

    for name, path in paths.items():
        if name not in documents:
            errors.append(f"无法执行 {name} 的 OOXML 结构检查")
            continue
        root = document_root(path)
        body_nodes = root.xpath("/w:document/w:body", namespaces=NS)
        if len(body_nodes) != 1:
            errors.append(f"{name} 的 w:body 数量不是 1")
            continue
        body = body_nodes[0]
        direct_children = list(body)
        invalid_children = [
            display_qname(child) for child in direct_children if child.tag not in allowed_body_children
        ]
        child_counts: dict[str, int] = {}
        for child in direct_children:
            label = display_qname(child)
            child_counts[label] = child_counts.get(label, 0) + 1

        math_paragraphs = root.xpath("//m:oMathPara", namespaces=NS)
        invalid_math_parents: list[str] = []
        for node in math_paragraphs:
            parent = node.getparent()
            if parent is None:
                invalid_math_parents.append("<none>")
            elif parent.tag != f"{{{NS['w']}}}p":
                invalid_math_parents.append(display_qname(parent))
        math_namespace = root.nsmap.get("m")
        namespace_ok = math_namespace == NS["m"]
        details.append(
            f"{name}: body直接子元素={child_counts}, 非法子元素={invalid_children}, "
            f"m:oMathPara={len(math_paragraphs)}, 非w:p父元素={invalid_math_parents}, "
            f"xmlns:m={math_namespace!r}"
        )
        if invalid_children:
            errors.append(f"{name} 的 w:body 含非法直接子元素: {invalid_children}")
        if invalid_math_parents:
            errors.append(f"{name} 含未包裹在 w:p 中的 m:oMathPara")
        if not namespace_ok:
            errors.append(f"{name} 未在 w:document 根元素正确声明 xmlns:m")

    results.append(
        CheckResult(
            8,
            not errors,
            "五个产物的 w:body 子元素、独立公式父节点及数学命名空间检查",
            details + errors,
        )
    )

    # 9. Two template sections, filing/notice headers, PAGE fields, and form identity.
    details = []
    errors = []
    expected_forms = {
        "100001权利要求书.docx": ("权利要求书", "100001"),
        "100002说明书.docx": ("说明书", "100002"),
        "100003说明书附图.docx": ("说明书附图", "100003"),
        "100004说明书摘要.docx": ("说明书摘要", "100004"),
        "摘要附图.docx": ("说明书附图", "100003"),
    }
    for name, (expected_header, expected_code) in expected_forms.items():
        if name not in documents:
            errors.append(f"无法执行 {name} 的页眉页脚检查")
            continue
        path = paths[name]
        style_resolver = StyleResolver(path)
        with zipfile.ZipFile(path) as package:
            members = package.namelist()
            header_parts = sorted(
                member
                for member in members
                if re.fullmatch(r"word/header\d+\.xml", member)
            )
            footer_parts = sorted(
                member
                for member in members
                if re.fullmatch(r"word/footer\d+\.xml", member)
            )
            root = etree.fromstring(package.read("word/document.xml"))
            relationships = etree.fromstring(package.read("word/_rels/document.xml.rels"))
            section_properties = root.xpath("//w:sectPr", namespaces=NS)
            section_count = len(section_properties)
            header_reference_count = len(root.xpath("//w:headerReference", namespaces=NS))
            footer_reference_count = len(root.xpath("//w:footerReference", namespaces=NS))
            header_relationships = [
                rel
                for rel in relationships.xpath("./pr:Relationship", namespaces=NS)
                if rel.get("Type", "").endswith("/header")
            ]
            footer_relationships = [
                rel
                for rel in relationships.xpath("./pr:Relationship", namespaces=NS)
                if rel.get("Type", "").endswith("/footer")
            ]
            relationships_by_id = {
                rel.get("Id"): rel
                for rel in relationships.xpath("./pr:Relationship", namespaces=NS)
            }

            def relationship_member(relationship_id: str | None) -> str:
                relationship = relationships_by_id.get(relationship_id)
                if relationship is None:
                    return ""
                target = relationship.get("Target", "").replace("\\", "/")
                if target.startswith("/"):
                    return target.lstrip("/")
                return target if target.startswith("word/") else f"word/{target}"

            filing_header_member = ""
            notice_header_member = ""
            filing_footer_member = ""
            if section_count == 2:
                filing_header_ids = section_properties[0].xpath(
                    "./w:headerReference[@w:type='default']/@r:id", namespaces=NS
                )
                notice_header_ids = section_properties[-1].xpath(
                    "./w:headerReference[@w:type='default']/@r:id", namespaces=NS
                )
                filing_footer_ids = section_properties[0].xpath(
                    "./w:footerReference[@w:type='default']/@r:id", namespaces=NS
                )
                filing_header_member = relationship_member(
                    filing_header_ids[0] if len(filing_header_ids) == 1 else None
                )
                notice_header_member = relationship_member(
                    notice_header_ids[0] if len(notice_header_ids) == 1 else None
                )
                filing_footer_member = relationship_member(
                    filing_footer_ids[0] if len(filing_footer_ids) == 1 else None
                )

            if filing_header_member in members:
                filing_header_root = etree.fromstring(package.read(filing_header_member))
                header_text = "".join(
                    filing_header_root.xpath("//w:t/text()", namespaces=NS)
                )
                header_text_runs = filing_header_root.xpath("//w:r[w:t]", namespaces=NS)
                header_format_ok = (
                    bool(header_text_runs)
                    and bool(filing_header_root.xpath("//w:b", namespaces=NS))
                    and all(
                        run.xpath("string(./w:rPr/w:sz/@w:val)", namespaces=NS) == "28"
                        and run.xpath(
                            "string(./w:rPr/w:rFonts/@w:eastAsia)", namespaces=NS
                        )
                        == "黑体"
                        for run in header_text_runs
                    )
                )
            else:
                header_text = ""
                header_format_ok = False

            if notice_header_member in members:
                notice_header_root = etree.fromstring(package.read(notice_header_member))
                notice_header_text = "".join(
                    notice_header_root.xpath("//w:t/text()", namespaces=NS)
                )
            else:
                notice_header_text = ""

            footer_text = ""
            page_instructions: list[str] = []
            page_field_size_ok = True
            footer_page_status: list[str] = []
            for footer_member in footer_parts:
                footer_root = etree.fromstring(package.read(footer_member))
                part_text = "".join(footer_root.xpath("//w:t/text()", namespaces=NS))
                instructions = footer_root.xpath("//w:instrText/text()", namespaces=NS)
                page_runs = [
                    run
                    for run in footer_root.xpath("//w:r[w:instrText]", namespaces=NS)
                    if re.search(
                        r"\bPAGE\b",
                        "".join(run.xpath("./w:instrText/text()", namespaces=NS)),
                        flags=re.IGNORECASE,
                    )
                ]
                part_has_page = bool(page_runs)
                part_size_ok = part_has_page and all(
                    style_resolver.run_value(run, run.getparent(), "sz") == "18"
                    and style_resolver.run_value(run, run.getparent(), "szCs")
                    == "18"
                    for run in page_runs
                )
                footer_page_status.append(
                    f"{footer_member}:PAGE={part_has_page},9pt={part_size_ok}"
                )
                page_field_size_ok = page_field_size_ok and part_size_ok
                page_instructions.extend(instructions)
                if footer_member == filing_footer_member:
                    footer_text = part_text
            has_page_field = bool(footer_parts) and all(
                "PAGE=True" in status for status in footer_page_status
            )

        details.append(
            f"{name}: sectPr={section_count}, headerReference={header_reference_count}, "
            f"footerReference={footer_reference_count}, headerParts={header_parts}, "
            f"footerParts={footer_parts}, headerRels={len(header_relationships)}, "
            f"footerRels={len(footer_relationships)}, filingHeaderPart={filing_header_member}, "
            f"header={header_text!r}, noticeHeaderPart={notice_header_member}, "
            f"noticeHeader={notice_header_text!r}, "
            f"footerText={footer_text!r}, PAGE={page_instructions}, "
            f"footerPAGE={footer_page_status}, header黑体14pt加粗={header_format_ok}, "
            f"页码9pt={page_field_size_ok}"
        )
        expected_footer_count = 2 if name == "100004说明书摘要.docx" else 1
        if section_count != 2:
            errors.append(f"{name} 的 sectPr 数量不是 2")
        if header_reference_count != 2:
            errors.append(f"{name} 应有正文和注意事项两个页眉引用")
        if footer_reference_count != expected_footer_count:
            errors.append(
                f"{name} 页脚引用应为 {expected_footer_count}，实际为 {footer_reference_count}"
            )
        if len(header_parts) != 2 or len(footer_parts) != expected_footer_count:
            errors.append(f"{name} 未保留模板原有页眉页脚部件数量")
        if len(header_relationships) != 2 or len(footer_relationships) != expected_footer_count:
            errors.append(f"{name} 未保留模板原有页眉页脚 rels 关系")
        if "注意事项" in header_text:
            errors.append(f"{name} 第一节页眉误含注意事项")
        if header_text != expected_header:
            errors.append(f"{name} 页眉应为 {expected_header!r}，实际为 {header_text!r}")
        if notice_header_text != "注意事项":
            errors.append(f"{name} 文末第二节页眉不是‘注意事项’")
        if not header_format_ok:
            errors.append(f"{name} 第一节页眉不是加粗黑体 14pt")
        if expected_code not in footer_text or "2023.03" not in footer_text:
            errors.append(f"{name} 页脚缺少表单代码 {expected_code} 或日期")
        if not has_page_field:
            errors.append(f"{name} 存在未写入 PAGE 域的页脚部件")
        if not page_field_size_ok:
            errors.append(f"{name} 页码域不是 9pt")

    results.append(
        CheckResult(
            9,
            not errors,
            "正文与注意事项双节、模板页眉页脚、PAGE 域及表单代码检查",
            details + errors,
        )
    )

    # 10. Explicit typography, OMML size, captions, tables, and body paragraph layout.
    details = []
    errors = []

    def paragraph_text(paragraph_node) -> str:
        return "".join(paragraph_node.xpath(".//w:t/text()", namespaces=NS))

    def visible_runs(paragraph_node) -> list[etree._Element]:
        return paragraph_node.xpath("./w:r[w:t[string-length(.) > 0]]", namespaces=NS)

    def run_format_matches(
        run_node,
        *,
        resolver: StyleResolver,
        size: str,
        east_asia: str,
        latin: str,
    ) -> bool:
        paragraphs = run_node.xpath("ancestor::w:p[1]", namespaces=NS)
        if len(paragraphs) != 1:
            return False
        paragraph = paragraphs[0]
        return (
            resolver.run_value(run_node, paragraph, "sz") == size
            and resolver.run_value(run_node, paragraph, "szCs") == size
            and resolver.run_value(run_node, paragraph, "rFonts", "eastAsia")
            == east_asia
            and resolver.run_value(run_node, paragraph, "rFonts", "ascii")
            == latin
            and resolver.run_value(run_node, paragraph, "rFonts", "hAnsi")
            == latin
        )

    def paragraph_is_single_spaced_with_before_60(
        paragraph_node,
        expected_alignment: str,
        resolver: StyleResolver,
    ) -> bool:
        # 格式基准行距：固定值 24 磅（line=480, exact）。
        # 独立公式段与图片段例外为 atLeast，不在本函数检查范围（无 w:t 文本）。
        return (
            resolver.paragraph_value(paragraph_node, "spacing", "before") == "60"
            and resolver.paragraph_value(paragraph_node, "spacing", "line") == "480"
            and resolver.paragraph_value(paragraph_node, "spacing", "lineRule")
            == "exact"
            and resolver.paragraph_value(paragraph_node, "jc")
            == expected_alignment
        )

    claims_root = document_root(paths["100001权利要求书.docx"])
    description_root = document_root(paths["100002说明书.docx"])
    abstract_root = document_root(paths["100004说明书摘要.docx"])
    claims_styles = StyleResolver(paths["100001权利要求书.docx"])
    description_styles = StyleResolver(paths["100002说明书.docx"])
    abstract_styles = StyleResolver(paths["100004说明书摘要.docx"])

    claims_paragraphs = filing_paragraph_nodes(claims_root)
    claims_runs = [run for paragraph in claims_paragraphs for run in visible_runs(paragraph)]
    claims_format_ok = sum(
        run_format_matches(
            run,
            resolver=claims_styles,
            size="24",
            east_asia="宋体",
            latin="Times New Roman",
        )
        for run in claims_runs
    )
    claims_ratio = claims_format_ok / len(claims_runs) if claims_runs else 0.0
    claims_prose = [paragraph for paragraph in claims_paragraphs if paragraph_text(paragraph)]
    claims_layout_ok = sum(
        paragraph_is_single_spaced_with_before_60(
            paragraph, "both", claims_styles
        )
        for paragraph in claims_prose
    )

    excluded_description_paragraphs = {EXPECTED_TITLE, *EXPECTED_HEADINGS}
    description_paragraphs = filing_paragraph_nodes(description_root)
    description_body_paragraphs = [
        paragraph
        for paragraph in description_paragraphs
        if paragraph_text(paragraph)
        and paragraph_text(paragraph) not in excluded_description_paragraphs
        and not re.match(r"^表\d+、", paragraph_text(paragraph))
    ]
    description_runs = [
        run for paragraph in description_body_paragraphs for run in visible_runs(paragraph)
    ]
    description_format_ok = sum(
        run_format_matches(
            run,
            resolver=description_styles,
            size="24",
            east_asia="宋体",
            latin="Times New Roman",
        )
        for run in description_runs
    )
    description_ratio = description_format_ok / len(description_runs) if description_runs else 0.0
    description_layout_ok = sum(
        paragraph_is_single_spaced_with_before_60(
            paragraph, "both", description_styles
        )
        for paragraph in description_body_paragraphs
    )

    title_status: list[str] = []
    # 格式基准：名称宋体 14pt 不加粗；节标题与正文同为宋体 12pt 不加粗。
    for expected_text, expected_size in (
        (EXPECTED_TITLE, "28"),
        *((heading, "24") for heading in EXPECTED_HEADINGS),
    ):
        matching = [
            paragraph
            for paragraph in description_paragraphs
            if paragraph_text(paragraph) == expected_text
        ]
        formatted = len(matching) == 1 and bool(visible_runs(matching[0])) and all(
            run_format_matches(
                run,
                resolver=description_styles,
                size=expected_size,
                east_asia="宋体",
                latin="Times New Roman",
            )
            and not description_styles.run_boolean(run, matching[0], "b")
            for run in visible_runs(matching[0])
        )
        title_status.append(f"{expected_text}:{formatted}")
        if not formatted:
            errors.append(
                f"说明书名称或节标题 {expected_text!r} 不符合宋体不加粗规格"
            )

    abstract_filing_paragraphs = filing_paragraph_nodes(abstract_root)
    abstract_runs = [
        run for paragraph in abstract_filing_paragraphs for run in visible_runs(paragraph)
    ]
    abstract_format_ok = sum(
        run_format_matches(
            run,
            resolver=abstract_styles,
            size="24",
            east_asia="宋体",
            latin="Times New Roman",
        )
        for run in abstract_runs
    )
    abstract_paragraphs = [
        paragraph
        for paragraph in abstract_filing_paragraphs
        if paragraph_text(paragraph)
    ]
    abstract_layout_ok = all(
        paragraph_is_single_spaced_with_before_60(
            paragraph, "both", abstract_styles
        )
        for paragraph in abstract_paragraphs
    )

    table_runs = description_root.xpath("//w:tbl//w:r[w:t]", namespaces=NS)
    table_format_ok = sum(
        run_format_matches(
            run,
            resolver=description_styles,
            size="21",
            east_asia="宋体",
            latin="Times New Roman",
        )
        for run in table_runs
    )

    math_runs = claims_root.xpath("//m:r", namespaces=NS) + description_root.xpath(
        "//m:r", namespaces=NS
    )
    math_size_ok = sum(
        run.xpath("string(./w:rPr/w:sz/@w:val)", namespaces=NS) == "24"
        and run.xpath("string(./w:rPr/w:szCs/@w:val)", namespaces=NS) == "24"
        for run in math_runs
    )

    # 图题来自 patent.json 的 figures；摘要附图那份不带图号标注，故为空。
    caption_expectations = {
        "100003说明书附图.docx": list(EXPECTED_CAPTIONS),
        "摘要附图.docx": [],
    }
    caption_status: list[str] = []
    for name, expected_captions in caption_expectations.items():
        root = document_root(paths[name])
        caption_styles = StyleResolver(paths[name])
        paragraphs = filing_paragraph_nodes(root)
        actual_captions = [
            paragraph_text(paragraph)
            for paragraph in paragraphs
            if re.match(r"^图\d+", paragraph_text(paragraph))
        ]
        captions_ok = actual_captions == expected_captions
        for paragraph in paragraphs:
            if paragraph_text(paragraph) in expected_captions:
                captions_ok = captions_ok and (
                    paragraph.xpath("string(./w:pPr/w:jc/@w:val)", namespaces=NS) == "center"
                    and bool(visible_runs(paragraph))
                    and all(
                        run_format_matches(
                            run,
                            resolver=caption_styles,
                            size="21",
                            east_asia="宋体",
                            latin="Times New Roman",
                        )
                        for run in visible_runs(paragraph)
                    )
                )
        caption_status.append(f"{name}:{actual_captions},格式={captions_ok}")
        if not captions_ok:
            errors.append(f"{name} 图号不是宋体 10.5pt 居中")

    details.extend(
        [
            f"权利要求书正文 run：合格={claims_format_ok}/{len(claims_runs)}, 比例={claims_ratio:.3f}；"
            f"段前60twip、单倍、两端对齐={claims_layout_ok}/{len(claims_prose)}",
            f"说明书正文 run：sz=24且字体合格={description_format_ok}/{len(description_runs)}, "
            f"比例={description_ratio:.3f}；段落格式合格={description_layout_ok}/"
            f"{len(description_body_paragraphs)}",
            f"说明书名称及五个节标题={title_status}",
            f"摘要正文 12pt 合格={abstract_format_ok}/{len(abstract_runs)}，"
            f"段落格式={abstract_layout_ok}",
            f"表格单元格 10.5pt 合格={table_format_ok}/{len(table_runs)}",
            f"OMML 12pt 合格={math_size_ok}/{len(math_runs)}",
            f"附图图号={caption_status}",
        ]
    )
    if claims_ratio < 0.90:
        errors.append("权利要求书正文显式 12pt 宋体/Times New Roman run 少于 90%")
    if description_ratio < 0.90:
        errors.append("说明书正文 sz=24 的 run 少于 90%")
    if claims_layout_ok != len(claims_prose):
        errors.append("权利要求书正文未全部采用段前 60 twip、单倍行距、两端对齐")
    if description_layout_ok != len(description_body_paragraphs):
        errors.append("说明书正文未全部采用段前 60 twip、单倍行距、两端对齐")
    if abstract_format_ok != len(abstract_runs) or not abstract_layout_ok:
        errors.append("说明书摘要正文不是宋体 12pt 或段落格式不符")
    if table_format_ok != len(table_runs):
        errors.append("表格单元格文字不是宋体 10.5pt")
    if math_size_ok != len(math_runs):
        errors.append("OMML 公式 run 未全部显式设置为 12pt")

    results.append(
        CheckResult(
            10,
            not errors,
            "正文、标题、摘要、表格、图号、公式及段落格式检查",
            details + errors,
        )
    )

    # 11. Every required OMML argument slot must contain a non-empty expression.
    details = []
    errors = []
    slot_requirements = {
        "nary": ("sub", "sup", "e"),
        "f": ("num", "den"),
        "rad": ("e",),  # m:deg is intentionally allowed to be empty for square roots.
        "d": ("e",),
        "sSub": ("e", "sub"),
        "sSup": ("e", "sup"),
        "sSubSup": ("e", "sub", "sup"),
    }
    total_counts = {name: 0 for name in slot_requirements}
    total_empty_slots: list[str] = []
    for name, path in paths.items():
        if name not in documents:
            errors.append(f"无法执行 {name} 的 OMML 参数槽检查")
            continue
        root = document_root(path)
        file_counts: dict[str, int] = {}
        file_empty_slots: list[str] = []
        for structure_name, required_slots in slot_requirements.items():
            structures = root.xpath(f"//m:{structure_name}", namespaces=NS)
            file_counts[structure_name] = len(structures)
            total_counts[structure_name] += len(structures)
            for structure_index, structure in enumerate(structures, start=1):
                for slot_name in required_slots:
                    slots = structure.xpath(f"./m:{slot_name}", namespaces=NS)
                    if not slots:
                        file_empty_slots.append(
                            f"m:{structure_name}[{structure_index}]/m:{slot_name}:missing"
                        )
                        continue
                    for slot_index, slot in enumerate(slots, start=1):
                        slot_text = "".join(slot.itertext()).strip()
                        if not slot_text:
                            file_empty_slots.append(
                                f"m:{structure_name}[{structure_index}]/m:{slot_name}"
                                f"[{slot_index}]:empty"
                            )
        details.append(f"{name}: 结构计数={file_counts}, 空参数槽={file_empty_slots}")
        total_empty_slots.extend(f"{name}:{item}" for item in file_empty_slots)

    details.append(f"合计结构={total_counts}, 合计空参数槽={len(total_empty_slots)}")
    if total_counts["nary"] == 0:
        errors.append("未检出任何 m:nary，无法证明求和结构已检查")
    if total_empty_slots:
        errors.append(f"发现空 OMML 参数槽: {total_empty_slots}")
    results.append(
        CheckResult(
            11,
            not errors,
            "OMML n-ary、分式、根式、定界符及上下标参数槽非空检查",
            details + errors,
        )
    )

    # 12. Round-4 template fidelity, notice-tail order, indentation, clean fields,
    #     and fixed-size n-ary operators.
    details = []
    errors = []

    def body_text(node: etree._Element) -> str:
        return "".join(node.xpath(".//w:t/text()", namespaces=NS)).strip()

    def paragraph_indent_kind(paragraph: etree._Element) -> str:
        indent_nodes = paragraph.xpath("./w:pPr/w:ind", namespaces=NS)
        if len(indent_nodes) != 1:
            return "none"
        indent = indent_nodes[0]
        first_line = indent.get(f"{{{NS['w']}}}firstLine")
        first_line_chars = indent.get(f"{{{NS['w']}}}firstLineChars")
        left = indent.get(f"{{{NS['w']}}}left")
        if first_line_chars == "200":
            try:
                # Word may round the equivalent two-character value from
                # 480 to 482 twips when it saves the file.  The character
                # setting is authoritative; the small twip tolerance avoids
                # a false regression while still rejecting other indents.
                if abs(int(first_line or "0") - 480) <= 2:
                    return "first_line"
            except ValueError:
                pass
        if left == "432" and first_line == "-432":
            return "claim_hanging"
        return (
            f"other(firstLine={first_line},firstLineChars={first_line_chars},left={left})"
        )

    def page_field_has_cached_result(footer_root: etree._Element) -> bool:
        page_paragraphs = [
            paragraph
            for paragraph in footer_root.xpath(".//w:p", namespaces=NS)
            if any(
                re.search(r"\bPAGE\b", text, flags=re.IGNORECASE)
                for text in paragraph.xpath(".//w:instrText/text()", namespaces=NS)
            )
        ]
        if not page_paragraphs:
            return False
        for paragraph in page_paragraphs:
            runs = paragraph.xpath("./w:r", namespaces=NS)
            begin_indexes = [
                index
                for index, run in enumerate(runs)
                if run.xpath("./w:fldChar[@w:fldCharType='begin']", namespaces=NS)
            ]
            separate_indexes = [
                index
                for index, run in enumerate(runs)
                if run.xpath("./w:fldChar[@w:fldCharType='separate']", namespaces=NS)
            ]
            end_indexes = [
                index
                for index, run in enumerate(runs)
                if run.xpath("./w:fldChar[@w:fldCharType='end']", namespaces=NS)
            ]
            if not begin_indexes or not separate_indexes or not end_indexes:
                return False
            begin, separate, end = begin_indexes[0], separate_indexes[0], end_indexes[-1]
            if not begin < separate < end:
                return False
            result_text = "".join(
                text
                for run in runs[separate + 1 : end]
                for text in run.xpath(".//w:t/text()", namespaces=NS)
            ).strip()
            if result_text != "1":
                return False
        return True

    indentation_roots = {
        "100001权利要求书.docx": document_root(paths["100001权利要求书.docx"]),
        "100002说明书.docx": document_root(paths["100002说明书.docx"]),
    }
    for name, root in indentation_roots.items():
        paragraphs = filing_paragraph_nodes(root)
        prose_paragraphs: list[etree._Element] = []
        indent_status: list[str] = []
        matching = 0
        for paragraph in paragraphs:
            text = body_text(paragraph)
            if not text or paragraph.xpath("./m:oMathPara", namespaces=NS):
                continue
            if name == "100002说明书.docx" and text in {EXPECTED_TITLE, *EXPECTED_HEADINGS}:
                continue
            prose_paragraphs.append(paragraph)
            kind = paragraph_indent_kind(paragraph)
            # 格式基准：权利要求条款与正文一律首行缩进两字符。
            expected_kind = "first_line"
            indent_status.append(f"{kind}={expected_kind}")
            if kind == expected_kind:
                matching += 1
        ratio = matching / len(prose_paragraphs) if prose_paragraphs else 0.0
        details.append(
            f"{name}: 非标题非公式正文缩进合格={matching}/{len(prose_paragraphs)}, "
            f"覆盖率={ratio:.3f}, 类型={indent_status}"
        )
        if ratio < 0.90:
            errors.append(f"{name} 正文首行或悬挂缩进覆盖率低于 90%")

    abstract_root_for_indent = document_root(paths["100004说明书摘要.docx"])
    abstract_prose = [
        paragraph
        for paragraph in filing_paragraph_nodes(abstract_root_for_indent)
        if body_text(paragraph)
    ]
    abstract_indent_ok = (
        len(abstract_prose) == 1
        and paragraph_indent_kind(abstract_prose[0]) == "first_line"
    )
    details.append(
        f"100004说明书摘要.docx: 单段首行缩进2字符={abstract_indent_ok}"
    )
    if not abstract_indent_ok:
        errors.append("说明书摘要单段未设置 480 twip/200 字符首行缩进")

    total_nary_grow = 0
    total_dirty = 0
    for name, path in paths.items():
        if name not in documents:
            errors.append(f"无法执行 {name} 的轮4硬检查")
            continue
        with zipfile.ZipFile(path) as package:
            members = package.namelist()
            root = etree.fromstring(package.read("word/document.xml"))
            body = root.xpath("/w:document/w:body", namespaces=NS)[0]
            body_children = list(body)
            transition_indexes = [
                index
                for index, child in enumerate(body_children)
                if child.xpath("./w:pPr/w:sectPr", namespaces=NS)
            ]
            section_count = len(root.xpath("//w:sectPr", namespaces=NS))
            notice_body_text = ""
            notice_at_tail = False
            notice_keywords = {
                "100001权利要求书.docx": "提交权利要求书",
                "100002说明书.docx": "提交说明书，",
                "100003说明书附图.docx": "提交说明书附图",
                "100004说明书摘要.docx": "提交说明书摘要",
                "摘要附图.docx": "提交说明书附图",
            }
            if len(transition_indexes) == 1 and body_children[-1].tag == f"{{{NS['w']}}}sectPr":
                transition_index = transition_indexes[0]
                notice_nodes = body_children[transition_index + 1 : -1]
                notice_body_text = "".join(body_text(node) for node in notice_nodes)
                notice_at_tail = (
                    bool(notice_nodes)
                    and "申请" in notice_body_text
                    and notice_keywords[name] in notice_body_text
                )

            relationships = etree.fromstring(package.read("word/_rels/document.xml.rels"))
            relationships_by_id = {
                rel.get("Id"): rel
                for rel in relationships.xpath("./pr:Relationship", namespaces=NS)
            }
            final_sections = root.xpath("//w:sectPr", namespaces=NS)
            notice_header_text = ""
            if final_sections:
                notice_header_ids = final_sections[-1].xpath(
                    "./w:headerReference[@w:type='default']/@r:id", namespaces=NS
                )
                if len(notice_header_ids) == 1:
                    relationship = relationships_by_id.get(notice_header_ids[0])
                    if relationship is not None:
                        target = relationship.get("Target", "").lstrip("/")
                        member = target if target.startswith("word/") else f"word/{target}"
                        if member in members:
                            notice_header_root = etree.fromstring(package.read(member))
                            notice_header_text = "".join(
                                notice_header_root.xpath("//w:t/text()", namespaces=NS)
                            )

            header1_ok = False
            header1_stats = "missing"
            if "word/header1.xml" in members:
                header1_root = etree.fromstring(package.read("word/header1.xml"))
                pict_count = len(header1_root.xpath("//w:pict", namespaces=NS))
                spacing_130 = len(
                    header1_root.xpath("//w:spacing[@w:val='130']", namespaces=NS)
                )
                header1_ok = pict_count > 0 and spacing_130 > 0
                header1_stats = f"pict={pict_count},spacing130={spacing_130}"

            footer_parts = sorted(
                member
                for member in members
                if re.fullmatch(r"word/footer\d+\.xml", member)
            )
            footer_status: list[str] = []
            footer_ok = bool(footer_parts)
            for footer_member in footer_parts:
                footer_root = etree.fromstring(package.read(footer_member))
                pict_count = len(footer_root.xpath("//w:pict", namespaces=NS))
                cached_page = page_field_has_cached_result(footer_root)
                footer_status.append(
                    f"{footer_member}:pict={pict_count},PAGE结果1={cached_page}"
                )
                footer_ok = footer_ok and pict_count > 0 and cached_page

            dirty_count = 0
            for member in members:
                if not member.endswith(".xml"):
                    continue
                try:
                    xml_root = etree.fromstring(package.read(member))
                except etree.XMLSyntaxError:
                    continue
                dirty_count += sum(
                    1
                    for element in xml_root.iter()
                    for attribute, value in element.attrib.items()
                    if etree.QName(attribute).localname == "dirty" and value == "true"
                )
            nary_grow = len(root.xpath("//m:naryPr/m:grow", namespaces=NS))
            total_dirty += dirty_count
            total_nary_grow += nary_grow
            details.append(
                f"{name}: header1[{header1_stats}], footer={footer_status}, "
                f"sectPr={section_count}, 文末注意事项正文={notice_at_tail}, "
                f"末节页眉={notice_header_text!r}, dirty=true={dirty_count}, "
                f"m:naryPr/m:grow={nary_grow}"
            )
            if not header1_ok:
                errors.append(f"{name} header1 缺少 w:pict 或 w:spacing=130")
            if not footer_ok:
                errors.append(f"{name} 页脚未保留 w:pict 或 PAGE 域无预写结果 1")
            if section_count != 2 or not notice_at_tail or notice_header_text != "注意事项":
                errors.append(f"{name} 未在文末保留完整注意事项节")
            if dirty_count:
                errors.append(f"{name} 仍含 w:dirty=\"true\"")
            if nary_grow:
                errors.append(f"{name} 仍含 m:naryPr/m:grow")

    details.append(
        f"五文件合计 dirty=true={total_dirty}, m:naryPr/m:grow={total_nary_grow}"
    )
    results.append(
        CheckResult(
            12,
            not errors,
            "轮4模板VML、文末注意事项、段落缩进、域状态及求和符大小检查",
            details + errors,
        )
    )

    return results


def main(argv: list[str] | None = None) -> int:
    configure(project_dir_from_args("对五标准件 docx 做格式体检", argv))
    print(f"[PYTHON] {sys.executable}")
    print(f"[VERIFY DIR] {WORD_DIR}")
    print(f"[TITLE] {EXPECTED_TITLE}")
    results = run_checks()
    for result in results:
        status = "PASS" if result.passed else "FAIL"
        print(f"[{status}] {result.number}. {result.summary}")
        for detail in result.details:
            print(f"  - {detail}")
    passed = sum(result.passed for result in results)
    if passed == len(results):
        print(f"ALL CHECKS PASSED: {passed}/{len(results)}")
        return 0
    print(f"CHECKS FAILED: {passed}/{len(results)} passed")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
