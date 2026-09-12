#!/usr/bin/env python3
"""Build five CNIPA-style DOCX files from a Markdown patent draft.

The script intentionally starts from the supplied CNIPA DOCX templates.  It
replaces the blank filing-page body with the patent application content while
retaining the template's original first-section header/footer parts and its
second, final instruction (``注意事项``) section.

Formula route: LaTeX -> MathML (latex2mathml) -> editable OMML (this module).
"""

from __future__ import annotations

import argparse
import html as html_stdlib
import os
import re
import sys
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Twips
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree, html as lxml_html
from PIL import Image


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
from patent_config import PatentConfig, add_project_arg  # noqa: E402
from patent_config import load as load_patent_config  # noqa: E402

# 以下全部由 configure() 在运行时按 <项目目录>/patent.json 填入。
# 本模块不写死任何具体专利的名称、附图或路径。
CONFIG: PatentConfig | None = None
PROJECT_DIR: Path | None = None
SOURCE_MD: Path | None = None
TEMPLATE_DIR: Path | None = None
FIGURE_DIR: Path | None = None
OUTPUT_DIR: Path | None = None
TEMPLATE_DESCRIPTION: Path | None = None
TEMPLATE_DRAWINGS: Path | None = None
TEMPLATE_ABSTRACT: Path | None = None
ABSTRACT_FIGURE: Path | None = None
FIGURE_SPECS: tuple[tuple[Path, str], ...] = ()
FIGURES: tuple[Path, ...] = ()
OUT_CLAIMS: Path | None = None
OUT_DESCRIPTION: Path | None = None
OUT_DRAWINGS: Path | None = None
OUT_ABSTRACT: Path | None = None
OUT_ABSTRACT_DRAWING: Path | None = None


def configure(cfg: PatentConfig) -> None:
    """按 patent.json 填入本次要生成哪一篇专利。"""
    global CONFIG, PROJECT_DIR, SOURCE_MD, TEMPLATE_DIR, FIGURE_DIR, OUTPUT_DIR
    global TEMPLATE_DESCRIPTION, TEMPLATE_DRAWINGS, TEMPLATE_ABSTRACT
    global ABSTRACT_FIGURE, FIGURE_SPECS, FIGURES
    global OUT_CLAIMS, OUT_DESCRIPTION, OUT_DRAWINGS, OUT_ABSTRACT, OUT_ABSTRACT_DRAWING
    global EXPECTED_TITLE

    CONFIG = cfg
    PROJECT_DIR = cfg.project_dir
    SOURCE_MD = cfg.source_md
    TEMPLATE_DIR = cfg.template_dir
    FIGURE_DIR = cfg.figure_dir
    OUTPUT_DIR = cfg.word_dir
    TEMPLATE_DESCRIPTION = cfg.template("100002说明书.docx")
    TEMPLATE_DRAWINGS = cfg.template("100003说明书附图.docx")
    TEMPLATE_ABSTRACT = cfg.template("100004说明书摘要.docx")
    ABSTRACT_FIGURE = cfg.abstract_figure
    FIGURE_SPECS = cfg.figure_specs
    FIGURES = cfg.figures
    OUT_CLAIMS = cfg.output("100001权利要求书.docx")
    OUT_DESCRIPTION = cfg.output("100002说明书.docx")
    OUT_DRAWINGS = cfg.output("100003说明书附图.docx")
    OUT_ABSTRACT = cfg.output("100004说明书摘要.docx")
    OUT_ABSTRACT_DRAWING = cfg.output("摘要附图.docx")
    EXPECTED_TITLE = cfg.title
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 官方 100001 权利要求书模板（.doc）的注意事项六条全文；因该模板为旧版 .doc 无法作
# python-docx 底稿，权利要求书借用 100002 版式，注意事项文本须替换为本组。
CLAIMS_NOTICE_TEXTS = (
    "一、申请发明专利或者实用新型专利应当提交权利要求书，一式一份。",
    "二、权利要求书应当打字或者印刷，字迹应当整齐清晰，呈黑色，符合制版要求，不得涂改，"
    "字高应当不低于3.5毫米，行距应当在 2.5毫米至3.5毫米之间，权利要求书首页用此页，"
    "续页可使用同样大小和质量相当的白纸。纸张应当纵向使用，只限使用正面，四周应当留有"
    "页边距：左侧和顶部各25毫米，右侧和底部各15毫米。",
    "三、权利要求书应当说明发明或者实用新型的技术特征，清楚和简要地表述请求保护的范围。"
    "权利要求书有几项权利要求时，应当用阿拉伯数字顺序编号，编号前不得冠以“权利要求”"
    "或者“权项”等词。",
    "四、权利要求书中使用的科技术语应当与说明书中使用的一致，可以有化学式或者数学式，"
    "必要时可以有表格，但不得有插图。不得使用“如说明书……部分所述”或者"
    "“如图……所示”等用语。",
    "五、每一项权利要求仅允许在权利要求的结尾处使用句号。",
    "六、权利要求书应当在每页下框线居中位置顺序编写页码。",
)

EXPECTED_TITLE = ""  # 由 configure() 按 patent.json 填入
EXPECTED_DESCRIPTION_HEADINGS = ["技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式"]

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

BODY_EAST_ASIA_FONT = "宋体"
BODY_LATIN_FONT = "Times New Roman"
TITLE_FONT = "黑体"
MATH_FONT = "Cambria Math"
BODY_SIZE_HALF_POINTS = 24
TITLE_SIZE_HALF_POINTS = 28
ABSTRACT_SIZE_HALF_POINTS = 21
CAPTION_SIZE_HALF_POINTS = 21
FOOTER_SIZE_HALF_POINTS = 18
BODY_SPACE_BEFORE_TWIPS = 60
# 格式基准：固定值 24 磅行距（见 references/02-撰写规范.md 第 11 节）；独立公式段例外用
# 「最小值 24 磅」，否则带上下限的求和/分式会被固定行高裁剪。
BODY_LINE_TWIPS = 480
BODY_LINE_RULE = "exact"
FORMULA_LINE_RULE = "atLeast"
BODY_FIRST_LINE_TWIPS = 480
BODY_FIRST_LINE_CHARS = 200
CLAIM_LEFT_TWIPS = 432
CLAIM_FIRST_LINE_TWIPS = -432


def m_tag(local_name: str) -> str:
    return f"{{{M_NS}}}{local_name}"


def w_tag(local_name: str) -> str:
    return f"{{{W_NS}}}{local_name}"


def local_name(element: etree._Element) -> str:
    return etree.QName(element).localname


def append_converted(parent: etree._Element, elements: Iterable[etree._Element]) -> None:
    for element in elements:
        parent.append(element)


class MathMLToOMML:
    """Convert the MathML subset emitted by latex2mathml for this draft to OMML."""

    SUPPORTED_TAGS = {
        "math",
        "mrow",
        "mi",
        "mn",
        "mo",
        "mtext",
        "mspace",
        "mfrac",
        "msqrt",
        "msub",
        "msup",
        "msubsup",
        "munder",
        "munderover",
        "mover",
    }

    NARY_CHARS = {"∑", "∏", "∫", "∬", "∭", "⋂", "⋃"}

    def convert_latex(self, latex: str) -> etree._Element:
        mathml_text = latex_to_mathml(latex)
        mathml_root = etree.fromstring(mathml_text.encode("utf-8"))
        tags = {local_name(node) for node in mathml_root.iter()}
        unsupported = tags - self.SUPPORTED_TAGS
        if unsupported:
            raise ValueError(f"unsupported MathML tags: {sorted(unsupported)}")

        omath = etree.Element(m_tag("oMath"))
        append_converted(omath, self.convert_node(mathml_root))
        if len(omath) == 0:
            raise ValueError("formula converted to an empty OMML object")
        return omath

    def convert_node(self, node: etree._Element) -> list[etree._Element]:
        tag = local_name(node)

        if tag == "math":
            return self.convert_children(node)

        if tag == "mrow":
            children = list(node)
            if self.is_fenced_row(children):
                return [self.make_delimiter(children[0], children[-1], children[1:-1])]
            return self.convert_children(node)

        if tag in {"mi", "mn", "mo", "mtext"}:
            text = "".join(node.itertext())
            style = None
            script = None
            mathvariant = node.get("mathvariant")
            if tag in {"mn", "mo", "mtext"} or mathvariant == "normal":
                style = "p"
            if mathvariant == "sans-serif":
                style = "p"
                script = "sans-serif"
            return [self.make_run(text, style=style, script=script)]

        if tag == "mspace":
            width = node.get("width", "")
            if width.startswith("negative"):
                return []
            if width == "1em":
                return [self.make_run("\u2003", style="p")]
            return [self.make_run("\u2009", style="p")]

        children = list(node)

        if tag == "mfrac":
            if len(children) != 2:
                raise ValueError("mfrac must contain numerator and denominator")
            fraction = etree.Element(m_tag("f"))
            fraction_pr = etree.SubElement(fraction, m_tag("fPr"))
            fraction_type = etree.SubElement(fraction_pr, m_tag("type"))
            fraction_type.set(m_tag("val"), "bar")
            numerator = etree.SubElement(fraction, m_tag("num"))
            denominator = etree.SubElement(fraction, m_tag("den"))
            append_converted(numerator, self.convert_node(children[0]))
            append_converted(denominator, self.convert_node(children[1]))
            return [fraction]

        if tag == "msqrt":
            radical = etree.Element(m_tag("rad"))
            radical_pr = etree.SubElement(radical, m_tag("radPr"))
            degree_hide = etree.SubElement(radical_pr, m_tag("degHide"))
            degree_hide.set(m_tag("val"), "1")
            etree.SubElement(radical, m_tag("deg"))
            radicand = etree.SubElement(radical, m_tag("e"))
            append_converted(radicand, self.convert_children(node))
            return [radical]

        if tag == "msub":
            if len(children) != 2:
                raise ValueError("msub must contain base and subscript")
            base_text = "".join(children[0].itertext())
            if base_text in self.NARY_CHARS:
                raise ValueError("n-ary operator must be converted with its following operand sequence")
            subscript = etree.Element(m_tag("sSub"))
            base = etree.SubElement(subscript, m_tag("e"))
            sub = etree.SubElement(subscript, m_tag("sub"))
            append_converted(base, self.convert_node(children[0]))
            append_converted(sub, self.convert_node(children[1]))
            return [subscript]

        if tag == "msup":
            if len(children) != 2:
                raise ValueError("msup must contain base and superscript")
            superscript = etree.Element(m_tag("sSup"))
            base = etree.SubElement(superscript, m_tag("e"))
            sup = etree.SubElement(superscript, m_tag("sup"))
            append_converted(base, self.convert_node(children[0]))
            append_converted(sup, self.convert_node(children[1]))
            return [superscript]

        if tag == "msubsup":
            if len(children) != 3:
                raise ValueError("msubsup must contain base, subscript, and superscript")
            base_text = "".join(children[0].itertext())
            if base_text in self.NARY_CHARS:
                raise ValueError("n-ary operator must be converted with its following operand sequence")
            script = etree.Element(m_tag("sSubSup"))
            base = etree.SubElement(script, m_tag("e"))
            sub = etree.SubElement(script, m_tag("sub"))
            sup = etree.SubElement(script, m_tag("sup"))
            append_converted(base, self.convert_node(children[0]))
            append_converted(sub, self.convert_node(children[1]))
            append_converted(sup, self.convert_node(children[2]))
            return [script]

        if tag == "mover":
            if len(children) != 2:
                raise ValueError("mover must contain base and accent")
            accent = etree.Element(m_tag("acc"))
            accent_pr = etree.SubElement(accent, m_tag("accPr"))
            accent_char = etree.SubElement(accent_pr, m_tag("chr"))
            raw_accent = "".join(children[1].itertext())
            accent_char.set(m_tag("val"), "\u0302" if raw_accent == "^" else raw_accent)
            base = etree.SubElement(accent, m_tag("e"))
            append_converted(base, self.convert_node(children[0]))
            return [accent]

        if tag in {"munder", "munderover"}:
            raise ValueError("limit construct must be converted as an n-ary sequence")

        raise ValueError(f"unhandled MathML tag: {tag}")

    def convert_children(self, node: etree._Element) -> list[etree._Element]:
        return self.convert_sequence(list(node))

    def convert_sequence(self, nodes: Sequence[etree._Element]) -> list[etree._Element]:
        """Convert sibling MathML nodes while binding n-ary operands.

        MathML encodes ``\\sum`` limits in one node and leaves the summand as
        following siblings. OMML instead requires that summand inside m:nary/m:e.
        At the current grouping level a n-ary operator therefore consumes every
        remaining sibling. Recursing on that operand sequence makes a later
        nested sum consume the remainder of the outer sum's expression.
        """
        converted: list[etree._Element] = []
        for index, child in enumerate(nodes):
            if self.is_nary_limit_node(child):
                converted.append(self.make_nary_from_sequence(child, nodes[index + 1 :]))
                return converted
            converted.extend(self.convert_node(child))
        return converted

    def is_nary_limit_node(self, node: etree._Element) -> bool:
        tag = local_name(node)
        if tag not in {"msub", "msubsup", "munder", "munderover"}:
            return False
        children = list(node)
        return bool(children) and "".join(children[0].itertext()) in self.NARY_CHARS

    def make_nary_from_sequence(
        self,
        limit_node: etree._Element,
        operand_nodes: Sequence[etree._Element],
    ) -> etree._Element:
        children = list(limit_node)
        tag = local_name(limit_node)
        operator = "".join(children[0].itertext())
        if tag in {"msub", "munder"} and len(children) == 2:
            sub_node = children[1]
            sup_node = None
        elif tag in {"msubsup", "munderover"} and len(children) == 3:
            sub_node = children[1]
            sup_node = children[2]
        else:
            raise ValueError(f"malformed n-ary limit node: {tag}")
        return self.make_nary(operator, sub_node, sup_node, operand_nodes)

    @staticmethod
    def is_fenced_row(children: Sequence[etree._Element]) -> bool:
        if len(children) < 2:
            return False
        first, last = children[0], children[-1]
        return (
            local_name(first) == "mo"
            and local_name(last) == "mo"
            and first.get("fence") == "true"
            and first.get("form") == "prefix"
            and last.get("fence") == "true"
            and last.get("form") == "postfix"
        )

    def make_delimiter(
        self,
        first: etree._Element,
        last: etree._Element,
        inner: Sequence[etree._Element],
    ) -> etree._Element:
        delimiter = etree.Element(m_tag("d"))
        delimiter_pr = etree.SubElement(delimiter, m_tag("dPr"))
        begin = etree.SubElement(delimiter_pr, m_tag("begChr"))
        begin.set(m_tag("val"), "".join(first.itertext()))
        end = etree.SubElement(delimiter_pr, m_tag("endChr"))
        end.set(m_tag("val"), "".join(last.itertext()))
        grow = etree.SubElement(delimiter_pr, m_tag("grow"))
        grow.set(m_tag("val"), "1")
        expression = etree.SubElement(delimiter, m_tag("e"))
        append_converted(expression, self.convert_sequence(inner))
        return delimiter

    def make_nary(
        self,
        operator: str,
        sub_node: etree._Element,
        sup_node: etree._Element | None,
        operand_nodes: Sequence[etree._Element],
    ) -> etree._Element:
        nary = etree.Element(m_tag("nary"))
        nary_pr = etree.SubElement(nary, m_tag("naryPr"))
        char = etree.SubElement(nary_pr, m_tag("chr"))
        char.set(m_tag("val"), operator)
        limit_location = etree.SubElement(nary_pr, m_tag("limLoc"))
        limit_location.set(m_tag("val"), "undOvr")
        sub = etree.SubElement(nary, m_tag("sub"))
        sup = etree.SubElement(nary, m_tag("sup"))
        expression = etree.SubElement(nary, m_tag("e"))
        append_converted(sub, self.convert_node(sub_node))
        if sup_node is not None:
            append_converted(sup, self.convert_node(sup_node))
        else:
            sup_hide = etree.SubElement(nary_pr, m_tag("supHide"))
            sup_hide.set(m_tag("val"), "1")
        operand = self.convert_sequence(operand_nodes)
        if not operand:
            raise ValueError(f"n-ary operator {operator!r} has no operand in its MathML group")
        append_converted(expression, operand)
        return nary

    @staticmethod
    def make_run(text: str, style: str | None = None, script: str | None = None) -> etree._Element:
        run = etree.Element(m_tag("r"))
        if style is not None or script is not None:
            run_pr = etree.SubElement(run, m_tag("rPr"))
            if script is not None:
                script_el = etree.SubElement(run_pr, m_tag("scr"))
                script_el.set(m_tag("val"), script)
            if style is not None:
                style_el = etree.SubElement(run_pr, m_tag("sty"))
                style_el.set(m_tag("val"), style)
        # Word uses the embedded w:rPr as the base character size for OMML.
        # Keep it after m:rPr and before m:t to match the Office Math schema.
        word_run_pr = etree.SubElement(run, w_tag("rPr"))
        fonts = etree.SubElement(word_run_pr, w_tag("rFonts"))
        for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(w_tag(attribute), MATH_FONT)
        size = etree.SubElement(word_run_pr, w_tag("sz"))
        size.set(w_tag("val"), str(BODY_SIZE_HALF_POINTS))
        size_cs = etree.SubElement(word_run_pr, w_tag("szCs"))
        size_cs.set(w_tag("val"), str(BODY_SIZE_HALF_POINTS))
        text_el = etree.SubElement(run, m_tag("t"))
        if text[:1].isspace() or text[-1:].isspace():
            text_el.set(f"{{{XML_NS}}}space", "preserve")
        text_el.text = text
        return run


MATH_CONVERTER = MathMLToOMML()


@dataclass
class ConversionStats:
    block_success: int = 0
    inline_success: int = 0
    fallbacks: list[dict[str, str]] | None = None

    def __post_init__(self) -> None:
        if self.fallbacks is None:
            self.fallbacks = []


STATS = ConversionStats()


@dataclass(frozen=True)
class BlockToken:
    kind: str
    value: str


@dataclass(frozen=True)
class HtmlCellSpec:
    row: int
    column: int
    rowspan: int
    colspan: int
    text: str
    bold: bool


INLINE_TOKEN_RE = re.compile(
    r"(?P<bold>\*\*(?P<bold_text>.+?)\*\*)"
    r"|(?P<math>\$(?P<math_text>[^$\n]+?)\$)"
    r"|(?P<code>`(?P<code_text>.+?)`)"
)


def validate_inputs() -> None:
    required_files = [
        SOURCE_MD,
        TEMPLATE_DESCRIPTION,
        TEMPLATE_DRAWINGS,
        TEMPLATE_ABSTRACT,
        *FIGURES,
        ABSTRACT_FIGURE,
    ]
    if PROJECT_DIR is None:
        raise RuntimeError("尚未调用 configure()，请先加载 patent.json")
    if not PROJECT_DIR.is_dir():
        raise FileNotFoundError(f"项目目录不存在：{PROJECT_DIR}")
    for path in required_files:
        if not path.is_file():
            raise FileNotFoundError(f"required input file not found: {path}")
        # Read at least one byte so an unreadable placeholder fails before any write.
        with path.open("rb") as handle:
            handle.read(1)
        print(f"[INPUT] {path}")
    if not OUTPUT_DIR.parent.is_dir():
        raise FileNotFoundError(f"output parent not found: {OUTPUT_DIR.parent}")
    OUTPUT_DIR.mkdir(parents=False, exist_ok=True)
    if not OUTPUT_DIR.is_dir():
        raise NotADirectoryError(OUTPUT_DIR)
    print(f"[OUTPUT DIR] {OUTPUT_DIR}")


def section_between(markdown: str, start_heading: str, end_heading: str) -> str:
    start = markdown.find(start_heading)
    if start < 0:
        raise ValueError(f"missing heading: {start_heading}")
    start = markdown.find("\n", start)
    end = markdown.find(end_heading, start)
    if start < 0 or end < 0:
        raise ValueError(f"could not extract section {start_heading}")
    return markdown[start + 1 : end]


def content_lines(raw_section: str) -> list[str]:
    return [
        line.strip()
        for line in raw_section.splitlines()
        if line.strip()
        and not line.lstrip().startswith(">")
        and line.strip() != "---"
    ]


def parse_source(markdown: str) -> dict[str, object]:
    title_section = section_between(markdown, "## 一、发明名称", "## 二、摘要")
    abstract_section = section_between(markdown, "## 二、摘要", "## 三、权利要求书")
    claims_section = section_between(markdown, "## 三、权利要求书", "## 四、说明书")
    description_section = section_between(markdown, "## 四、说明书", "## 五、摘要图（说明）")

    title_candidates = content_lines(title_section)
    if len(title_candidates) != 1:
        raise ValueError(f"expected one invention title, got {title_candidates}")
    title = title_candidates[0]
    if title != EXPECTED_TITLE:
        raise ValueError(f"unexpected invention title: {title}")

    abstract_candidates = [
        line for line in content_lines(abstract_section) if not line.startswith("摘要附图：")
    ]
    if not abstract_candidates:
        raise ValueError("abstract body is empty")
    abstract = "".join(abstract_candidates)

    claim_numbers = [
        int(value)
        for value in re.findall(r"^\*\*(\d+)、\*\*", claims_section, flags=re.MULTILINE)
    ]
    if claim_numbers != list(range(1, 12)):
        raise ValueError(f"expected claims 1-11, got {claim_numbers}")

    description_headings = re.findall(r"^###\s+(.+?)\s*$", description_section, flags=re.MULTILINE)
    if description_headings != EXPECTED_DESCRIPTION_HEADINGS:
        raise ValueError(f"unexpected description headings: {description_headings}")

    relevant = claims_section + "\n" + description_section
    source_formula_blocks = re.findall(
        r"^\$\$\s*\n(.*?)\n\$\$\s*$", relevant, flags=re.MULTILINE | re.DOTALL
    )
    if len(source_formula_blocks) != 51:
        raise ValueError(f"expected 51 display formulas, got {len(source_formula_blocks)}")

    table_count = len(re.findall(r"^<table>\s*$", description_section, flags=re.MULTILINE))
    if table_count != 2:
        raise ValueError(f"expected two HTML tables, got {table_count}")

    return {
        "title": title,
        "abstract": abstract,
        "claims": claims_section,
        "description": description_section,
        "formula_block_count": len(source_formula_blocks),
    }


def tokenize_blocks(raw_section: str) -> list[BlockToken]:
    lines = raw_section.splitlines()
    tokens: list[BlockToken] = []
    index = 0

    def is_boundary(line: str) -> bool:
        stripped = line.strip()
        return (
            not stripped
            or line.lstrip().startswith(">")
            or stripped in {"---", "$$"}
            or stripped.startswith("<table")
            or stripped.startswith("### ")
            or stripped.startswith("## ")
        )

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped or line.lstrip().startswith(">") or stripped == "---":
            index += 1
            continue
        if stripped.startswith("## "):
            index += 1
            continue
        if stripped.startswith("### "):
            tokens.append(BlockToken("heading", stripped[4:].strip()))
            index += 1
            continue
        if stripped == "$$":
            formula_lines: list[str] = []
            index += 1
            while index < len(lines) and lines[index].strip() != "$$":
                formula_lines.append(lines[index].strip())
                index += 1
            if index >= len(lines):
                raise ValueError("unclosed display formula")
            formula = " ".join(part for part in formula_lines if part)
            if not formula:
                raise ValueError("empty display formula")
            tokens.append(BlockToken("formula", formula))
            index += 1
            continue
        if stripped.startswith("<table"):
            table_lines = [line]
            index += 1
            while index < len(lines):
                table_lines.append(lines[index])
                if lines[index].strip() == "</table>":
                    break
                index += 1
            if not table_lines[-1].strip() == "</table>":
                raise ValueError("unclosed HTML table")
            tokens.append(BlockToken("table", "\n".join(table_lines)))
            index += 1
            continue

        paragraph_lines = [stripped]
        index += 1
        while index < len(lines) and not is_boundary(lines[index]):
            paragraph_lines.append(lines[index].strip())
            index += 1
        tokens.append(BlockToken("paragraph", " ".join(paragraph_lines)))

    return tokens


@dataclass
class NoticeSectionPayload:
    """Template XML temporarily detached while generated content is written."""

    transition_paragraph: etree._Element
    notice_children: list[etree._Element]


def clean_template(template_path: Path) -> DocumentObject:
    """Prepare the filing section and retain the template notice section.

    The supplied templates store the first section properties in an empty
    transition paragraph, followed by the complete notice-page body and the
    final section properties.  ``python-docx`` always appends new content just
    before the body-level ``w:sectPr``.  We therefore detach the transition and
    notice tail, expose a temporary body-level copy of the first ``w:sectPr``
    while building, and restore the original two-section tail immediately
    before saving.
    """
    document = Document(str(template_path))
    body = document._element.body
    children = list(body)
    transition_indexes = [
        index
        for index, child in enumerate(children)
        if child.tag == w_tag("p") and child.xpath("./w:pPr/w:sectPr")
    ]
    if len(transition_indexes) != 1:
        raise ValueError(
            f"template must contain one filing/notice transition paragraph: "
            f"{template_path}, found {len(transition_indexes)}"
        )
    transition_index = transition_indexes[0]
    transition_paragraph = deepcopy(children[transition_index])
    transition_sect_pr = transition_paragraph.xpath("./w:pPr/w:sectPr")
    if len(transition_sect_pr) != 1:
        raise ValueError(f"malformed first section transition: {template_path}")
    notice_children = [deepcopy(child) for child in children[transition_index + 1 :]]
    if not notice_children or notice_children[-1].tag != w_tag("sectPr"):
        raise ValueError(f"template notice tail does not end in w:sectPr: {template_path}")
    if not any("".join(child.itertext()).strip() for child in notice_children[:-1]):
        raise ValueError(f"template notice tail contains no instruction text: {template_path}")

    retained_first_section = deepcopy(transition_sect_pr[0])
    for child in children:
        body.remove(child)
    body.append(retained_first_section)
    setattr(
        document,
        "_cnipa_notice_payload",
        NoticeSectionPayload(transition_paragraph, notice_children),
    )

    if len(document.sections) != 1:
        raise ValueError("prepared template should temporarily contain one section")
    if document.sections[0].orientation != WD_ORIENT.PORTRAIT:
        raise ValueError("CNIPA template filing page is expected to be portrait")
    return document


def restore_notice_section(document: DocumentObject) -> None:
    """Restore the template transition paragraph and final notice section."""
    payload = getattr(document, "_cnipa_notice_payload", None)
    if payload is None:
        raise ValueError("document has no detached notice-section payload")
    body = document._element.body
    direct_section_properties = body.xpath("./w:sectPr")
    if len(direct_section_properties) != 1:
        raise ValueError(
            f"prepared document should have one temporary body sectPr, got "
            f"{len(direct_section_properties)}"
        )
    body.remove(direct_section_properties[0])
    body.append(deepcopy(payload.transition_paragraph))
    for child in payload.notice_children:
        body.append(deepcopy(child))
    delattr(document, "_cnipa_notice_payload")

    section_properties = body.xpath(".//w:sectPr")
    if len(section_properties) != 2 or body[-1].tag != w_tag("sectPr"):
        raise ValueError("restored document does not contain exactly two ordered sections")


def replace_notice_texts(document: DocumentObject, texts: Sequence[str]) -> None:
    """Swap the retained notice-section body text, keeping template styling.

    Used by the claims build: its layout base is the 100002 description
    template, whose notice items must be replaced by the 100001 claim notes.
    The first numbered notice paragraph serves as the styling prototype.
    """
    payload = getattr(document, "_cnipa_notice_payload", None)
    if payload is None:
        raise ValueError("document has no detached notice-section payload")
    children = payload.notice_children
    if not children or children[-1].tag != w_tag("sectPr"):
        raise ValueError("notice payload does not end in w:sectPr")
    prototypes = [
        child
        for child in children[:-1]
        if child.tag == w_tag("p")
        and re.match(r"^[一二三四五六七八九十]、", "".join(child.itertext()).strip())
    ]
    if not prototypes:
        raise ValueError("no numbered notice paragraph available as styling prototype")
    prototype = prototypes[0]
    prototype_ppr = prototype.find(qn("w:pPr"))
    prototype_rpr = None
    for run in prototype.findall(qn("w:r")):
        candidate = run.find(qn("w:rPr"))
        if run.find(qn("w:t")) is not None:
            prototype_rpr = candidate
            break
    new_children: list[etree._Element] = []
    for text in texts:
        paragraph = OxmlElement("w:p")
        if prototype_ppr is not None:
            paragraph.append(deepcopy(prototype_ppr))
        run = OxmlElement("w:r")
        if prototype_rpr is not None:
            run.append(deepcopy(prototype_rpr))
        text_el = OxmlElement("w:t")
        text_el.set(f"{{{XML_NS}}}space", "preserve")
        text_el.text = text
        run.append(text_el)
        paragraph.append(run)
        new_children.append(paragraph)
    new_children.append(children[-1])
    payload.notice_children = new_children


def get_or_add_child(parent: etree._Element, qualified_name: str) -> etree._Element:
    child = parent.find(qn(qualified_name))
    if child is None:
        child = OxmlElement(qualified_name)
        parent.append(child)
    return child


def apply_run_properties(
    run_properties: etree._Element,
    *,
    east_asia_font: str,
    latin_font: str,
    size_half_points: int,
) -> None:
    fonts = get_or_add_child(run_properties, "w:rFonts")
    fonts.set(qn("w:ascii"), latin_font)
    fonts.set(qn("w:hAnsi"), latin_font)
    fonts.set(qn("w:eastAsia"), east_asia_font)
    fonts.set(qn("w:cs"), latin_font)
    size = get_or_add_child(run_properties, "w:sz")
    size.set(qn("w:val"), str(size_half_points))
    size_cs = get_or_add_child(run_properties, "w:szCs")
    size_cs.set(qn("w:val"), str(size_half_points))


def format_run(
    run,
    *,
    east_asia_font: str = BODY_EAST_ASIA_FONT,
    latin_font: str = BODY_LATIN_FONT,
    size_half_points: int = BODY_SIZE_HALF_POINTS,
) -> None:
    run_properties = run._element.get_or_add_rPr()
    apply_run_properties(
        run_properties,
        east_asia_font=east_asia_font,
        latin_font=latin_font,
        size_half_points=size_half_points,
    )


def format_paragraph(
    paragraph,
    *,
    alignment,
    east_asia_font: str,
    latin_font: str,
    size_half_points: int,
    space_before_twips: int = BODY_SPACE_BEFORE_TWIPS,
    indent_kind: str | None = None,
    line_twips: int | None = BODY_LINE_TWIPS,
    line_rule: str | None = None,
) -> None:
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = Twips(space_before_twips)
    paragraph.paragraph_format.space_after = Twips(0)
    paragraph_properties = paragraph._p.get_or_add_pPr()
    spacing = get_or_add_child(paragraph_properties, "w:spacing")
    if line_twips is None:
        # 表格单元格等紧凑场景：单倍行距。
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
    else:
        spacing.set(qn("w:line"), str(line_twips))
        spacing.set(qn("w:lineRule"), line_rule or BODY_LINE_RULE)
    default_run_properties = get_or_add_child(paragraph_properties, "w:rPr")
    apply_run_properties(
        default_run_properties,
        east_asia_font=east_asia_font,
        latin_font=latin_font,
        size_half_points=size_half_points,
    )
    indentation = get_or_add_child(paragraph_properties, "w:ind")
    for attribute_name in (
        "left",
        "right",
        "start",
        "end",
        "firstLine",
        "firstLineChars",
        "hanging",
        "hangingChars",
    ):
        indentation.attrib.pop(qn(f"w:{attribute_name}"), None)
    if indent_kind == "first_line":
        indentation.set(qn("w:firstLine"), str(BODY_FIRST_LINE_TWIPS))
        indentation.set(qn("w:firstLineChars"), str(BODY_FIRST_LINE_CHARS))
    elif indent_kind == "claim_hanging":
        indentation.set(qn("w:left"), str(CLAIM_LEFT_TWIPS))
        indentation.set(qn("w:firstLine"), str(CLAIM_FIRST_LINE_TWIPS))
    elif indent_kind is not None:
        raise ValueError(f"unknown paragraph indentation kind: {indent_kind}")
    if not indentation.attrib:
        paragraph_properties.remove(indentation)


def append_page_field(paragraph_element: etree._Element) -> None:
    def append_field_run(child: etree._Element) -> None:
        run = OxmlElement("w:r")
        run_properties = OxmlElement("w:rPr")
        apply_run_properties(
            run_properties,
            east_asia_font=BODY_EAST_ASIA_FONT,
            latin_font=BODY_LATIN_FONT,
            size_half_points=FOOTER_SIZE_HALF_POINTS,
        )
        run.append(run_properties)
        run.append(child)
        paragraph_element.append(run)

    field_begin = OxmlElement("w:fldChar")
    field_begin.set(qn("w:fldCharType"), "begin")
    append_field_run(field_begin)

    instruction = OxmlElement("w:instrText")
    instruction.set(f"{{{XML_NS}}}space", "preserve")
    instruction.text = " PAGE \\* Arabic "
    append_field_run(instruction)

    field_separate = OxmlElement("w:fldChar")
    field_separate.set(qn("w:fldCharType"), "separate")
    append_field_run(field_separate)

    result = OxmlElement("w:t")
    result.text = "1"
    append_field_run(result)

    field_end = OxmlElement("w:fldChar")
    field_end.set(qn("w:fldCharType"), "end")
    append_field_run(field_end)


def replace_text_across_nodes(
    root: etree._Element,
    old_text: str,
    new_text: str,
    *,
    context: str,
) -> None:
    """Replace package-part text without rebuilding any surrounding XML."""
    text_nodes = root.xpath(".//w:t")
    combined = "".join(node.text or "" for node in text_nodes)
    if combined == new_text:
        return
    if combined.count(old_text) != 1:
        raise ValueError(
            f"{context}: expected one {old_text!r} in visible text {combined!r}"
        )
    replaced = combined.replace(old_text, new_text, 1)
    original_lengths = [len(node.text or "") for node in text_nodes]
    cursor = 0
    for index, node in enumerate(text_nodes):
        if index == len(text_nodes) - 1:
            segment = replaced[cursor:]
        else:
            segment = replaced[cursor : cursor + original_lengths[index]]
        node.text = segment
        cursor += len(segment)


def related_xml_part(
    document: DocumentObject,
    section_properties: etree._Element,
    reference_tag: str,
    relationship_type: str,
):
    references = section_properties.xpath(f"./w:{reference_tag}[@w:type='default']")
    if len(references) != 1:
        raise ValueError(
            f"first section must have one default w:{reference_tag}, got {len(references)}"
        )
    relationship_id = references[0].get(qn("r:id"))
    relationship = document.part.rels.get(relationship_id)
    if relationship is None or relationship.reltype != relationship_type:
        raise ValueError(f"invalid {reference_tag} relationship: {relationship_id}")
    return relationship.target_part


def prepare_footer_page_field(footer_element: etree._Element) -> None:
    """Use the template's trailing blank paragraph for a centered PAGE field."""
    page_instructions = footer_element.xpath(".//w:instrText")
    if any(re.search(r"\bPAGE\b", node.text or "", flags=re.IGNORECASE) for node in page_instructions):
        for dirty_node in footer_element.xpath(".//*[@w:dirty='true']"):
            dirty_node.attrib.pop(qn("w:dirty"), None)
        return

    candidates = [
        paragraph
        for paragraph in footer_element.xpath("./w:p")
        if not "".join(paragraph.xpath(".//w:t/text()"))
        and not paragraph.xpath(".//w:pict|.//w:drawing|.//w:fldChar")
    ]
    if candidates:
        paragraph = candidates[-1]
    else:
        paragraph = OxmlElement("w:p")
        footer_element.append(paragraph)
    paragraph_properties = paragraph.find(qn("w:pPr"))
    if paragraph_properties is None:
        paragraph_properties = OxmlElement("w:pPr")
        paragraph.insert(0, paragraph_properties)
    justification = get_or_add_child(paragraph_properties, "w:jc")
    justification.set(qn("w:val"), "center")
    spacing = get_or_add_child(paragraph_properties, "w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    append_page_field(paragraph)


def configure_header_footer(
    document: DocumentObject,
    *,
    header_text: str,
    form_code: str,
) -> None:
    if len(document.sections) != 1:
        raise ValueError("template part configuration requires the prepared first section")
    section_properties = document._element.body.sectPr
    if section_properties is None:
        raise ValueError("prepared template has no first section properties")

    header_part = related_xml_part(
        document, section_properties, "headerReference", RT.HEADER
    )
    header_element = header_part._element
    current_header_text = "".join(header_element.xpath(".//w:t/text()"))
    replace_text_across_nodes(
        header_element,
        current_header_text,
        header_text,
        context=f"{header_part.partname} filing header",
    )
    if not header_element.xpath(".//w:pict"):
        raise ValueError(f"template filing header lost w:pict: {header_part.partname}")
    if not header_element.xpath(".//w:b"):
        raise ValueError(f"template filing header lost w:b: {header_part.partname}")
    header_spacings = header_element.xpath(".//w:rPr/w:spacing")
    if not header_spacings:
        raise ValueError(f"template filing header has no character spacing: {header_part.partname}")
    # The round-4 acceptance contract explicitly requires 130 twentieths of a
    # point in every filing header.  Preserve every copied run and VML object,
    # changing only this existing spacing attribute where a source template
    # uses its legacy 89/90 value.
    for spacing in header_spacings:
        spacing.set(qn("w:val"), "130")

    footer_part = related_xml_part(
        document, section_properties, "footerReference", RT.FOOTER
    )
    footer_element = footer_part._element
    footer_text = "".join(footer_element.xpath(".//w:t/text()"))
    existing_codes = re.findall(r"100\d{3}", footer_text)
    if len(existing_codes) != 1:
        raise ValueError(
            f"{footer_part.partname}: expected one form code, got {existing_codes}"
        )
    replace_text_across_nodes(
        footer_element,
        existing_codes[0],
        form_code,
        context=f"{footer_part.partname} filing form code",
    )
    if not footer_element.xpath(".//w:pict"):
        raise ValueError(f"template filing footer lost w:pict: {footer_part.partname}")

    footer_parts = {
        relationship.target_part.partname: relationship.target_part
        for relationship in document.part.rels.values()
        if relationship.reltype == RT.FOOTER
    }
    for part in footer_parts.values():
        if not part._element.xpath(".//w:pict"):
            raise ValueError(f"template footer lost w:pict: {part.partname}")
        prepare_footer_page_field(part._element)


def validate_template_header_footer_parts(document: DocumentObject) -> None:
    """Ensure both template sections and all referenced parts remain intact."""
    section_properties = document._element.xpath("//w:sectPr")
    if len(section_properties) != 2:
        raise ValueError(f"expected two template sectPr elements, got {len(section_properties)}")
    header_ids = document._element.xpath("//w:headerReference/@r:id")
    footer_ids = document._element.xpath("//w:footerReference/@r:id")
    if len(header_ids) != 2 or not footer_ids:
        raise ValueError(
            f"expected two header references and at least one footer reference, "
            f"got headers={header_ids}, footers={footer_ids}"
        )
    for relationship_id in set(header_ids + footer_ids):
        if relationship_id not in document.part.rels:
            raise ValueError(f"missing retained header/footer relationship: {relationship_id}")


def add_plain_run(
    paragraph,
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    east_asia_font: str = BODY_EAST_ASIA_FONT,
    latin_font: str = BODY_LATIN_FONT,
    size_half_points: int = BODY_SIZE_HALF_POINTS,
):
    if not text:
        return None
    run = paragraph.add_run(html_stdlib.unescape(text))
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    format_run(
        run,
        east_asia_font=east_asia_font,
        latin_font=latin_font,
        size_half_points=size_half_points,
    )
    return run


def add_rich_text(
    paragraph,
    markdown_text: str,
    context: str,
    *,
    east_asia_font: str = BODY_EAST_ASIA_FONT,
    latin_font: str = BODY_LATIN_FONT,
    size_half_points: int = BODY_SIZE_HALF_POINTS,
) -> None:
    position = 0
    for match in INLINE_TOKEN_RE.finditer(markdown_text):
        add_plain_run(
            paragraph,
            markdown_text[position : match.start()],
            east_asia_font=east_asia_font,
            latin_font=latin_font,
            size_half_points=size_half_points,
        )
        if match.group("bold") is not None:
            add_plain_run(
                paragraph,
                match.group("bold_text"),
                bold=True,
                east_asia_font=east_asia_font,
                latin_font=latin_font,
                size_half_points=size_half_points,
            )
        elif match.group("math") is not None:
            latex = match.group("math_text")
            try:
                paragraph._p.append(MATH_CONVERTER.convert_latex(latex))
                STATS.inline_success += 1
            except Exception as exc:  # explicit, reported degradation path
                add_plain_run(
                    paragraph,
                    latex,
                    italic=True,
                    east_asia_font=east_asia_font,
                    latin_font=latin_font,
                    size_half_points=size_half_points,
                )
                assert STATS.fallbacks is not None
                STATS.fallbacks.append(
                    {"kind": "inline", "latex": latex, "context": context, "error": str(exc)}
                )
        else:
            add_plain_run(
                paragraph,
                match.group("code_text"),
                east_asia_font=east_asia_font,
                latin_font=latin_font,
                size_half_points=size_half_points,
            )
        position = match.end()
    add_plain_run(
        paragraph,
        markdown_text[position:],
        east_asia_font=east_asia_font,
        latin_font=latin_font,
        size_half_points=size_half_points,
    )


def add_body_paragraph(
    document: DocumentObject,
    text: str,
    context: str,
    *,
    centered: bool = False,
    size_half_points: int = BODY_SIZE_HALF_POINTS,
    indent_kind: str = "first_line",
):
    paragraph = document.add_paragraph()
    format_paragraph(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.JUSTIFY,
        east_asia_font=BODY_EAST_ASIA_FONT,
        latin_font=BODY_LATIN_FONT,
        size_half_points=size_half_points,
        space_before_twips=BODY_SPACE_BEFORE_TWIPS,
        indent_kind=indent_kind,
    )
    add_rich_text(
        paragraph,
        text,
        context,
        east_asia_font=BODY_EAST_ASIA_FONT,
        latin_font=BODY_LATIN_FONT,
        size_half_points=size_half_points,
    )
    return paragraph


def add_heading(document: DocumentObject, text: str):
    # 格式基准：节标题与正文完全同样式（宋体 12pt、不加粗、首行缩进两字符、两端对齐）。
    paragraph = document.add_paragraph()
    format_paragraph(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        east_asia_font=BODY_EAST_ASIA_FONT,
        latin_font=BODY_LATIN_FONT,
        size_half_points=BODY_SIZE_HALF_POINTS,
        space_before_twips=BODY_SPACE_BEFORE_TWIPS,
        indent_kind="first_line",
    )
    paragraph.paragraph_format.keep_with_next = True
    run = paragraph.add_run(text)
    format_run(
        run,
        east_asia_font=BODY_EAST_ASIA_FONT,
        latin_font=BODY_LATIN_FONT,
        size_half_points=BODY_SIZE_HALF_POINTS,
    )
    return paragraph


def add_display_formula(document: DocumentObject, latex: str, context: str) -> bool:
    try:
        omath = MATH_CONVERTER.convert_latex(latex)
        math_paragraph = etree.Element(m_tag("oMathPara"))
        math_paragraph_pr = etree.SubElement(math_paragraph, m_tag("oMathParaPr"))
        justification = etree.SubElement(math_paragraph_pr, m_tag("jc"))
        justification.set(m_tag("val"), "center")
        math_paragraph.append(omath)
        # m:oMathPara is paragraph content, not a valid direct child of w:body.
        # Wrapping it in a centered w:p is required for strict Microsoft Word
        # OOXML validation, even though python-docx can reopen the invalid form.
        word_paragraph = document.add_paragraph()
        format_paragraph(
            word_paragraph,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            east_asia_font=BODY_EAST_ASIA_FONT,
            latin_font=BODY_LATIN_FONT,
            size_half_points=BODY_SIZE_HALF_POINTS,
            space_before_twips=BODY_SPACE_BEFORE_TWIPS,
            line_rule=FORMULA_LINE_RULE,
        )
        word_paragraph._p.append(math_paragraph)
        STATS.block_success += 1
        return True
    except Exception as exc:  # explicit, reported degradation path
        paragraph = document.add_paragraph()
        format_paragraph(
            paragraph,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            east_asia_font=BODY_EAST_ASIA_FONT,
            latin_font=BODY_LATIN_FONT,
            size_half_points=BODY_SIZE_HALF_POINTS,
            space_before_twips=BODY_SPACE_BEFORE_TWIPS,
            line_rule=FORMULA_LINE_RULE,
        )
        add_plain_run(paragraph, latex, italic=True)
        assert STATS.fallbacks is not None
        STATS.fallbacks.append(
            {"kind": "display", "latex": latex, "context": context, "error": str(exc)}
        )
        return False


def parse_html_table(table_html: str) -> tuple[int, int, list[HtmlCellSpec]]:
    table_element = lxml_html.fragment_fromstring(table_html)
    rows = table_element.xpath(".//tr")
    if not rows:
        raise ValueError("HTML table contains no rows")

    occupied: set[tuple[int, int]] = set()
    specs: list[HtmlCellSpec] = []
    max_columns = 0
    for row_index, row in enumerate(rows):
        column_index = 0
        for cell in row.xpath("./th|./td"):
            while (row_index, column_index) in occupied:
                column_index += 1
            rowspan = int(cell.get("rowspan", "1"))
            colspan = int(cell.get("colspan", "1"))
            text = " ".join("".join(cell.itertext()).split())
            bold = local_name(cell) == "th" or bool(cell.xpath(".//strong"))
            specs.append(
                HtmlCellSpec(
                    row=row_index,
                    column=column_index,
                    rowspan=rowspan,
                    colspan=colspan,
                    text=text,
                    bold=bold,
                )
            )
            for occupied_row in range(row_index, row_index + rowspan):
                for occupied_column in range(column_index, column_index + colspan):
                    occupied.add((occupied_row, occupied_column))
            column_index += colspan
            max_columns = max(max_columns, column_index)

    for row_index in range(len(rows)):
        covered = {column for row, column in occupied if row == row_index}
        if covered != set(range(max_columns)):
            raise ValueError(f"HTML table row {row_index} does not cover {max_columns} columns")
    return len(rows), max_columns, specs


def set_cell_margins(cell, top: int = 40, start: int = 40, bottom: int = 40, end: int = 40) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin_name, margin_value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin_name}"))
        if node is None:
            node = OxmlElement(f"w:{margin_name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(margin_value))
        node.set(qn("w:type"), "dxa")


def apply_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge_el = OxmlElement(f"w:{edge}")
        edge_el.set(qn("w:val"), "single")
        edge_el.set(qn("w:sz"), "4")
        edge_el.set(qn("w:space"), "0")
        edge_el.set(qn("w:color"), "000000")
        borders.append(edge_el)


def mark_header_row_repeat(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_html_table(document: DocumentObject, table_html: str, table_number: int) -> None:
    row_count, column_count, specs = parse_html_table(table_html)
    if row_count != 9 or column_count != 13:
        raise ValueError(
            f"table {table_number}: expected 9 rows and 13 columns, got {row_count}x{column_count}"
        )

    table = document.add_table(rows=row_count, cols=column_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    section = document.sections[0]
    available_width = section.page_width - section.left_margin - section.right_margin
    first_column_width = int(Mm(22))
    data_column_width = int((available_width - first_column_width) / 12)
    widths = [first_column_width] + [data_column_width] * 12
    for column_index, width in enumerate(widths):
        table.columns[column_index].width = width
        for row in table.rows:
            row.cells[column_index].width = width

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    table._tbl.tblPr.append(layout)
    apply_table_borders(table)

    for spec in specs:
        if spec.rowspan > 1 or spec.colspan > 1:
            table.cell(spec.row, spec.column).merge(
                table.cell(spec.row + spec.rowspan - 1, spec.column + spec.colspan - 1)
            )

    for spec in specs:
        cell = table.cell(spec.row, spec.column)
        cell.text = ""
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell)
        paragraph = cell.paragraphs[0]
        format_paragraph(
            paragraph,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            east_asia_font=BODY_EAST_ASIA_FONT,
            latin_font=BODY_LATIN_FONT,
            size_half_points=ABSTRACT_SIZE_HALF_POINTS,
            space_before_twips=0,
            line_twips=None,
        )
        run = paragraph.add_run(spec.text)
        run.bold = spec.bold
        format_run(
            run,
            east_asia_font=BODY_EAST_ASIA_FONT,
            latin_font=BODY_LATIN_FONT,
            size_half_points=ABSTRACT_SIZE_HALF_POINTS,
        )

    mark_header_row_repeat(table.rows[0])
    mark_header_row_repeat(table.rows[1])


def add_picture_scaled(document: DocumentObject, image_path: Path, label: str | None = None) -> None:
    section = document.sections[0]
    available_width = section.page_width - section.left_margin - section.right_margin - int(Mm(4))
    available_height = section.page_height - section.top_margin - section.bottom_margin - int(Mm(18))
    with Image.open(image_path) as image:
        pixel_width, pixel_height = image.size
    target_width = int(available_width)
    target_height = int(target_width * pixel_height / pixel_width)
    if target_height > available_height:
        target_height = int(available_height)
        target_width = int(target_height * pixel_width / pixel_height)

    paragraph = document.add_paragraph()
    format_paragraph(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        east_asia_font=BODY_EAST_ASIA_FONT,
        latin_font=BODY_LATIN_FONT,
        size_half_points=CAPTION_SIZE_HALF_POINTS,
        space_before_twips=0,
        line_rule=FORMULA_LINE_RULE,
    )
    paragraph.paragraph_format.keep_with_next = label is not None
    run = paragraph.add_run()
    run.add_picture(str(image_path), width=target_width, height=target_height)

    if label is not None:
        caption = document.add_paragraph()
        format_paragraph(
            caption,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            east_asia_font=BODY_EAST_ASIA_FONT,
            latin_font=BODY_LATIN_FONT,
            size_half_points=CAPTION_SIZE_HALF_POINTS,
            space_before_twips=0,
        )
        caption_run = caption.add_run(label)
        format_run(
            caption_run,
            east_asia_font=BODY_EAST_ASIA_FONT,
            latin_font=BODY_LATIN_FONT,
            size_half_points=CAPTION_SIZE_HALF_POINTS,
        )


def save_atomic(document: DocumentObject, destination: Path) -> None:
    if destination.parent.resolve() != OUTPUT_DIR.resolve():
        raise ValueError(f"refusing to write outside output directory: {destination}")
    restore_notice_section(document)
    validate_template_header_footer_parts(document)
    if document._element.nsmap.get("m") != M_NS:
        raise ValueError("document root does not declare the required OMML math namespace")
    temporary = destination.with_name(f".{destination.stem}.tmp.docx")
    if temporary.exists():
        temporary.unlink()
    try:
        document.save(str(temporary))
        # Immediate reopen catches corrupt packages before replacing a prior good result.
        Document(str(temporary))
        os.replace(temporary, destination)
    finally:
        if temporary.exists():
            temporary.unlink()
    print(f"[WROTE] {destination}")


def build_claims(parsed: dict[str, object]) -> int:
    document = clean_template(TEMPLATE_DESCRIPTION)
    replace_notice_texts(document, CLAIMS_NOTICE_TEXTS)
    configure_header_footer(document, header_text="权利要求书", form_code="100001")
    document.core_properties.title = "权利要求书"
    tokens = tokenize_blocks(str(parsed["claims"]))
    block_count = 0
    for token in tokens:
        context = f"100001权利要求书:{token.kind}"
        if token.kind == "paragraph":
            # 格式基准：权利要求条款与普通段落一样首行缩进两字符（悬挂缩进仅见于注意事项页）。
            add_body_paragraph(
                document,
                token.value,
                context,
                indent_kind="first_line",
            )
        elif token.kind == "formula":
            add_display_formula(document, token.value, context)
            block_count += 1
        elif token.kind == "table":
            raise ValueError("claims unexpectedly contain a table")
        elif token.kind == "heading":
            raise ValueError("claims unexpectedly contain a subsection heading")
    save_atomic(document, OUT_CLAIMS)
    return block_count


def build_description(parsed: dict[str, object]) -> int:
    document = clean_template(TEMPLATE_DESCRIPTION)
    configure_header_footer(document, header_text="说明书", form_code="100002")
    document.core_properties.title = "说明书"

    # 格式基准：发明名称宋体 14pt、不加粗、居中。
    title_paragraph = document.add_paragraph()
    format_paragraph(
        title_paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        east_asia_font=BODY_EAST_ASIA_FONT,
        latin_font=BODY_LATIN_FONT,
        size_half_points=TITLE_SIZE_HALF_POINTS,
        space_before_twips=0,
    )
    title_run = title_paragraph.add_run(str(parsed["title"]))
    format_run(
        title_run,
        east_asia_font=BODY_EAST_ASIA_FONT,
        latin_font=BODY_LATIN_FONT,
        size_half_points=TITLE_SIZE_HALF_POINTS,
    )
    document.add_paragraph()  # CNIPA template instruction: one blank line after the title.

    tokens = tokenize_blocks(str(parsed["description"]))
    block_count = 0
    table_count = 0
    for token in tokens:
        context = f"100002说明书:{token.kind}"
        if token.kind == "heading":
            add_heading(document, token.value)
        elif token.kind == "paragraph":
            centered = token.value.startswith("表一 ") or token.value.startswith("表二 ")
            paragraph = add_body_paragraph(document, token.value, context, centered=centered)
            if centered:
                paragraph.paragraph_format.keep_with_next = True
        elif token.kind == "formula":
            add_display_formula(document, token.value, context)
            block_count += 1
        elif token.kind == "table":
            table_count += 1
            add_html_table(document, token.value, table_count)

    if table_count != 2:
        raise ValueError(f"description build expected two tables, got {table_count}")
    save_atomic(document, OUT_DESCRIPTION)
    return block_count


def build_drawings() -> None:
    document = clean_template(TEMPLATE_DRAWINGS)
    configure_header_footer(document, header_text="说明书附图", form_code="100003")
    document.core_properties.title = "说明书附图"
    for index, (figure, label) in enumerate(FIGURE_SPECS, start=1):
        add_picture_scaled(document, figure, label=label)
        if index != len(FIGURE_SPECS):
            document.add_page_break()
    save_atomic(document, OUT_DRAWINGS)


def build_abstract(parsed: dict[str, object]) -> None:
    document = clean_template(TEMPLATE_ABSTRACT)
    configure_header_footer(document, header_text="说明书摘要", form_code="100004")
    document.core_properties.title = "说明书摘要"
    # 格式基准：摘要正文与说明书正文同为宋体 12pt。
    add_body_paragraph(
        document,
        str(parsed["abstract"]),
        "100004说明书摘要:paragraph",
        size_half_points=BODY_SIZE_HALF_POINTS,
    )
    save_atomic(document, OUT_ABSTRACT)


def build_abstract_drawing() -> None:
    document = clean_template(TEMPLATE_DRAWINGS)
    # Per the v1 baseline, the abstract drawing reuses the 100003 drawing form.
    # Whether the agency ultimately wants a dedicated abstract-drawing form is
    # a filing-policy question and is recorded in the delivery report.
    configure_header_footer(document, header_text="说明书附图", form_code="100003")
    document.core_properties.title = "摘要附图"
    # 格式基准：摘要附图文件只放指定图（图4），不带图号标注。
    add_picture_scaled(document, ABSTRACT_FIGURE, label=None)
    save_atomic(document, OUT_ABSTRACT_DRAWING)


def count_claims(claims_source: str) -> int:
    """数权利要求条数：行首的「N、」编号，允许 Markdown 加粗包裹。"""
    return len(re.findall(r"^\s*(?:\*\*)?\d+、", claims_source, flags=re.MULTILINE))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按 patent.json 从 md 草案生成 CNIPA 五标准件 docx"
    )
    add_project_arg(parser)
    args = parser.parse_args(argv)
    configure(load_patent_config(args.project))

    print(f"[PYTHON] {sys.executable}")
    print(f"[PROJECT] {PROJECT_DIR}")
    print(f"[TITLE] {EXPECTED_TITLE}（{len(EXPECTED_TITLE)} 字）")
    validate_inputs()
    markdown = SOURCE_MD.read_text(encoding="utf-8")
    parsed = parse_source(markdown)
    claims_formula_blocks = build_claims(parsed)
    description_formula_blocks = build_description(parsed)
    build_drawings()
    build_abstract(parsed)
    build_abstract_drawing()

    total_formula_blocks = claims_formula_blocks + description_formula_blocks
    print(
        f"[SOURCE] claims={count_claims(parsed['claims'])}, "
        f"display formulas={parsed['formula_block_count']}, "
        f"figures={len(FIGURE_SPECS)}"
    )
    print(
        "[FORMULAS] "
        f"display blocks emitted={total_formula_blocks}, "
        f"display OMML successes={STATS.block_success}, "
        f"inline OMML successes={STATS.inline_success}"
    )
    if STATS.fallbacks:
        print(f"[FORMULA FALLBACKS] {len(STATS.fallbacks)}")
        for fallback in STATS.fallbacks:
            print(
                "  - "
                f"{fallback['kind']} {fallback['latex']!r} at {fallback['context']}: "
                f"{fallback['error']}"
            )
    else:
        print("[FORMULA FALLBACKS] none")

    if total_formula_blocks != parsed["formula_block_count"]:
        raise ValueError(
            f"display formula token mismatch: built {total_formula_blocks}, "
            f"source {parsed['formula_block_count']}"
        )
    print("BUILD COMPLETED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
