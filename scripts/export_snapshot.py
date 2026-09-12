#!/usr/bin/env python3
"""Export deterministic Markdown snapshots from the five editable DOCX files."""

from __future__ import annotations

import posixpath
import sys
import zipfile
from pathlib import Path, PurePosixPath

from docx import Document
from docx.table import Table
from lxml import etree


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
from patent_config import project_dir_from_args  # noqa: E402

# 由 configure() 在运行时填入。
PROJECT_DIR: Path | None = None
WORD_DIR: Path | None = None
SNAPSHOT_DIR: Path | None = None


def configure(project_dir: Path) -> None:
    global PROJECT_DIR, WORD_DIR, SNAPSHOT_DIR
    PROJECT_DIR = project_dir
    WORD_DIR = project_dir / "word"
    SNAPSHOT_DIR = project_dir / "snapshots"

DOCX_NAMES = [
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
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

W_P = f"{{{NS['w']}}}p"
W_TBL = f"{{{NS['w']}}}tbl"
M_OMATH = f"{{{NS['m']}}}oMath"
M_OMATH_PARA = f"{{{NS['m']}}}oMathPara"
MC_ALTERNATE_CONTENT = f"{{{NS['mc']}}}AlternateContent"
MC_CHOICE = f"{{{NS['mc']}}}Choice"
MC_FALLBACK = f"{{{NS['mc']}}}Fallback"


def xpath(node: etree._Element, expression: str):
    """Run XPath on either an lxml element or a python-docx OXML element."""
    try:
        return node.xpath(expression, namespaces=NS)
    except TypeError:
        return node.xpath(expression)


class SnapshotExporter:
    """Read one DOCX package and render its document body in source order."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.document = Document(str(path))
        with zipfile.ZipFile(path) as package:
            self.members = set(package.namelist())
            relationships = etree.fromstring(
                package.read("word/_rels/document.xml.rels")
            )
        self.image_names = self._image_relationships(relationships)

    def _image_relationships(
        self, relationships: etree._Element
    ) -> dict[str, str]:
        image_names: dict[str, str] = {}
        for relationship in relationships.xpath(
            "./pr:Relationship", namespaces=NS
        ):
            if not relationship.get("Type", "").endswith("/image"):
                continue
            relationship_id = relationship.get("Id")
            target = relationship.get("Target", "").replace("\\", "/")
            if not relationship_id or not target:
                continue
            if target.startswith("/"):
                member = target.lstrip("/")
            else:
                member = posixpath.normpath(posixpath.join("word", target))
            if member in self.members:
                image_names[relationship_id] = PurePosixPath(member).name
        return image_names

    @staticmethod
    def _formula_text(node: etree._Element) -> str:
        return "".join(xpath(node, ".//m:t/text()"))

    def _image_marker(self, node: etree._Element) -> str:
        if node.tag == f"{{{NS['a']}}}blip":
            candidates = [
                node.get(f"{{{NS['r']}}}embed"),
                node.get(f"{{{NS['r']}}}link"),
            ]
        elif node.tag == f"{{{NS['v']}}}imagedata":
            candidates = [node.get(f"{{{NS['r']}}}id")]
        else:
            return ""
        return "".join(
            f"![图片](media/{self.image_names[relationship_id]})"
            for relationship_id in candidates
            if relationship_id in self.image_names
        )

    def _inline_text(self, node: etree._Element) -> str:
        if node.tag in {M_OMATH, M_OMATH_PARA}:
            return f"${self._formula_text(node)}$"

        image_marker = self._image_marker(node)
        if image_marker:
            return image_marker

        if node.tag == f"{{{NS['w']}}}t":
            return (node.text or "").replace("\r", " ").replace("\n", " ")
        if node.tag == f"{{{NS['w']}}}tab":
            return "\t"
        if node.tag in {
            f"{{{NS['w']}}}br",
            f"{{{NS['w']}}}cr",
        }:
            return " "
        if node.tag == f"{{{NS['w']}}}noBreakHyphen":
            return "‑"
        if node.tag == f"{{{NS['w']}}}softHyphen":
            return "\u00ad"
        if node.tag == f"{{{NS['w']}}}sym":
            value = node.get(f"{{{NS['w']}}}char", "")
            try:
                return chr(int(value, 16))
            except ValueError:
                return ""

        # AlternateContent carries two equivalent representations.  Rendering
        # both would duplicate drawings and text, so prefer Choice, then Fallback.
        if node.tag == MC_ALTERNATE_CONTENT:
            selected = next(
                (child for child in node if child.tag == MC_CHOICE),
                None,
            )
            if selected is None:
                selected = next(
                    (child for child in node if child.tag == MC_FALLBACK),
                    None,
                )
            return self._inline_text(selected) if selected is not None else ""

        return "".join(self._inline_text(child) for child in node)

    def paragraph_text(self, paragraph: etree._Element) -> str:
        standalone_math = [
            child
            for child in paragraph
            if child.tag in {M_OMATH, M_OMATH_PARA}
        ]
        meaningful_nonmath = [
            child
            for child in paragraph
            if child.tag
            not in {
                f"{{{NS['w']}}}pPr",
                f"{{{NS['w']}}}bookmarkStart",
                f"{{{NS['w']}}}bookmarkEnd",
                f"{{{NS['w']}}}proofErr",
                M_OMATH,
                M_OMATH_PARA,
            }
            and self._inline_text(child)
        ]
        if standalone_math and not meaningful_nonmath:
            formula = "".join(self._formula_text(node) for node in standalone_math)
            text = f"$${formula}$$"
        else:
            text = "".join(self._inline_text(child) for child in paragraph)

        page_break_nodes = xpath(paragraph, "./w:pPr/w:pageBreakBefore")
        page_break_before = any(
            node.get(f"{{{NS['w']}}}val", "true").lower()
            not in {"0", "false", "off", "no"}
            for node in page_break_nodes
        )
        return f"[分页] {text}" if page_break_before else text

    def _cell_text(self, cell: etree._Element) -> str:
        parts: list[str] = []
        for child in cell:
            if child.tag == W_P:
                parts.append(self.paragraph_text(child))
            elif child.tag == W_TBL:
                nested_table = Table(child, self.document)
                for row in nested_table.rows:
                    parts.append(" ".join(item.text for item in row.cells))
        return " ".join(parts).replace("\r", " ").replace("\n", " ")

    @staticmethod
    def _markdown_cell(text: str) -> str:
        return text.replace("|", r"\|")

    def table_lines(self, table_element: etree._Element) -> list[str]:
        table = Table(table_element, self.document)
        rows = [
            [self._markdown_cell(self._cell_text(cell._tc)) for cell in row.cells]
            for row in table.rows
        ]
        if not rows:
            return []
        width = max(len(row) for row in rows)
        rows = [row + [""] * (width - len(row)) for row in rows]

        def markdown_row(row: list[str]) -> str:
            return "| " + " | ".join(row) + " |"

        return [
            markdown_row(rows[0]),
            markdown_row(["---"] * width),
            *(markdown_row(row) for row in rows[1:]),
        ]

    def render(self) -> list[str]:
        lines = [
            "<!-- 自动生成的只读快照。事实源是 "
            f"../word/{self.path.name}，请勿手改本文件。"
            "由 scripts/export_snapshot.py 生成。 -->"
        ]
        for child in self.document.element.body.iterchildren():
            if child.tag == W_P:
                lines.append(self.paragraph_text(child))
            elif child.tag == W_TBL:
                lines.extend(self.table_lines(child))
        return lines


def export_snapshot(docx_path: Path) -> tuple[Path, int]:
    lines = SnapshotExporter(docx_path).render()
    snapshot_path = SNAPSHOT_DIR / f"{docx_path.stem}.md"
    content = ("\n".join(lines) + "\n").encode("utf-8")
    if not snapshot_path.is_file() or snapshot_path.read_bytes() != content:
        snapshot_path.write_bytes(content)
    return snapshot_path, len(lines)


def main(argv: list[str] | None = None) -> int:
    configure(project_dir_from_args("把五标准件 docx 导出为只读 md 快照", argv))
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    missing = [name for name in DOCX_NAMES if not (WORD_DIR / name).is_file()]
    if missing:
        print(f"缺少 DOCX：{missing}", file=sys.stderr)
        return 1

    for name in DOCX_NAMES:
        snapshot_path, line_count = export_snapshot(WORD_DIR / name)
        print(f"{snapshot_path}: {line_count} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
