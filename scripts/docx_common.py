#!/usr/bin/env python3
"""docx 底层读写的共用工具：read_revisions / write_tracked_change / diff_docx 共用。

只做三件事：解包 docx 的 XML 部件、提取文本（含公式）、安全地写回。
不依赖 python-docx，因为要直接操作 w:ins / w:del 这类 python-docx 不暴露的结构。
"""

from __future__ import annotations

import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}

W = NS["w"]
M = NS["m"]
W15 = NS["w15"]

DOCUMENT_PART = "word/document.xml"
COMMENTS_PART = "word/comments.xml"
PEOPLE_PART = "word/people.xml"


def q(tag: str) -> str:
    """把 'w:ins' 变成 '{namespace}ins'。"""
    prefix, _, local = tag.partition(":")
    return f"{{{NS[prefix]}}}{local}"


class DocxError(RuntimeError):
    pass


# ---------------------------------------------------------------- 锁文件


def check_not_open(docx: Path) -> None:
    """写入前必查：Word 开着时写入零报错，但用户一保存就把改动静默覆盖。

    事后用 lsof 查不出任何异常（os.replace 换掉了目录项，Word 抓的是旧 inode），
    所以这个检查只在写入前有意义。
    """
    lock = docx.parent / f"~${docx.name}"
    # Word 对长文件名的锁文件会截断前两个字符，两种形态都查
    alt = docx.parent / f"~${docx.name[2:]}"
    for candidate in (lock, alt):
        if candidate.exists():
            raise DocxError(
                f"检测到 Word 锁文件 {candidate.name}，说明该文件正开在 Word 里。\n"
                f"请先关闭 Word 再运行。此时强行写入会零报错地成功，"
                f"但用户保存一次就会把改动静默覆盖。"
            )


def backup(docx: Path, tag: str = "改前") -> Path:
    stamp = datetime.now().strftime("%H%M%S")
    target = docx.with_name(f"{docx.stem}_备份_{tag}_{stamp}{docx.suffix}")
    shutil.copy2(docx, target)
    return target


# ---------------------------------------------------------------- 读


def part_names(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as zf:
        return zf.namelist()


def read_part(docx: Path, name: str) -> bytes | None:
    with zipfile.ZipFile(docx) as zf:
        if name not in zf.namelist():
            return None
        return zf.read(name)


def parse_part(docx: Path, name: str):
    data = read_part(docx, name)
    if data is None:
        return None
    return etree.fromstring(data)


def document_root(docx: Path):
    root = parse_part(docx, DOCUMENT_PART)
    if root is None:
        raise DocxError(f"{docx} 里没有 {DOCUMENT_PART}")
    return root


# ---------------------------------------------------------------- 文本提取


def element_text(element, include_math: bool = True) -> str:
    """收集元素下的全部文字。

    必须同时收 m:t：批注正文里可能嵌 OMML 公式，只收 w:t 会静默丢字。
    真实教训：两条批注里的关键公式记号就是这样丢掉的，一度看不懂批注在问什么。
    """
    tags = [q("w:t"), q("w:delText")]
    if include_math:
        tags.append(q("m:t"))
    parts: list[str] = []
    for node in element.iter(*tags):
        if node.text:
            parts.append(node.text)
    return "".join(parts)


def body_paragraphs(root) -> list:
    """正文顶层段落。不含表格内段落，与 python-docx 的 paragraphs 口径一致。"""
    return root.findall(f".//{q('w:body')}/{q('w:p')}")


def all_paragraphs(root) -> list:
    return root.findall(f".//{q('w:p')}")


def paragraph_text(paragraph, include_math: bool = True) -> str:
    return element_text(paragraph, include_math=include_math)


# ---------------------------------------------------------------- 修订 id


def max_revision_id(root) -> int:
    """现有 w:id 的最大值。新写的修订要从它 +1 起，避免撞号。"""
    ids = []
    for node in root.iter():
        value = node.get(q("w:id"))
        if value is None:
            continue
        try:
            ids.append(int(value))
        except ValueError:
            continue
    return max(ids) if ids else 0


# ---------------------------------------------------------------- 写


def write_parts(src: Path, dst: Path, replacements: dict[str, bytes]) -> None:
    """把 src 复制成 dst，其中 replacements 指定的部件换成新字节，其余原样保留。

    逐条目复制而不是整包重压，避免动到没必要动的部件。
    """
    with zipfile.ZipFile(src) as zin:
        infos = zin.infolist()
        payload = {info.filename: zin.read(info.filename) for info in infos}
    for name, data in replacements.items():
        if name not in payload:
            raise DocxError(f"{src} 里没有部件 {name}，不能替换")
        payload[name] = data

    tmp = dst.with_suffix(dst.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, payload[info.filename])
    tmp.replace(dst)


def serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
