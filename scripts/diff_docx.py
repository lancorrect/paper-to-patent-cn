#!/usr/bin/env python3
"""比对两份 docx：谁在这一轮新增了哪些修订，段落文本有何变化。

典型用途：代理机构返回新一轮稿子时，先跟上一轮比一比，看这轮到底改了什么。

    python diff_docx.py 上一轮.docx 这一轮.docx
    python diff_docx.py 旧.docx 新.docx --author 专利组1   # 只看某个作者
    python diff_docx.py 旧.docx 新.docx --paragraphs        # 连段落文本差异一起看

**按「作者 + 文本」聚合做差集，不按 w:id 比。** Word 保存时会重编 w:id，实测一天之内
重编过两次，其中一条批注的新 id 恰好撞上另一条的旧号；段落索引也随内容增删漂移。
按内容比是唯一稳的办法。
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from docx_common import (  # noqa: E402
    COMMENTS_PART,
    DocxError,
    all_paragraphs,
    document_root,
    element_text,
    paragraph_text,
    parse_part,
    q,
)


def revision_bag(docx: Path, kind: str, author: str | None) -> Counter:
    """(作者, 文本) -> 次数。空白文本的段落标记修订不计入。"""
    root = document_root(docx)
    bag: Counter = Counter()
    for node in root.iter(q(f"w:{kind}")):
        who = node.get(q("w:author"))
        if author and who != author:
            continue
        text = element_text(node)
        if not text.strip():
            continue
        bag[(who, text)] += 1
    return bag


def comment_count(docx: Path) -> int:
    root = parse_part(docx, COMMENTS_PART)
    return 0 if root is None else len(list(root.iter(q("w:comment"))))


def author_summary(docx: Path) -> dict:
    root = document_root(docx)
    summary: dict = {}
    for kind in ("ins", "del"):
        for node in root.iter(q(f"w:{kind}")):
            if not element_text(node).strip():
                continue
            summary[(kind, node.get(q("w:author")))] = (
                summary.get((kind, node.get(q("w:author"))), 0) + 1
            )
    return summary


def paragraph_diff(old: Path, new: Path) -> list[str]:
    """段落级差异。Word 保存会归一化 XML，所以只能比文本，不能比字节。"""
    old_texts = [paragraph_text(p) for p in all_paragraphs(document_root(old))]
    new_texts = [paragraph_text(p) for p in all_paragraphs(document_root(new))]
    lines: list[str] = []
    if len(old_texts) != len(new_texts):
        lines.append(f"段落数：{len(old_texts)} -> {len(new_texts)}")
    for index, (a, b) in enumerate(zip(old_texts, new_texts)):
        if a != b:
            lines.append(f"  第 {index} 段：\n    - {a[:90]}\n    + {b[:90]}")
    return lines


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="比对两份 docx 的修订与文本")
    parser.add_argument("old", type=Path)
    parser.add_argument("new", type=Path)
    parser.add_argument("--author", help="只看这个作者的修订")
    parser.add_argument("--paragraphs", action="store_true", help="连段落文本差异一起列")
    args = parser.parse_args(argv)

    for path in (args.old, args.new):
        if not path.is_file():
            print(f"文件不存在：{path}", file=sys.stderr)
            return 2

    try:
        print(f"旧：{args.old.name}")
        print(f"新：{args.new.name}")
        print(f"批注数：{comment_count(args.old)} -> {comment_count(args.new)}")

        old_summary, new_summary = author_summary(args.old), author_summary(args.new)
        keys = sorted(set(old_summary) | set(new_summary))
        print("\n带文字的修订按作者：")
        for kind, who in keys:
            before = old_summary.get((kind, who), 0)
            after = new_summary.get((kind, who), 0)
            mark = "" if before == after else f"   ({after - before:+d})"
            print(f"  {kind} · {who}：{before} -> {after}{mark}")

        for kind in ("ins", "del"):
            old_bag = revision_bag(args.old, kind, args.author)
            new_bag = revision_bag(args.new, kind, args.author)
            added = new_bag - old_bag
            removed = old_bag - new_bag
            print(f"\n=== 新增的 w:{kind}（{sum(added.values())} 处）===")
            for (who, text), count in added.most_common():
                suffix = f" x{count}" if count > 1 else ""
                print(f"  [{who}]{suffix} {text[:100]}")
            if removed:
                print(f"--- 不再出现的 w:{kind}（{sum(removed.values())} 处）---")
                for (who, text), count in removed.most_common():
                    suffix = f" x{count}" if count > 1 else ""
                    print(f"  [{who}]{suffix} {text[:100]}")

        if args.paragraphs:
            print("\n=== 段落文本差异 ===")
            lines = paragraph_diff(args.old, args.new)
            print("\n".join(lines) if lines else "  无")
        return 0
    except DocxError as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
