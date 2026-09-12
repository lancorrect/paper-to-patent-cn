#!/usr/bin/env python3
"""读出 docx 里的批注与修订痕迹。

代理机构返回的稿子带 w:ins / w:del 修订和 comments.xml 批注，用这个脚本先摸清
这一轮到底提了什么，再决定要不要逐段细读。

    python read_revisions.py <docx>              # 摘要
    python read_revisions.py <docx> --full       # 连修订文本一起列
    python read_revisions.py <docx> --json       # 机读

三条要点写在代码里，别绕过：
1. 批注正文可能嵌 OMML 公式，只收 w:t 会静默丢字，必须同时收 m:t；
2. 批注 id 与段落索引都会被 Word 重编（实测一天内重编两次），**引用批注按内容不按编号**；
3. people.xml 里的作者名是写修订时必须用的身份，这里一并反解出来。
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from docx_common import (  # noqa: E402
    COMMENTS_PART,
    PEOPLE_PART,
    DocxError,
    all_paragraphs,
    document_root,
    element_text,
    max_revision_id,
    paragraph_text,
    parse_part,
    q,
)


def read_people(docx: Path) -> list[dict]:
    """反解 Word 登录账户名。写修订时 w:author 必须用这里的名字。"""
    root = parse_part(docx, PEOPLE_PART)
    if root is None:
        return []
    people = []
    for person in root.iter(q("w15:person")):
        info = person.find(q("w15:presenceInfo"))
        people.append(
            {
                "author": person.get(q("w15:author")),
                "provider": info.get(q("w15:providerId")) if info is not None else None,
                "user_id": info.get(q("w15:userId")) if info is not None else None,
            }
        )
    return people


def read_comments(docx: Path) -> list[dict]:
    root = parse_part(docx, COMMENTS_PART)
    if root is None:
        return []
    comments = []
    for node in root.iter(q("w:comment")):
        comments.append(
            {
                "id": node.get(q("w:id")),
                "author": node.get(q("w:author")),
                "date": node.get(q("w:date")),
                # 同时收 w:t 与 m:t，否则批注里的公式会静默丢失
                "text": element_text(node).strip(),
            }
        )
    return comments


def anchor_map(doc_root) -> dict[str, dict]:
    """把批注 id 关联到它锚在哪一段、锚定的文字是什么。"""
    anchors: dict[str, dict] = {}
    paragraphs = all_paragraphs(doc_root)
    index_of = {id(p): i for i, p in enumerate(paragraphs)}

    for node in doc_root.iter(q("w:commentReference")):
        cid = node.get(q("w:id"))
        paragraph = node
        while paragraph is not None and paragraph.tag != q("w:p"):
            paragraph = paragraph.getparent()
        if paragraph is None:
            continue
        anchors.setdefault(cid, {})["paragraph_index"] = index_of.get(id(paragraph))
        anchors[cid]["paragraph_text"] = paragraph_text(paragraph).strip()

    # 锚定区间的文字：commentRangeStart 到 commentRangeEnd 之间
    for start in doc_root.iter(q("w:commentRangeStart")):
        cid = start.get(q("w:id"))
        collected: list[str] = []
        node = start
        while True:
            node = node.getnext()
            if node is None:
                break
            if node.tag == q("w:commentRangeEnd") and node.get(q("w:id")) == cid:
                break
            collected.append(element_text(node))
        text = "".join(collected).strip()
        if text:
            anchors.setdefault(cid, {})["anchor_text"] = text
    return anchors


def read_revisions(doc_root) -> dict:
    result = {"ins": [], "del": []}
    for kind in ("ins", "del"):
        for node in doc_root.iter(q(f"w:{kind}")):
            text = element_text(node)
            if not text.strip():
                continue  # 段落标记的插入/删除，没有文字内容
            result[kind].append(
                {
                    "id": node.get(q("w:id")),
                    "author": node.get(q("w:author")),
                    "date": node.get(q("w:date")),
                    "text": text,
                }
            )
    return result


def collect(docx: Path) -> dict:
    doc_root = document_root(docx)
    comments = read_comments(docx)
    anchors = anchor_map(doc_root)
    for comment in comments:
        comment.update(anchors.get(comment["id"], {}))
    revisions = read_revisions(doc_root)
    return {
        "file": str(docx),
        "people": read_people(docx),
        "comments": comments,
        "revisions": revisions,
        "max_revision_id": max_revision_id(doc_root),
        "paragraph_count": len(all_paragraphs(doc_root)),
    }


def render(data: dict, full: bool) -> None:
    print(f"文件：{data['file']}")
    print(f"段落数：{data['paragraph_count']}    现有 w:id 最大值：{data['max_revision_id']}"
          f"（新修订从 {data['max_revision_id'] + 1} 起）")

    people = data["people"]
    if people:
        print("\npeople.xml 登记的作者（写修订时 w:author 必须用这里的名字）：")
        for person in people:
            print(f"  - {person['author']}    {person['user_id'] or ''}")
    else:
        print("\npeople.xml 无登记作者（该文件可能没有本机账户写过修订）")

    ins, dele = data["revisions"]["ins"], data["revisions"]["del"]
    print(f"\n修订：插入 {len(ins)} 处，删除 {len(dele)} 处")
    for kind, items in (("插入", ins), ("删除", dele)):
        by_author = Counter(item["author"] for item in items)
        for author, count in by_author.most_common():
            print(f"  {kind} · {author}：{count} 处")

    comments = data["comments"]
    print(f"\n批注：{len(comments)} 条")
    if comments:
        print("注意：批注 id 会被 Word 重编，段落索引也会漂移。**引用批注按内容，不按编号。**")
    for comment in comments:
        print(f"\n  [id={comment['id']} 作者={comment['author']} {comment.get('date', '')}]")
        if comment.get("anchor_text"):
            print(f"  锚定：{comment['anchor_text'][:60]}")
        elif comment.get("paragraph_text"):
            print(f"  所在段（第 {comment.get('paragraph_index')} 段）："
                  f"{comment['paragraph_text'][:60]}")
        print(f"  内容：{comment['text']}")

    if full:
        print("\n---- 修订全文 ----")
        for kind, items in (("INS", ins), ("DEL", dele)):
            for item in items:
                print(f"  [{kind} {item['author']}] {item['text']}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="读出 docx 的批注与修订痕迹")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--full", action="store_true", help="连每处修订的文本一起列出")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    args = parser.parse_args(argv)

    if not args.docx.is_file():
        print(f"文件不存在：{args.docx}", file=sys.stderr)
        return 2
    try:
        data = collect(args.docx)
    except DocxError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
    else:
        render(data, args.full)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
