#!/usr/bin/env python3
"""在带批注的 docx 里以 Word 修订模式写入改动。

代理机构要求在原先带批注和审阅的文件里改、保留全部痕迹，所以不能接受全部修订后再改，
也不能导出 md 改完重新打包。只能直接写 w:ins / w:del。

    # 先看会命中什么，不写文件
    python write_tracked_change.py 稿子.docx --anchor "原文片段" --replace "新文字" --dry-run

    # 真写（自动备份 + 双向验证）
    python write_tracked_change.py 稿子.docx --anchor "原文片段" --replace "新文字"
    python write_tracked_change.py 稿子.docx --anchor "某段结尾" --append "补充一句。"
    python write_tracked_change.py 稿子.docx --anchor "要删的片段" --delete

四条安全约束，都是真实事故换来的，不要绕过：

1. **锚点命中次数必须恰好为 1**，否则中止。Word 保存会重排 run 边界，锚点会漂移；
   曾经靠这条断言挡住一次误改。正确处置是重新定位锚点，不是放宽匹配。
2. **写入前必查 Word 锁文件**。开着时写入零报错、验收全过，用户一保存就静默覆盖。
3. **w:author 必须是 Word 当前登录账户名**，否则 Word 把改动显示成「别人给的建议」，
   点击会弹接受/拒绝卡片，与手打效果不同。默认从 people.xml 反解。
4. **改完双向验证**：接受修订后结果正确，拒绝修订后逐字回到原样。
"""

from __future__ import annotations

import argparse
import sys
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from docx_common import (  # noqa: E402
    DOCUMENT_PART,
    PEOPLE_PART,
    DocxError,
    all_paragraphs,
    backup,
    check_not_open,
    document_root,
    max_revision_id,
    paragraph_text,
    parse_part,
    q,
    serialize,
    write_parts,
)


# ---------------------------------------------------------------- 作者身份


def resolve_author(docx: Path, explicit: str | None) -> str:
    people = parse_part(docx, PEOPLE_PART)
    names = []
    if people is not None:
        names = [p.get(q("w15:author")) for p in people.iter(q("w15:person"))]

    if explicit:
        if names and explicit not in names:
            print(
                f"[警告] «{explicit}» 不在 people.xml（现有：{names}）。\n"
                f"        若它不是 Word 当前登录账户名，改动会被显示成「别人给的建议」，"
                f"而不是自己的修订。",
                file=sys.stderr,
            )
        return explicit

    if len(names) == 1:
        print(f"[作者] 从 people.xml 取到唯一作者：{names[0]}")
        return names[0]
    raise DocxError(
        f"people.xml 里有 {len(names)} 个作者 {names}，无法自动判断。\n"
        f"请用 --author 指定 Word 当前登录账户名（不是中文姓名）。\n"
        f"先跑 read_revisions.py 看这份文件里登记了谁。"
    )


def ensure_person(docx: Path, author: str, user_id: str | None) -> bytes | None:
    """把作者补登记进 people.xml。返回新字节，无需改动则返回 None。"""
    root = parse_part(docx, PEOPLE_PART)
    if root is None:
        print(
            "[警告] 该文件没有 word/people.xml，无法登记作者。"
            "Word 可能把修订显示为外来建议。",
            file=sys.stderr,
        )
        return None
    existing = [p.get(q("w15:author")) for p in root.iter(q("w15:person"))]
    if author in existing:
        return None
    person = etree.SubElement(root, q("w15:person"))
    person.set(q("w15:author"), author)
    info = etree.SubElement(person, q("w15:presenceInfo"))
    info.set(q("w15:providerId"), "AD")
    info.set(q("w15:userId"), user_id or author)
    print(f"[people.xml] 已登记作者 {author}")
    return serialize(root)


# ---------------------------------------------------------------- 构造 run


def make_text_run(source_run, text: str, as_deleted: bool):
    """照抄原 run 的格式，换掉文字。删除态用 w:delText，插入态用 w:t。"""
    run = deepcopy(source_run)
    for child in list(run):
        if child.tag in (q("w:t"), q("w:delText")):
            run.remove(child)
    node = etree.SubElement(run, q("w:delText") if as_deleted else q("w:t"))
    node.text = text
    if text != text.strip():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def wrap(kind: str, run, rev_id: int, author: str, stamp: str):
    """用 w:ins 或 w:del 包住一个 run。"""
    holder = etree.Element(q(f"w:{kind}"))
    holder.set(q("w:id"), str(rev_id))
    holder.set(q("w:author"), author)
    holder.set(q("w:date"), stamp)
    holder.append(run)
    return holder


# ---------------------------------------------------------------- 定位


def locate(root, anchor: str):
    """找到含 anchor 的那个 w:t。命中必须恰好 1 次。"""
    hits = [node for node in root.iter(q("w:t")) if node.text and anchor in node.text]
    if len(hits) == 1:
        return hits[0]

    if not hits:
        # 是不是跨 run 了？给出可操作的提示，而不是放宽匹配
        for index, paragraph in enumerate(all_paragraphs(root)):
            if anchor in paragraph_text(paragraph):
                raise DocxError(
                    f"锚点命中 0 次，但它出现在第 {index} 段的段落文本里，"
                    f"说明这段文字被拆到了多个 run 中（Word 保存会重排 run 边界）。\n"
                    f"换一个更短、落在单个 run 内的锚点重试。"
                )
        raise DocxError(f"锚点命中 0 次，文件里没有这段文字：{anchor!r}")

    raise DocxError(
        f"锚点命中 {len(hits)} 次，不唯一。请加长锚点使其唯一。\n"
        f"绝不允许改成模糊匹配硬凑——那样会改错地方。"
    )


# ---------------------------------------------------------------- 修订模拟


def _accept(root) -> None:
    """模拟「接受全部修订」：删掉 w:del 整体，把 w:ins 解包。"""
    for node in list(root.iter(q("w:del"))):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    for node in list(root.iter(q("w:ins"))):
        parent = node.getparent()
        if parent is None:
            continue
        index = list(parent).index(node)
        for offset, child in enumerate(list(node)):
            parent.insert(index + offset, child)
        parent.remove(node)


def _reject(root) -> None:
    """模拟「拒绝全部修订」：删掉 w:ins 整体，把 w:del 解包且 delText 变回 t。"""
    for node in list(root.iter(q("w:ins"))):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    for node in list(root.iter(q("w:del"))):
        parent = node.getparent()
        if parent is None:
            continue
        for deleted in node.iter(q("w:delText")):
            deleted.tag = q("w:t")
        index = list(parent).index(node)
        for offset, child in enumerate(list(node)):
            parent.insert(index + offset, child)
        parent.remove(node)


def rendered_text(root, mode: str) -> list[str]:
    clone = deepcopy(root)
    if mode == "accept":
        _accept(clone)
    else:
        _reject(clone)
    return [paragraph_text(p) for p in all_paragraphs(clone)]


# ---------------------------------------------------------------- 改动


def apply_change(root, anchor: str, mode: str, payload: str, author: str) -> int:
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    next_id = max_revision_id(root) + 1

    target = locate(root, anchor)
    run = target.getparent()
    paragraph = run.getparent()
    if paragraph is None or paragraph.tag != q("w:p"):
        raise DocxError("锚点所在 run 不在普通段落里，可能位于表格或公式中，请人工处理")

    if mode == "append":
        new_run = make_text_run(run, payload, as_deleted=False)
        paragraph.append(wrap("ins", new_run, next_id, author, stamp))
        return 1

    text = target.text
    before, _, after = text.partition(anchor)
    index = list(paragraph).index(run)
    pieces = []
    if before:
        pieces.append(make_text_run(run, before, as_deleted=False))
    pieces.append(
        wrap("del", make_text_run(run, anchor, as_deleted=True), next_id, author, stamp)
    )
    used = 1
    if mode == "replace":
        pieces.append(
            wrap(
                "ins",
                make_text_run(run, payload, as_deleted=False),
                next_id + 1,
                author,
                stamp,
            )
        )
        used = 2
    if after:
        pieces.append(make_text_run(run, after, as_deleted=False))

    paragraph.remove(run)
    for offset, piece in enumerate(pieces):
        paragraph.insert(index + offset, piece)
    return used


# ---------------------------------------------------------------- 验证


def verify(before_root, after_root, anchor: str, mode: str, payload: str) -> list[str]:
    problems: list[str] = []

    before_reject = rendered_text(before_root, "reject")
    after_reject = rendered_text(after_root, "reject")
    if before_reject != after_reject:
        diff = [
            f"    第 {i} 段：{b!r} -> {a!r}"
            for i, (b, a) in enumerate(zip(before_reject, after_reject))
            if b != a
        ]
        problems.append("拒绝修订后与改前不一致：\n" + "\n".join(diff[:5]))

    before_accept = rendered_text(before_root, "accept")
    after_accept = rendered_text(after_root, "accept")
    changed = [i for i, (b, a) in enumerate(zip(before_accept, after_accept)) if b != a]
    if len(changed) != 1:
        problems.append(f"接受修订后应恰好 1 段变化，实际 {len(changed)} 段：{changed[:5]}")
    else:
        index = changed[0]
        got, old = after_accept[index], before_accept[index]
        expect = {
            "replace": old.replace(anchor, payload, 1),
            "delete": old.replace(anchor, "", 1),
            "append": old + payload,
        }[mode]
        if got != expect:
            problems.append(f"接受修订后第 {index} 段内容不符：\n    期望 {expect!r}\n    实际 {got!r}")
        else:
            print(f"[验证] 接受修订后第 {index} 段：{got.strip()[:70]}")
    return problems


# ---------------------------------------------------------------- 主流程


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="以 Word 修订模式在 docx 里写入改动")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--anchor", required=True, help="要定位的原文片段，必须唯一命中")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--replace", metavar="新文字", help="把锚点换成这段文字")
    group.add_argument("--append", metavar="新文字", help="在锚点所在段末尾追加")
    group.add_argument("--delete", action="store_true", help="删除锚点这段文字")
    parser.add_argument("--author", help="Word 当前登录账户名，默认从 people.xml 反解")
    parser.add_argument("--user-id", help="补登记 people.xml 时用的 userId")
    parser.add_argument("--dry-run", action="store_true", help="只报告命中位置，不写文件")
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args(argv)

    if not args.docx.is_file():
        print(f"文件不存在：{args.docx}", file=sys.stderr)
        return 2

    mode = "replace" if args.replace is not None else ("append" if args.append else "delete")
    payload = args.replace or args.append or ""

    try:
        root = document_root(args.docx)
        target = locate(root, args.anchor)
        paragraph = target.getparent().getparent()
        index = all_paragraphs(root).index(paragraph)
        print(f"[命中] 唯一命中，第 {index} 段：{paragraph_text(paragraph).strip()[:70]}")
        if args.dry_run:
            print("[dry-run] 未写入任何文件")
            return 0

        # 写入前必查，事后查不出来
        check_not_open(args.docx)

        author = resolve_author(args.docx, args.author)
        before_root = deepcopy(root)
        count = apply_change(root, args.anchor, mode, payload, author)
        print(f"[改动] {mode}，新增 {count} 处修订，作者 {author}")

        problems = verify(before_root, root, args.anchor, mode, payload)
        if problems:
            print("\n[验证失败] 未写入文件：", file=sys.stderr)
            for problem in problems:
                print(f"  - {problem}", file=sys.stderr)
            return 1

        replacements = {DOCUMENT_PART: serialize(root)}
        people_bytes = ensure_person(args.docx, author, args.user_id)
        if people_bytes is not None:
            replacements[PEOPLE_PART] = people_bytes

        if not args.no_backup:
            print(f"[备份] {backup(args.docx).name}")
        write_parts(args.docx, args.docx, replacements)
        print("[完成] 已写入。请在 Word 里目视确认渲染正常，公式类改动尤其要看。")
        return 0

    except DocxError as exc:
        print(f"\n[中止] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
