#!/usr/bin/env python3
"""文风校订的不变量校验：比较改写前后两份 md 草案，确认只动了表达、没动内容。

用法：
    python check_style_invariants.py 草案.改前.md 草案.md --terms 术语表.txt

术语表每行一个词，# 开头的行是注释。术语确实需要增减时，用 --allow-term 逐个放行。
用户决定修改权利要求后，用 --allow-claims 放行第 3 项，差异仍会打印出来。

这个脚本只能证明数字、公式、术语、固定句式没变，查不出句子含义有没有变，
所以它替代不了人工审核（见 references/文风校订.md）。

退出码：0 全部通过；1 有校验项不过；2 输入有问题（文件缺失、章节切不出来）。
"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

ABSTRACT = "二、摘要"
CLAIMS = "三、权利要求书"
DESCRIPTION = "四、说明书"
FIELD = "技术领域"
SUMMARY = "发明内容"
FIGURES = "附图说明"
EMBODIMENT = "具体实施方式"

MAX_ABSTRACT_CHARS = 300
MAX_LISTED_DIFFS = 8

DISPLAY_FORMULA = re.compile(r"\$\$(.+?)\$\$", re.DOTALL)
INLINE_FORMULA = re.compile(r"\$([^$\n]+)\$")
TABLE = re.compile(r"<table>.*?</table>", re.DOTALL)
NUMBER = re.compile(r"\d+(?:\.\d+)?%?")
STEP_LABEL = re.compile(r"^\*{0,2}S(\d+)\*{0,2}[：:]", re.MULTILINE)
CLAIM_START = re.compile(r"^\*{0,2}(\d+)、\*{0,2}", re.MULTILINE)
FIGURE_NUMBER = re.compile(r"图\s*\d+")
WHITESPACE = re.compile(r"\s")
SEPARATOR = re.compile(r"\s*-{3,}\s*")

FROZEN_PATTERNS = {
    "数据流向句": re.compile(r"[^。\n]*提供给步骤[^。\n]*。"),
    "公式引出句": re.compile(r"[^。\n]*计算过程如下[：:]"),
    "符号说明段": re.compile(r"^其中[，,].*$", re.MULTILINE),
    "兜底声明": re.compile(r"^以上实施例仅用于说明.*$", re.MULTILINE),
}


class InputProblem(RuntimeError):
    """输入文件不可用。说清楚缺什么，由调用方报错退出。"""


@dataclass(frozen=True)
class Draft:
    name: str
    body: str
    sections: dict[str, str]
    subsections: dict[str, str]

    def units(self) -> dict[str, str]:
        """按二级标题切开，说明书再按三级标题切开。数字和行内公式按这个粒度比较。"""
        units = {name: text for name, text in self.sections.items() if name != DESCRIPTION}
        units.update(self.subsections)
        return units


@dataclass(frozen=True)
class CheckResult:
    number: int
    title: str
    passed: bool
    details: tuple[str, ...] = ()


def split_by_heading(text: str, marker: str) -> dict[str, str]:
    pattern = re.compile(rf"^{re.escape(marker)}(?!#)\s*(.+?)\s*$", re.MULTILINE)
    matches = list(pattern.finditer(text))
    return {
        match.group(1): text[match.end() : matches[index + 1].start() if index + 1 < len(matches) else len(text)]
        for index, match in enumerate(matches)
    }


def load_draft(path: Path) -> Draft:
    """读 md 草案。> 开头的批注行和 --- 分隔线不进 Word，也不参与校验。"""
    if not path.is_file():
        raise InputProblem(f"找不到文件：{path}")
    lines = path.read_text(encoding="utf-8").splitlines()
    body = "\n".join(line for line in lines if not line.startswith(">") and not SEPARATOR.fullmatch(line))

    sections = split_by_heading(body, "## ")
    missing = [name for name in (ABSTRACT, CLAIMS, DESCRIPTION) if name not in sections]
    if missing:
        raise InputProblem(f"{path} 缺少二级标题：{missing}，标题要与 scripts/template.md 一字不差")

    subsections = split_by_heading(sections[DESCRIPTION], "### ")
    missing = [name for name in (FIELD, SUMMARY, FIGURES, EMBODIMENT) if name not in subsections]
    if missing:
        raise InputProblem(f"{path} 的说明书缺少三级标题：{missing}")

    return Draft(path.name, body, sections, subsections)


def load_terms(path: Path) -> list[str]:
    if not path.is_file():
        raise InputProblem(f"找不到术语表：{path}")
    lines = (line.strip() for line in path.read_text(encoding="utf-8").splitlines())
    terms = [line for line in lines if line and not line.startswith("#")]
    if not terms:
        raise InputProblem(f"术语表是空的：{path}")
    return terms


def strip_math(text: str) -> str:
    """去掉表格和公式。它们有专门的校验项，这样同一处改动不会重复报错。"""
    without_tables = TABLE.sub("", text)
    without_display = DISPLAY_FORMULA.sub("", without_tables)
    return INLINE_FORMULA.sub("", without_display)


def describe_counter_diff(before: Counter, after: Counter) -> str:
    only_before = sorted((before - after).elements())
    only_after = sorted((after - before).elements())
    return f"改前多 {only_before[:MAX_LISTED_DIFFS]}，改后多 {only_after[:MAX_LISTED_DIFFS]}"


def first_difference(before: str, after: str, context: int = 18) -> str:
    limit = min(len(before), len(after))
    index = next((i for i in range(limit) if before[i] != after[i]), limit)
    start = max(0, index - context)
    return f"改前「{before[start:index + context]}」，改后「{after[start:index + context]}」"


def compare_counters_by_unit(
    before: Draft, after: Draft, extract: callable, label: str
) -> list[str]:
    units_before, units_after = before.units(), after.units()
    problems = []
    for unit in sorted(set(units_before) | set(units_after)):
        counter_before = Counter(extract(units_before.get(unit, "")))
        counter_after = Counter(extract(units_after.get(unit, "")))
        if counter_before != counter_after:
            problems.append(f"{unit}的{label}：{describe_counter_diff(counter_before, counter_after)}")
    return problems


def check_numbers(before: Draft, after: Draft) -> CheckResult:
    problems = compare_counters_by_unit(
        before, after, lambda text: NUMBER.findall(strip_math(text)), "数字"
    )
    return CheckResult(1, "数字没有增、删、改", not problems, tuple(problems))


def check_formulas_and_tables(before: Draft, after: Draft) -> CheckResult:
    problems = []
    display_before = DISPLAY_FORMULA.findall(before.body)
    display_after = DISPLAY_FORMULA.findall(after.body)
    if len(display_before) != len(display_after):
        problems.append(f"独立公式数量：改前 {len(display_before)}，改后 {len(display_after)}")
    else:
        for index, (old, new) in enumerate(zip(display_before, display_after), start=1):
            if old != new:
                problems.append(f"第 {index} 个独立公式：{first_difference(old, new)}")

    def inline_formulas(text: str) -> list[str]:
        return INLINE_FORMULA.findall(DISPLAY_FORMULA.sub("", TABLE.sub("", text)))

    problems.extend(compare_counters_by_unit(before, after, inline_formulas, "行内公式"))

    if TABLE.findall(before.body) != TABLE.findall(after.body):
        problems.append("表格内容有变化")
    return CheckResult(2, "公式与表格逐字节相等", not problems, tuple(problems))


def check_claims(before: Draft, after: Draft, allow: bool) -> CheckResult:
    old, new = before.sections[CLAIMS], after.sections[CLAIMS]
    if hashlib.sha256(old.encode()).digest() == hashlib.sha256(new.encode()).digest():
        return CheckResult(3, "权利要求书整节不变", True)
    difference = first_difference(old, new)
    if allow:
        return CheckResult(3, "权利要求书整节不变", True, (f"已按 --allow-claims 放行：{difference}",))
    return CheckResult(3, "权利要求书整节不变", False, (difference,))


def check_terms(before: Draft, after: Draft, terms: Iterable[str], allowed: set[str]) -> CheckResult:
    problems, notes = [], []
    for term in terms:
        count_before, count_after = before.body.count(term), after.body.count(term)
        if count_before == 0:
            notes.append(f"提示：术语「{term}」在改前稿里没有出现，确认术语表写对了")
        if count_before == count_after:
            continue
        message = f"「{term}」：改前 {count_before} 次，改后 {count_after} 次"
        if term in allowed:
            notes.append(f"已放行 {message}")
        else:
            problems.append(message)
    return CheckResult(4, "术语出现次数不变", not problems, tuple(problems + notes))


def check_steps(before: Draft, after: Draft) -> CheckResult:
    problems = []
    for unit in (SUMMARY, EMBODIMENT):
        old = STEP_LABEL.findall(before.subsections[unit])
        new = STEP_LABEL.findall(after.subsections[unit])
        if old != new:
            problems.append(f"{unit}的步骤编号：改前 {old}，改后 {new}")

    old_region = scheme_region(before.subsections[SUMMARY])
    new_region = scheme_region(after.subsections[SUMMARY])
    if old_region != new_region:
        problems.append(f"技术方案步骤段不允许改动，标点也不改：{first_difference(old_region, new_region)}")
    return CheckResult(5, "S1–Sn 编号与顺序不变，技术方案步骤段不变", not problems, tuple(problems))


def scheme_region(summary: str) -> str:
    """取发明内容里从第一个步骤到最后一个步骤段末的文字，只去掉空白，标点参与比较。"""
    labels = list(STEP_LABEL.finditer(summary))
    if not labels:
        return ""
    paragraph_end = summary.find("\n\n", labels[-1].start())
    end = paragraph_end if paragraph_end != -1 else len(summary)
    return WHITESPACE.sub("", summary[labels[0].start() : end])


def check_frozen(before: Draft, after: Draft) -> CheckResult:
    problems = []
    old_description, new_description = before.sections[DESCRIPTION], after.sections[DESCRIPTION]
    for name, pattern in FROZEN_PATTERNS.items():
        old = Counter(pattern.findall(old_description))
        new = Counter(pattern.findall(new_description))
        if old != new:
            problems.append(f"{name}：{describe_counter_diff(old, new)}")

    if before.subsections[FIELD].strip() != after.subsections[FIELD].strip():
        problems.append(f"技术领域属冻结区：{first_difference(before.subsections[FIELD], after.subsections[FIELD])}")

    old_figures = FIGURE_NUMBER.findall(before.subsections[FIGURES])
    new_figures = FIGURE_NUMBER.findall(after.subsections[FIGURES])
    if old_figures != new_figures:
        problems.append(f"附图说明的图号：改前 {old_figures}，改后 {new_figures}")
    return CheckResult(6, "冻结句式、技术领域、图号不变", not problems, tuple(problems))


def check_hard_limits(after: Draft) -> CheckResult:
    problems = []
    abstract_lines = [line for line in after.sections[ABSTRACT].splitlines() if not line.startswith("摘要附图")]
    abstract_chars = len(re.sub(r"\s", "", "".join(abstract_lines)))
    if abstract_chars > MAX_ABSTRACT_CHARS:
        problems.append(f"摘要 {abstract_chars} 字，超过 {MAX_ABSTRACT_CHARS} 字")

    claims = after.sections[CLAIMS]
    starts = list(CLAIM_START.finditer(claims))
    if not starts:
        problems.append("切不出权利要求编号")
    for index, match in enumerate(starts):
        end = starts[index + 1].start() if index + 1 < len(starts) else len(claims)
        text = strip_math(claims[match.end() : end]).strip()
        if text.count("。") != 1 or not text.endswith("。"):
            problems.append(f"权利要求 {match.group(1)} 应当只在结尾有一个句号，实际 {text.count('。')} 个")
    return CheckResult(7, "摘要不超过 300 字，权项每条一个句号", not problems, tuple(problems))


def run_checks(before: Draft, after: Draft, terms: list[str], allowed: set[str], allow_claims: bool) -> list[CheckResult]:
    return [
        check_numbers(before, after),
        check_formulas_and_tables(before, after),
        check_claims(before, after, allow_claims),
        check_terms(before, after, terms, allowed),
        check_steps(before, after),
        check_frozen(before, after),
        check_hard_limits(after),
    ]


def print_report(results: list[CheckResult]) -> None:
    for result in results:
        print(f"[{'通过' if result.passed else '不过'}] {result.number}. {result.title}")
        for detail in result.details:
            print(f"    {detail}")
    failed = sum(not result.passed for result in results)
    print(f"结果：{len(results) - failed}/{len(results)} 通过")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="校验文风校订的改写没有动内容")
    parser.add_argument("before", type=Path, help="改写前的 md 草案")
    parser.add_argument("after", type=Path, help="改写后的 md 草案")
    parser.add_argument("--terms", type=Path, required=True, help="术语表，每行一个词")
    parser.add_argument("--allow-term", action="append", default=[], help="放行某个术语的次数变化，可重复")
    parser.add_argument("--allow-claims", action="store_true", help="用户已决定修改权利要求时放行第 3 项")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        before, after = load_draft(args.before), load_draft(args.after)
        terms = load_terms(args.terms)
    except InputProblem as problem:
        print(f"输入有问题：{problem}", file=sys.stderr)
        return 2
    results = run_checks(before, after, terms, set(args.allow_term), args.allow_claims)
    print_report(results)
    return 0 if all(result.passed for result in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
