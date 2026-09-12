#!/usr/bin/env python3
"""专利项目配置：build_docx.py / verify_docx.py / export_snapshot.py 共用。

三个脚本原本把发明名称、附图清单、目录路径写死在源码里（绑死第一篇专利）。
改为从专利版本目录下的 patent.json 读取，脚本本身不含任何具体专利的内容。

patent.json 放在专利的版本目录里，例如 <专利>/v5/patent.json：

    {
      "title": "一种……的方法",
      "source_md": "v5草案.md",
      "template_dir": "../../模板",
      "figures": [
        {"file": "model_architecture.png", "caption": "图1 整体架构示意图"},
        {"file": "module_detail.png",      "caption": "图2 图1 中××模块的结构示意图"},
        {"file": "training.png",           "caption": "图3 训练过程示意图"},
        {"file": "flowchart.png",          "caption": "图4 实现步骤流程图"}
      ],
      "abstract_figure": "flowchart.png"
    }

字段说明：
  title            发明名称，必须与 md 草案「## 一、发明名称」下的文字一致，不超过 25 字。
  source_md        md 草案文件名，相对版本目录。默认 "草案.md"。
  template_dir     CNIPA 官方空白模板目录，相对版本目录或绝对路径。默认 "../../模板"。
  figures          附图清单，顺序即图号顺序。file 相对 figures/ 目录，caption 是图号加图题。
  abstract_figure  摘要附图用哪张，填 figures 里的某个 file。
                   CNIPA 要求摘要附图必须是说明书附图之一，所以它必须出现在 figures 里。

目录约定（都相对版本目录，不可配置）：
  figures/     附图 PNG
  word/        输出的五个 docx
  snapshots/   只读 md 快照
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path

CONFIG_NAME = "patent.json"

# CNIPA 标准件号，固定不变
STANDARD_FILES = (
    "100001权利要求书.docx",
    "100002说明书.docx",
    "100003说明书附图.docx",
    "100004说明书摘要.docx",
    "摘要附图.docx",
)

DESCRIPTION_HEADINGS = ("技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式")

MAX_TITLE_CHARS = 25


class ConfigError(RuntimeError):
    """配置缺失或不合法。报错要说清楚缺什么、去哪补。"""


@dataclass(frozen=True)
class PatentConfig:
    project_dir: Path
    title: str
    source_md: Path
    template_dir: Path
    figure_dir: Path
    word_dir: Path
    snapshot_dir: Path
    figure_specs: tuple[tuple[Path, str], ...]
    abstract_figure: Path

    @property
    def figures(self) -> tuple[Path, ...]:
        return tuple(path for path, _ in self.figure_specs)

    def template(self, name: str) -> Path:
        return self.template_dir / name

    def output(self, name: str) -> Path:
        return self.word_dir / name


def load(project_dir: Path | str) -> PatentConfig:
    project_dir = Path(project_dir).expanduser().resolve()
    if not project_dir.is_dir():
        raise ConfigError(f"版本目录不存在：{project_dir}")

    config_path = project_dir / CONFIG_NAME
    if not config_path.is_file():
        raise ConfigError(
            f"缺少配置文件 {config_path}\n"
            f"从 skill 的 scripts/patent.json.example 复制一份过去，填上发明名称与附图清单。"
        )

    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ConfigError(f"{config_path} 不是合法的 JSON：{exc}") from exc

    title = str(raw.get("title", "")).strip()
    if not title:
        raise ConfigError(f"{config_path} 缺少 title（发明名称）")
    if len(title) > MAX_TITLE_CHARS:
        # 代理机构明确要求压到 25 字以内，这里只提醒不拦截
        print(
            f"[提醒] 发明名称 {len(title)} 字，超过建议上限 {MAX_TITLE_CHARS} 字：{title}",
            flush=True,
        )

    source_md = project_dir / str(raw.get("source_md", "草案.md"))
    if not source_md.is_file():
        raise ConfigError(f"md 草案不存在：{source_md}（在 {CONFIG_NAME} 的 source_md 里改）")

    template_dir = Path(str(raw.get("template_dir", "../../模板")))
    if not template_dir.is_absolute():
        template_dir = (project_dir / template_dir).resolve()
    if not template_dir.is_dir():
        raise ConfigError(f"官方模板目录不存在：{template_dir}")

    figure_dir = project_dir / "figures"
    raw_figures = raw.get("figures") or []
    if not raw_figures:
        raise ConfigError(f"{config_path} 的 figures 为空，至少要有一张附图")

    specs: list[tuple[Path, str]] = []
    for index, item in enumerate(raw_figures, start=1):
        if not isinstance(item, dict) or "file" not in item or "caption" not in item:
            raise ConfigError(f"{config_path} 的 figures 第 {index} 项要有 file 与 caption")
        path = figure_dir / str(item["file"])
        if not path.is_file():
            raise ConfigError(f"附图文件不存在：{path}")
        specs.append((path, str(item["caption"])))

    abstract_name = str(raw.get("abstract_figure", "")).strip()
    if not abstract_name:
        raise ConfigError(
            f"{config_path} 缺少 abstract_figure。"
            f"CNIPA 要求摘要附图必须是说明书附图之一，填 figures 里的某个 file。"
        )
    abstract_figure = figure_dir / abstract_name
    if abstract_figure not in [path for path, _ in specs]:
        raise ConfigError(
            f"abstract_figure «{abstract_name}» 不在 figures 清单里。"
            f"摘要附图必须同时编入说明书附图，否则不合规。"
        )

    return PatentConfig(
        project_dir=project_dir,
        title=title,
        source_md=source_md,
        template_dir=template_dir,
        figure_dir=figure_dir,
        word_dir=project_dir / "word",
        snapshot_dir=project_dir / "snapshots",
        figure_specs=tuple(specs),
        abstract_figure=abstract_figure,
    )


def add_project_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--project",
        required=True,
        metavar="DIR",
        help=f"专利版本目录（其中含 {CONFIG_NAME}、figures/、word/）",
    )


def project_dir_from_args(description: str, argv: list[str] | None = None) -> Path:
    """只需要目录、不需要完整配置的脚本用这个（verify / export）。"""
    parser = argparse.ArgumentParser(description=description)
    add_project_arg(parser)
    args = parser.parse_args(argv)
    project_dir = Path(args.project).expanduser().resolve()
    if not project_dir.is_dir():
        raise ConfigError(f"版本目录不存在：{project_dir}")
    return project_dir


def title_from_project(project_dir: Path) -> str:
    """verify 需要发明名称但不需要 md 与附图，单独取一次，缺 title 才报错。"""
    config_path = project_dir / CONFIG_NAME
    if not config_path.is_file():
        raise ConfigError(f"缺少配置文件 {config_path}")
    raw = json.loads(config_path.read_text(encoding="utf-8"))
    title = str(raw.get("title", "")).strip()
    if not title:
        raise ConfigError(f"{config_path} 缺少 title（发明名称）")
    return title


def captions_from_project(project_dir: Path) -> list[str]:
    """verify 只要图题，不需要图文件存在，单独取一次。"""
    config_path = project_dir / CONFIG_NAME
    if not config_path.is_file():
        raise ConfigError(f"缺少配置文件 {config_path}")
    raw = json.loads(config_path.read_text(encoding="utf-8"))
    return [str(item["caption"]) for item in (raw.get("figures") or [])]
