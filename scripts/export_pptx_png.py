#!/usr/bin/env python3
"""把版本目录的 figures/附图合集.pptx 导出成 patent.json 里列的 PNG，页序对应 figures 列表的顺序。

用法：
    python export_pptx_png.py --project ~/Desktop/专利/某专利/v1

只支持 macOS，本机要装 PowerPoint。由 PowerPoint 先存成 PDF，再用 PyMuPDF 按 300 dpi 转 PNG
并裁掉四周空白。渲染用的是 PowerPoint 本身。Quick Look 不一样，它不画上标、斜纹底和部分叉号，
只能作预览。

PowerPoint for Mac 是沙盒应用，实测：
- 副本要放在 Office 组容器里，原文件不经 PowerPoint 打开；
- AppleScript 不报错，但文件不落在指定路径，而是落在沙盒目录的 Data.pdf。所以要去那里取。
  曾只看指定路径，误判成「脚本导不出」。

退出码：0 成功；1 导出失败或页数对不上；2 有锁文件或残留文件，需要先处理。
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import fitz
import numpy as np
from PIL import Image

WORK = Path.home() / "Library/Group Containers/UBF8T346G9.Office/patent_export"
LANDED = Path.home() / "Library/Containers/com.microsoft.Powerpoint/Data.pdf"

SOURCE_NAME = "附图合集.pptx"
DPI = 300
WHITE_THRESHOLD = 250
PADDING = 12
WAIT_SECONDS = 30

SAVE_AS_PDF = """
on run argv
  set srcPosix to item 1 of argv
  set dstPosix to item 2 of argv
  tell application "Microsoft PowerPoint"
    open (POSIX file srcPosix)
    delay 3
    set p to presentation "figs.pptx"
    save p in (POSIX file dstPosix as text) as save as PDF
    delay 3
    close p saving no
  end tell
end run
"""


def export_pdf(source: Path, target: Path) -> None:
    WORK.mkdir(exist_ok=True)
    try:
        copy, requested = WORK / "figs.pptx", WORK / "figs.pdf"
        shutil.copy2(source, copy)
        subprocess.run(["osascript", "-e", SAVE_AS_PDF, str(copy), str(requested)], check=True, timeout=120)
        deadline = time.time() + WAIT_SECONDS
        while time.time() < deadline:
            for landed in (LANDED, requested):
                if landed.exists():
                    shutil.move(str(landed), target)
                    return
            time.sleep(1)
        raise RuntimeError(f"PowerPoint 没有报错，但 {LANDED} 和 {requested} 都没有出现 PDF")
    finally:
        shutil.rmtree(WORK, ignore_errors=True)


def trim_white(image: Image.Image) -> Image.Image:
    pixels = np.asarray(image.convert("L"))
    rows = np.where((pixels < WHITE_THRESHOLD).any(axis=1))[0]
    columns = np.where((pixels < WHITE_THRESHOLD).any(axis=0))[0]
    if rows.size == 0:
        raise RuntimeError("这一页是空白的")
    return image.crop(
        (
            max(int(columns[0]) - PADDING, 0),
            max(int(rows[0]) - PADDING, 0),
            min(int(columns[-1]) + PADDING + 1, image.width),
            min(int(rows[-1]) + PADDING + 1, image.height),
        )
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--project", required=True, type=Path, help="专利版本目录，内含 patent.json 和 figures/")
    args = parser.parse_args()

    project = args.project.expanduser().resolve()
    figures = project / "figures"
    source = figures / SOURCE_NAME
    if not source.exists():
        print(f"找不到 {source}", file=sys.stderr)
        return 1
    locks = list(figures.glob("~$*.pptx"))
    if locks:
        print(f"PowerPoint 正开着这些文件，请先关闭：{[lock.name for lock in locks]}", file=sys.stderr)
        return 2
    if LANDED.exists():
        print(f"{LANDED} 已存在，不知道是不是这次的产物，请先确认并移走", file=sys.stderr)
        return 2

    config = json.loads((project / "patent.json").read_text(encoding="utf-8"))
    names = [figure["file"] for figure in config["figures"]]
    with tempfile.TemporaryDirectory() as temp:
        pdf_path = Path(temp) / "figs.pdf"
        export_pdf(source, pdf_path)
        document = fitz.open(pdf_path)
        if len(document) != len(names):
            print(f"PPT 有 {len(document)} 页，patent.json 列了 {len(names)} 张图", file=sys.stderr)
            return 1
        for page, name in zip(document, names):
            pixmap = page.get_pixmap(dpi=DPI, alpha=False)
            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            trimmed = trim_white(image)
            trimmed.save(figures / name, dpi=(DPI, DPI))
            print(f"已导出 {name}：{trimmed.width}×{trimmed.height}")
        document.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
