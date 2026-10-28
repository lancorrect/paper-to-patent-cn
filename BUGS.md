# 已知 bug：脚本中残留的专利专属硬编码

- 发现日期：2026-10-27（首次将本 skill 用于第二项专利时暴露）
- 影响范围：`scripts/build_docx.py`、`scripts/verify_docx.py`
- **本目录（源码）未打补丁**——2026-10-27 经用户决定"代码无需同步"。
  已安装于 `~/.claude/skills/paper-to-patent-cn/scripts/` 的副本**已修复**。
  若日后要同步：直接用已装副本的两个文件覆盖本目录同名文件即可，或按文末 diff 逐条应用。

## 根因

这套脚本是从首次真实申请中直接长出来的，作者把"当时恰好成立的数字"写成了断言，
与 README 中"脚本里不含任何具体专利的内容"的声明矛盾。**换任何一篇论文必踩。**

## bug 清单（12 处，8 个逻辑问题）

| # | 位置 | 写死了什么 | 违背什么 | 修正为 |
|---|---|---|---|---|
| 1 | build · parse_source | 权项必须 `**N、**` 加粗格式 | `scripts/template.md` 示例是不加粗的 `1、`（文档自相矛盾） | `\*{0,2}` 兼容两种 |
| 2 | build · parse_source | 权项必须恰好 11 条（`range(1,12)`） | README 通用性声明 | 从 1 起连续即可 |
| 3 | build · parse_source | 独立公式必须恰好 51 个 | 同上 | ≥1 |
| 4 | build · parse_source + build_description | 表格必须恰好 2 张 | 同上 | ≥1 |
| 5 | build · add_html_table | 表必须 9 行 × 13 列 | 同上 | 任意非空 |
| 6 | build · add_html_table | 列宽 = 22mm 首列 + 12 等分列 | 同上 | 22mm 首列 + (n−1) 等分；单列时取满宽 |
| 7 | build · add_html_table | 表头跨页重复标记前两行（原表双行表头） | 同上 | 仅首行 |
| 8 | build · build_description + verify · check10 | 表题只认"表一 /表二 "（汉字+空格） | `references/02-撰写规范.md` 第 4 节自己沉淀的代理机构意见"序号用阿拉伯数字" | 正则 `^表\d+、` |
| 9 | verify · check3 | 权项必须 1–11 共 11 条 | README 通用性声明 | 从 1 起连续（且 ≥2） |
| 10 | verify · check5 | 表格必须 2 张，且各表行列必须完全一致 | 同上（多表形状不同即误报） | ≥1 张；行列数降级为仅报告 |
| 11 | verify · check6 | 说明书附图必须恰好 4 张 | patent.json 里明明有 figures 清单却不用 | 取 `len(EXPECTED_CAPTIONS)` |
| 12 | verify · check10 | 正文段落检查的表题排除项硬编码"表一 /表二 " | 同 #8 | 正则 `^表\d+、` |

修正原则只有两类，不引入新的具体数字：**从输入推导期望值**（如附图数取 patent.json），
或**只校验结构性质**（连续编号、非空）。

## 非 bug 但值得知道的限制

| 项 | 说明 |
|---|---|
| 公式转换器不支持 `mtable`/`cases` | `\begin{cases}` 与矩阵会退化为斜体纯文本。本例的解法：按"公式拆短"原则改写为条目式（（1）–（N）），数学记号走行内公式。若未来必须用矩阵/分段，需给 MathML→OMML 转换器补 `mtable → m:m` 支持 |
| `摘要附图.docx` 页眉沿用"说明书附图/100003" | 源自首次申请的交付形态；代理机构是否接受需逐案确认 |
| `100001权利要求书` 模板是旧版 .doc | build 借用 100002 版式并替换注意事项文本，属权宜之计（脚本注释已自述） |

## 完整补丁 diff（源码 → 已修复版）

以下 `---` 为本目录源码，`+++` 为已修复的已装副本。同步时整体应用即可。

```diff
--- a/scripts/build_docx.py	2026-09-12 11:54:57
+++ b/scripts/build_docx.py	2026-10-27 15:50:02
@@ -546,10 +546,12 @@
 
     claim_numbers = [
         int(value)
-        for value in re.findall(r"^\*\*(\d+)、\*\*", claims_section, flags=re.MULTILINE)
+        for value in re.findall(r"^\*{0,2}(\d+)、\*{0,2}", claims_section, flags=re.MULTILINE)
     ]
-    if claim_numbers != list(range(1, 12)):
-        raise ValueError(f"expected claims 1-11, got {claim_numbers}")
+    if not claim_numbers or claim_numbers != list(range(1, len(claim_numbers) + 1)):
+        raise ValueError(
+            f"claims must be numbered consecutively from 1, got {claim_numbers}"
+        )
 
     description_headings = re.findall(r"^###\s+(.+?)\s*$", description_section, flags=re.MULTILINE)
     if description_headings != EXPECTED_DESCRIPTION_HEADINGS:
@@ -559,12 +561,12 @@
     source_formula_blocks = re.findall(
         r"^\$\$\s*\n(.*?)\n\$\$\s*$", relevant, flags=re.MULTILINE | re.DOTALL
     )
-    if len(source_formula_blocks) != 51:
-        raise ValueError(f"expected 51 display formulas, got {len(source_formula_blocks)}")
+    if not source_formula_blocks:
+        raise ValueError("expected at least one display formula, got 0")
 
     table_count = len(re.findall(r"^<table>\s*$", description_section, flags=re.MULTILINE))
-    if table_count != 2:
-        raise ValueError(f"expected two HTML tables, got {table_count}")
+    if table_count < 1:
+        raise ValueError(f"expected at least one HTML table, got {table_count}")
 
     return {
         "title": title,
@@ -1333,9 +1335,9 @@
 
 def add_html_table(document: DocumentObject, table_html: str, table_number: int) -> None:
     row_count, column_count, specs = parse_html_table(table_html)
-    if row_count != 9 or column_count != 13:
+    if row_count < 1 or column_count < 1:
         raise ValueError(
-            f"table {table_number}: expected 9 rows and 13 columns, got {row_count}x{column_count}"
+            f"table {table_number}: expected a non-empty table, got {row_count}x{column_count}"
         )
 
     table = document.add_table(rows=row_count, cols=column_count)
@@ -1345,8 +1347,11 @@
     section = document.sections[0]
     available_width = section.page_width - section.left_margin - section.right_margin
     first_column_width = int(Mm(22))
-    data_column_width = int((available_width - first_column_width) / 12)
-    widths = [first_column_width] + [data_column_width] * 12
+    if column_count > 1:
+        data_column_width = int((available_width - first_column_width) / (column_count - 1))
+        widths = [first_column_width] + [data_column_width] * (column_count - 1)
+    else:
+        widths = [int(available_width)]
     for column_index, width in enumerate(widths):
         table.columns[column_index].width = width
         for row in table.rows:
@@ -1388,7 +1393,6 @@
         )
 
     mark_header_row_repeat(table.rows[0])
-    mark_header_row_repeat(table.rows[1])
 
 
 def add_picture_scaled(document: DocumentObject, image_path: Path, label: str | None = None) -> None:
@@ -1517,7 +1521,7 @@
         if token.kind == "heading":
             add_heading(document, token.value)
         elif token.kind == "paragraph":
-            centered = token.value.startswith("表一 ") or token.value.startswith("表二 ")
+            centered = re.match(r"^表\d+、", token.value) is not None
             paragraph = add_body_paragraph(document, token.value, context, centered=centered)
             if centered:
                 paragraph.paragraph_format.keep_with_next = True
@@ -1528,8 +1532,8 @@
             table_count += 1
             add_html_table(document, token.value, table_count)
 
-    if table_count != 2:
-        raise ValueError(f"description build expected two tables, got {table_count}")
+    if table_count < 1:
+        raise ValueError(f"description build expected at least one table, got {table_count}")
     save_atomic(document, OUT_DESCRIPTION)
     return block_count
 
```

```diff
--- a/scripts/verify_docx.py	2026-09-12 11:54:57
+++ b/scripts/verify_docx.py	2026-10-27 15:53:14
@@ -333,9 +333,9 @@
             if match:
                 claim_numbers.append(int(match.group(1)))
         details.append(f"识别到的权利要求编号={claim_numbers}")
-        if claim_numbers != list(range(1, 12)):
-            errors.append("权利要求编号不是完整的 1 至 11")
-    results.append(CheckResult(3, not errors, "权利要求书含 1 至 11 共 11 条", details + errors))
+        if len(claim_numbers) < 2 or claim_numbers != list(range(1, len(claim_numbers) + 1)):
+            errors.append("权利要求编号不是从 1 起的连续编号")
+    results.append(CheckResult(3, not errors, "权利要求书从 1 起连续编号", details + errors))
 
     # 4. Report OMML counts in claims and description; reject empty formulas.
     details = []
@@ -382,8 +382,8 @@
         document = documents[name]
         tables = document.tables
         details.append(f"w:tbl 数量={len(tables)}")
-        if len(tables) != 2:
-            errors.append(f"说明书表格数应为 2，实际为 {len(tables)}")
+        if len(tables) < 1:
+            errors.append(f"说明书表格数应至少为 1，实际为 {len(tables)}")
         dimensions: list[tuple[int, int]] = []
         for index, table in enumerate(tables, start=1):
             row_count = len(table.rows)
@@ -415,10 +415,8 @@
                 )
             if empty_header_columns:
                 errors.append(f"表{index} 存在空表头列: {empty_header_columns}")
-        if len(dimensions) > 1 and len(set(dimensions)) != 1:
-            errors.append(f"各表行列数不一致: {dimensions}")
-        elif dimensions:
-            details.append(f"各表行列数一致={len(set(dimensions)) == 1}: {dimensions}")
+        if dimensions:
+            details.append(f"各表行列数={dimensions}")
     results.append(
         CheckResult(5, not errors, "说明书表格数量、行列一致性及非空表头自检", details + errors)
     )
@@ -426,7 +424,10 @@
     # 6. Drawing and media counts.
     details = []
     errors = []
-    for name, expected in (("100003说明书附图.docx", 4), ("摘要附图.docx", 1)):
+    for name, expected in (
+        ("100003说明书附图.docx", len(EXPECTED_CAPTIONS)),
+        ("摘要附图.docx", 1),
+    ):
         if name not in documents:
             errors.append(f"{name} 未成功打开")
             continue
@@ -810,8 +811,7 @@
         for paragraph in description_paragraphs
         if paragraph_text(paragraph)
         and paragraph_text(paragraph) not in excluded_description_paragraphs
-        and not paragraph_text(paragraph).startswith("表一 ")
-        and not paragraph_text(paragraph).startswith("表二 ")
+        and not re.match(r"^表\d+、", paragraph_text(paragraph))
     ]
     description_runs = [
         run for paragraph in description_body_paragraphs for run in visible_runs(paragraph)
```

