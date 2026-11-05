# scripts 用法

全部脚本用同一个 Python 环境跑（见 `../config.local.md`）。依赖：`python-docx`、`lxml`、`latex2mathml`、`Pillow`；导出附图还要 `pymupdf`、`numpy`。

## 骨架与配置

| 文件 | 作用 |
|---|---|
| `template.md` | 空白 md 骨架。**新专利从它复制**到 `<专利>/v<N>/草案.md`，不要从零写 |
| `patent.json.example` | 项目配置样例，复制到版本目录改名 `patent.json` |
| `patent_config.py` | 三个生成类脚本共用的配置加载，不单独运行 |
| `docx_common.py` | 三个 docx 工具共用的底层读写，不单独运行 |

`patent.json` 放在专利版本目录里，脚本靠 `--project` 找到它：

```json
{
  "title": "一种……的方法",
  "source_md": "草案.md",
  "template_dir": "../../模板",
  "figures": [{"file": "flow.png", "caption": "图1 实现步骤流程图"}],
  "abstract_figure": "flow.png"
}
```

目录约定（相对版本目录，不可配置）：`figures/` 放附图 PNG，`word/` 是输出，`snapshots/` 放只读快照。

## 导出附图

```bash
# figures/附图合集.pptx → patent.json 里列的 PNG，300 dpi，裁白边
python export_pptx_png.py --project ~/Desktop/专利/某专利/v1
```

只支持 macOS，本机要装 PowerPoint。PPT 页序要和 `figures` 列表一致，页数对不上会报错。`figures/` 下有 `~$*.pptx` 锁文件时拒绝运行，要先关掉 PowerPoint 里的这个文件。

## 成稿三件套

```bash
# md 草案 → 五个标准件 docx
python build_docx.py --project ~/Desktop/专利/某专利/v1

# 格式体检（字体字号行距、OMML 参数槽、页眉页脚、图题……）
python verify_docx.py --project ~/Desktop/专利/某专利/v1

# docx → 只读 md 快照，供 git diff 阅读。输出幂等，可反复跑
python export_snapshot.py --project ~/Desktop/专利/某专利/v1
```

`build_docx.py` 只在**首次成稿**时用。事实源转正为 docx 之后不要再跑，会覆盖手改（见 `../references/05-事实源与快照.md`）。

## 文风校订的校验

```bash
python check_style_invariants.py 草案.改前.md 草案.md --terms 术语表.txt
```

比较改写前后两份 md，七项全过才算只动了表达：数字；公式与表格；权利要求书整节；术语出现次数；S1–Sn 编号以及技术方案步骤段不变；冻结句式、技术领域和图号；摘要字数与权项句号。

用户决定修改权利要求后加 `--allow-claims`，术语确需增减时加 `--allow-term 某词`（可重复）。退出码 0 为全过，1 为有项不过，2 为输入有问题。

它查不出句子含义的变化，人工审核的要点见 `../references/文风校订.md`。

## docx 批注与修订三件套

```bash
# 摸清一轮返修提了什么：批注、修订作者分布、people.xml 里的账户名
python read_revisions.py 稿子.docx
python read_revisions.py 稿子.docx --full     # 连每处修订的文本
python read_revisions.py 稿子.docx --json     # 机读

# 两份稿子比对：这一轮新增了哪些修订
python diff_docx.py 上一轮.docx 这一轮.docx
python diff_docx.py 旧.docx 新.docx --author 某作者 --paragraphs

# 以 Word 修订模式写入改动
python write_tracked_change.py 稿子.docx --anchor "原文片段" --replace "新文字" --dry-run
python write_tracked_change.py 稿子.docx --anchor "原文片段" --replace "新文字"
python write_tracked_change.py 稿子.docx --anchor "某段结尾" --append "补充一句。"
python write_tracked_change.py 稿子.docx --anchor "要删的片段" --delete
```

`write_tracked_change.py` 的四条安全约束，不要绕过：

1. **锚点命中必须恰好 1 次**，否则中止。命中 0 次时会提示是否跨 run，换个更短的锚点，别放宽匹配。
2. **写入前查 Word 锁文件**，检测到就中止。
3. **作者名默认从 `people.xml` 反解**，取不到唯一值时必须用 `--author` 指定 Word 当前登录账户名。
4. **自动双向验证**：接受修订后结果正确、拒绝修订后逐字回到原样，任一不过就不写文件。

先用 `--dry-run` 看命中位置，再真写。真写会自动备份为 `<原名>_备份_改前_<时分秒>.docx`。

**公式类改动风险最高**，脚本改完仍要真开 Word 目视确认渲染没降级。
