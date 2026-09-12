# paper-to-patent-cn

把一篇学术论文改写成可交付专利代理机构的中国发明专利申请文件。

这是一个 [Claude Code](https://claude.com/claude-code) skill。它覆盖从论文分析到 Word 五标准件打包的完整链路，并带六个可直接运行的 Python 脚本。

## 它是怎么来的

每一条规则都来自一次真实的申请：一篇深度学习方法论文，经七个月改写、两轮代理机构返修后递交 CNIPA。文档里凡是写「代理机构实际改过」「实测」「踩过的坑」的地方，都是实际发生过的意见或实际造成过返工的错误，不是推测。

所以它的目标不是「能写出一份专利」，而是**一次写成的稿子就达到代理机构两轮意见改完之后的水平**。

## 安装

```bash
git clone <这个仓库> ~/.claude/skills/paper-to-patent-cn
# 或者装到某个项目里
git clone <这个仓库> <项目>/.claude/skills/paper-to-patent-cn
```

然后复制配置模板并填好：

```bash
cd ~/.claude/skills/paper-to-patent-cn
cp config.local.md.example config.local.md
```

`config.local.md` 装本机与个人相关的一切（Word 登录账户名、Python 环境路径、参照样例在哪、代理机构信息）。它已被 `.gitignore` 忽略，不会入库。skill 的其余文件一律不写死个人信息。

依赖：`python-docx`、`lxml`、`latex2mathml`、`Pillow`。

## 六个阶段

| 阶段 | 做什么 | 文档 |
|---|---|---|
| 1 论文分析与骨架映射 | 论文 → 映射表 + 权利要求结构 | `references/01-论文映射分析.md` |
| 2 撰写 | 映射表 → 五件全文 md 草案 | `references/02-撰写规范.md` |
| 3 附图 | 结构规格 → drawio 合集 → 导出图 | `references/03-附图.md` |
| 4 Word 五件打包 | md → 五个 docx | `references/04-Word打包.md` |
| 5 事实源转正与快照 | docx 成为事实源，md 快照供 diff | `references/05-事实源与快照.md` |
| 6 工作流与日志 | 版本目录、日志、提交 | `references/06-工作流与日志.md` |

其中 `01` 和 `02` 是重点。`01` 讲怎么用专利结构当模具去拆论文，以及**哪些实现细节绝对不能从论文推断**；`02` 最厚，把代理机构的全部意见、格式基准、CNIPA 硬规则化成了「怎么写」的正面要求。

## 脚本

```bash
# 成稿：md 草案 → 五个标准件 → 格式体检 → 只读快照
python scripts/build_docx.py      --project <专利>/v1
python scripts/verify_docx.py     --project <专利>/v1
python scripts/export_snapshot.py --project <专利>/v1

# 返修：读批注与修订 / 比对两轮稿子 / 以 Word 修订模式写入
python scripts/read_revisions.py 稿子.docx
python scripts/diff_docx.py 上一轮.docx 这一轮.docx
python scripts/write_tracked_change.py 稿子.docx --anchor "原文片段" --replace "新文字" --dry-run
```

前三个从 `<专利>/v<N>/patent.json` 读发明名称与附图清单，脚本里不含任何具体专利的内容。写新专利时从 `scripts/template.md` 复制 md 骨架，从 `scripts/patent.json.example` 复制配置。

详见 `scripts/README.md`。

### 三个 docx 工具解决的实际问题

- `read_revisions.py`——读批注时**同时收 `w:t` 与 `m:t`**。只收 `w:t` 会把批注里嵌的公式静默丢掉，实际发生过，导致一度看不懂批注在问什么。
- `write_tracked_change.py`——四条安全约束：锚点命中必须恰好 1 次否则中止；写入前必查 Word 锁文件；`w:author` 必须是 Word 当前登录账户名（否则 Word 把改动显示成「别人给的建议」）；改完自动做接受/拒绝双向验证，任一不过就不写文件。
- `diff_docx.py`——按「作者 + 文本」聚合比对，**不按 `w:id`**。Word 保存会重编 id，实测一天之内重编过两次。

## 四条红线

1. 实验数据不杜撰。临时占位必须显著标注，递交前必须替换或整段删除。
2. 动 docx / drawio 前先确认没人开着。开着时写入零报错、校验全过，但用户一保存就把改动静默覆盖，事后 `lsof` 也查不出异常。
3. 实现细节必须向作者核实，不得从论文推断。论文为省篇幅一定会略去实现。
4. git 提交前必须经用户审核。

## 它沉淀不了什么

- **技术方案的真实细节**——只能规定必须问。
- **技术判断**——损失函数用哪种、某个细节写不写进说明书、术语叫什么、超参取多少，都得人来定。这些被设计成审核门，不是默认动作。
- **格式基准会过期**——参照样例可能是十几年前的旧表格，官方模板会改版。
- **代理机构的偏好只代表那一家。**
- 本流程假设**方法类论文 → 方法专利**。系统类、装置类的权利要求布局与附图规格都要另行调整。

## 免责

这是撰写辅助工具，不是专利代理人或律师的替代品。新颖性、创造性、侵权、申请时机这类法律判断，本 skill 不下结论。正式申请文件应当由有资质的专利代理人审核后再提交。
