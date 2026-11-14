# TexWord Sync · LaTeX–Word 双向协作

**人在 Word 中直接改内容，AI 通过 LaTeX 读取、理解和继续处理；人专注内容，AI 负责维护格式。**

Write directly in Word. Let AI read and process your edits through LaTeX, and manage formatting through templates.

用户可以像平常写文档一样，在 Word 中直接增删、改写正文、调整段落，也可以接受或拒绝 AI 的原生修订。同步时，程序将已保存的 Word 修改自动转换回 LaTeX，供 AI 读取并继续编辑，再把结果生成回 Word。用户自己的改写会成为后续协作的输入。

内容和格式的分工贯穿整个流程：用户提出内容与格式要求，在 Word 中专注表达；AI 通过 LaTeX、模板和生成流程维护样式、编号、引用及排版。在已接入且受支持的文档范围内，用户无需为每次内容修改重新手工整理格式。同步由工程命令触发，不是后台实时监听；首次接入和复杂对象仍有下述限制。

> **早期实验性项目。** 这个项目最初只是为满足作者个人的文档协作需求而开发，并从实际使用流程中提炼成技能。它远不够成熟和完善，既不是通用文档编辑器，也不能保证处理任意 Word/LaTeX 文档。开发者可在许可证允许的非商业范围内，根据自己的需求进一步完善、适配和验证。

**许可证：PolyForm Noncommercial 1.0.0。** 本仓库源码公开，在该许可证规定的非商业用途范围内授权使用、修改和分发；不授予商业用途许可。请阅读 [LICENSE](LICENSE) 的完整条款。由于包含非商业限制，本项目不自称符合标准开源定义。依赖项保留各自许可证。

## 能做什么

- **直接在 Word 中写作和改稿**：用户的增删改写可在同步时自动回读为 LaTeX，成为 AI 下一轮读取和处理的内容。
- **由 AI 维护格式**：用户专注内容与要求，AI 通过源模板和生成流程维护样式、编号、引用及排版。
- **持续双向协作**：将既有 Word 接入 LaTeX 工程，或以 Word 模板生成带源映射的工作稿；AI 修改后再生成回 Word。
- 合并双方不重叠的改动；重叠时保留 Word，并报告未写入的 LaTeX 建议。
- 在已支持的范围内保留引用域、可编辑公式、图片、表格和 Word 原生插入／删除修订。
- 在回写前检查编辑状态、候选文档和输入身份，保留用户接受、拒绝修订的决定。

## 前提与限制

编译需要 Python 3.10+、Pandoc，以及 [requirements.txt](assets/texword/requirements.txt) 中的 `lxml` 和 `python-docx`。实际验证使用 Python 3.12.14 与 Pandoc 2.19.2；这不是跨版本兼容性保证。完整桌面同步和原生全文比较需要 Windows、桌面 Microsoft Word 及 Windows PowerShell；当前桥接仍含 `C:/Windows` 假设。

没有 Word 的环境可使用适用的文件转换、映射回读与文件检查，但不能完成桌面状态检查和最终 Word 回写。Pandoc 最低声明版本为 2.12，复杂表格仍应使用经过项目验证的版本。

任意未映射的 Word 文档不能直接安全回写到另一份 LaTeX。复杂排版、任意自动编号、交错修订、复杂评论和未知对象并未全面支持；不支持的对象可能使流程停止。文件身份和候选检查降低风险，但不能保证任意中断后都可自动恢复。请先用独立副本验证，保留重要原件。

## 作为技能使用

将本仓库作为完整 `texword-sync` 技能目录放入你的技能目录，入口是根目录 [SKILL.md](SKILL.md)，保留 `assets`、`references` 和 `agents` 的相对位置。例如：

```powershell
git clone https://github.com/Get-it-right/texword-sync.git "$HOME/.codex/skills/texword-sync"
```

然后在支持本技能格式的助手中调用：

```text
使用 $texword-sync，将这份 Word 接入 LaTeX 双向协作流程。
使用 $texword-sync，回读 Word 的最新修改，只修改指定内容并同步回 Word。
```

已有工程优先复用自己的入口和配置，不重新接入或重建基线。

## 直接使用程序

```powershell
python -m pip install -r assets/texword/requirements.txt
python assets/texword/texword.py --help
python assets/texword/texword.py inspect "template.docx" --out "inventory.json"
```

新文档接入、正文范围、首次生成与日常回写见 [operations.md](references/operations.md)。已建立工程的日常顺序为：

```powershell
python "project/workflow.py" pull --source-only
# 在配置指定的 LaTeX 中完成修改
python "project/workflow.py" push
# 仅在程序要求暂停，且用户本次已保存并暂停之后：
python "project/workflow.py" push --write-window
```

`pull --source-only` 只读磁盘上的已保存内容，不读取 Word 未保存的编辑。`push` 自带写入前检查。返回 `needs_write_window` 是正常等待状态，退出码 0 不代表交付完成；以 `writeback_completed=true` 和当前结果为准。

## 验证范围

历史记录包含 72 项合成文档测试，涵盖合并、原生修订表示、生成与首次接入，详见 [validation.json](references/validation.json)。公开版另外验证了已更换为当前解释器的测试入口、脱敏文件和内容清单，见 [publication-checks.json](references/publication-checks.json)。

这些结果不证明真实 Word COM 在所有版本上可用，也不是任意真实文档的内容正确性保证。工作流测试中有模拟的桌面桥接；真实 Word 集成测试需显式启用，不默认操作桌面。

## 隐私与贡献

本次公开仅包含技能说明、程序源码与合成测试，没有用户的原始文档、论文草稿、模板文件、真实运行状态、私人路径、账号令牌或内部开发历史。公开提交使用仓库账号及 GitHub noreply 邮箱。

开发者可按需改进，但请先复现真实问题，在对应责任模块修复，并说明验证范围。提交 Issue 或 PR 时请用最小合成样例，先清除文档作者、批注、修订、路径、隐藏映射和其他隐私信息；不要上传未经脱敏的 DOCX、同步回执或运行日志。日后由此工具生成的工作文件可能包含嵌入源文本、映射和路径，本仓库的脱敏不等于你的生成文件会自动脱敏。

功能范围与限制见 [supported-content.md](references/supported-content.md)。欢迎非商业用途下的修复与改进；目前不承诺稳定接口、跨平台兼容、更新频率或长期维护。
