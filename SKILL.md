---
name: texword-sync
description: "支持 LaTeX 与 Word 双向协作：接入既有 DOCX 或 LaTeX/Word 模板工程，回读用户编辑、合并双方修改，并保留引用、公式、图片与原生修订。用于持续往返编辑和审阅；单次格式转换通常无需本技能。"
---

# TexWord Sync · LaTeX–Word 双向协作

助手在 LaTeX 修改正文，用户在 Word 编辑与接受／拒绝修订。程序负责 Word 回读、共同基线、三方合并、生成代码、候选校验和回写。目标是持续协作而非一次转换。

本技能由作者的个人文档协作需求驱动，属于早期实验性实现，远未成熟或完善。开发者可在许可证允许的非商业范围内按需改进，使用前应以副本验证。许可证：PolyForm Noncommercial 1.0.0，见仓库 LICENSE。

## 选择引擎与工程

先找文档已有的 `workflow.py`／项目调用入口、`workflow.json` 和源 adapter，复用已有 owner 与项目配置。不要为已接入文档重新 adopt、重建基线或换用另一套引擎。

新工程使用本技能附带的 [引擎](assets/texword/texword.py)；其运行文件和测试均随技能提供，无需访问原开发工作区。先阅读 [接入与命令](references/operations.md)。运行时路径相对本技能解析，不把当前用户、论文标题、正文范围或输出路径写成通用默认值。文档、模板、参考文献、格式及输出位置由调用者指定；首次提取可能产生 `report.tex` 等默认文件名，以生成的配置为准。

- 已有映射工程：直接进入日常流程。
- 未映射的既有 Word：`inspect` 定位正文范围，再 `adopt` 到新工程、生成独立工作稿。
- 已有 LaTeX 和 Word 模板：`adapt` + `compile --roundtrip`，先建立可回读的工作稿；完整交互需配置工程入口，见参考文档。
- 任意未映射 DOCX 不可直接当作另一份 LaTeX 的可靠回写源。

编译需要 Python、`lxml`、`python-docx>=1.2,<2` 和 Pandoc。完整桌面回写与 Word 原生比较依赖 Windows、桌面 Microsoft Word 和当前实现的 Windows PowerShell 桥接；其路径目前按 `C:/Windows` 设置。没有 Word 的环境只能使用适用的文件转换、映射回读和候选检查，不能宣称完成桌面同步。Pandoc 声明最低 2.12；复杂表格应复用项目已验证版本，2.19.2 是来源工程已有验证的版本，不承诺所有版本等价。

## 日常编辑

1. `workflow.py pull --source-only` 导入磁盘上已保存的 Word。它不调用桌面 Word，不要求用户暂停；结果仅覆盖 `saved_file_only`，不代表读取未保存内容。
2. 只在配置指定的 LaTeX 正文中完成本次修改。格式改动在源 adapter；不要人工改 Word、生成代码、`_assets` 或同步状态。沿用项目已授权的修订方式；用户要求审阅时使用原生删除／插入，不用批注建议替代。
3. 准备完成后直接执行一次 `workflow.py push`，复用其生成前内置检查，不例行先跑独立 `preflight`。
4. 如果返回 `needs_write_window`，说明具体原因，请用户本次保存并暂停编辑。收到本次确认后直接 `push --write-window`，由程序完成导入、合并、生成、检查、回写及光标恢复；不能把这个标志当作后续长期授权。
5. 读 JSON 结果和当前 `workflow_result.json`。退出码 0 也可能仅表示等待。`writeback_completed=true` 才表示工程本次交付完成，应立即通知用户可以继续编辑；同时列出 `skipped_latex_edits` 中未写入的建议。不要为已通过的本次检查追加重建或重复回读。

小改动围绕目标词句、表格或样式收敛，复用已通过的检查；超过数分钟尚无可交付结果时指出唯一阻塞和最短路线，不扩成全文整理。

## 合并、修订和保护

- 同步程序默认自动合并不重叠改动；重叠时保留 Word，报告被覆盖的 LaTeX 建议。不增加手工逐项冲突裁定。若项目需要另一策略，先确认需求并在唯一引擎 owner 实现，不能悄悄换策略。
- 保留用户接受、拒绝以及分别处理插入／删除的决定。普通批注的删除或解决不等于采纳建议；不能自动接受所有修订来方便回读。
- 原始模板、原图字节、引用域、共同基线与用户最新编辑须保留。未知或不支持对象在写入前停止，不能静默压平或丢弃。
- 引擎先生成临时候选并检查包、引用、映射和回读，再核对双方身份、替换目标并经 Word 完成检查。禁止手工改哈希、清理导入回执或修补 DOCX/XML 来越过失败。
- 已成功导入而后生成失败：保留该有效导入，修复真实原因后重试。检查失败时结合 `phase` 和 `word_replaced` 判断当前 Word 是否已替换；替换后桌面刷新失败按恢复路线 `check --word`，不要反复生成。
- `.workflow.lock` 保护同时写入；发现遗留锁先核实进程与状态，不删除活动锁。不能承诺任意进程中断都可自动恢复。

## 按需参考与维护

- 接入、模板、状态恢复与实际命令：[operations.md](references/operations.md)。
- 需要原生修订、复杂公式、图片或表格时：[supported-content.md](references/supported-content.md)。
- 本仓库内引擎位于 `assets/texword`；[engine-snapshot.json](references/engine-snapshot.json) 记录相对位置与逐文件 SHA256。已有文档工程继续复用其选定引擎，明确维护入口后再修改；本仓库不需要访问作者的私人开发目录。修改后刷新对应哈希，不保留个人机器路径。
- 本次验证环境、通过项目和局限见 [validation.json](references/validation.json)。
- 验证围绕实际改动使用附带合成 fixture。工作流测试的桌面调用是模拟的，不作为真实 Word COM 证据。交付说明区分技能结构检查、文件往返验证、真实桌面操作和用户验收。
