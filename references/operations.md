# 接入与运行

以下为 PowerShell 示例。`$engine` 替换为本技能实际目录下 `assets/texword` 的绝对路径；`python` 使用装有依赖的解释器。所有文档路径为调用者输入。引擎无 pip 包安装步骤；缺少 Python 依赖时按当前环境规则安装 `requirements.txt`。Pandoc 单独安装或通过 `PANDOC` 指定；不擅自降级系统版本。

## 既有工程

读取工程入口和 `workflow.json`，复用其引擎、`source`、`adapter`、`code`、`word`、可选 `bibliography` 和 `pandoc`。不要替换既有专用调用方。`PANDOC` 显式环境变量优先于工程配置。

```powershell
python "$project/workflow.py" pull --source-only
# 编辑配置指定的 LaTeX
python "$project/workflow.py" push
# 仅在本次用户已保存并暂停后：
python "$project/workflow.py" push --write-window
```

`status` 是保存文件预览，会写 `sync_check.json`；`preflight` 是桌面状态诊断，不保存。`check` 检查文件；`check --word` 会打开／刷新域／保存。`build` 强制重编译，适合仅改 adapter，仍走完整保护。`pull` 是会按需重建映射的完整同步，不等于无桌面副作用的 `pull --source-only`。

## 未映射 Word 接入

先检查结构；正文范围有实质歧义时向用户确认，不能按页码猜 XML 索引。新工程和工作稿路径应独立于原件。

```powershell
python "$engine/texword.py" inspect "$original" --out "$inventory"
python "$engine/texword.py" adopt "$original" "$project" --output-word "$workingWord" --body-start 1 --body-end 8
python "$engine/texword.py" compile "$project/template_adapter" "$project/report.tex" "$project/generate_word.py" --roundtrip
python "$project/generate_word.py" "$workingWord"
python "$project/workflow.py" check --baseline
# 需要桌面 Word 字段刷新和确认文件身份时：
python "$project/workflow.py" check --word
```

`1`、`8` 只是示例，必须换成 inspect 确认的 XML body 子节点范围。`adopt` 的范围参数是整数；`adapt` 另有唯一文字匹配参数。`adopt` 只提取与写配置，不等于工作稿已生成；必须先 compile 并执行生成器建立带映射的工作稿，不能直接对不存在的工作稿调用 workflow build。`check --word` 会操作桌面 Word，只在请求已授权时执行。`check --baseline` 仅用于首次接入原件保全比较，后续修改后不用其要求正文仍等于初稿。

生成的工程 `workflow.py` 记录引擎的绝对路径；移动技能或工程到另一机器时，更新该薄入口的导入路径以及调用者路径配置，不改映射哈希或伪造同步身份。

## 已有 LaTeX 与模板

```powershell
python "$engine/texword.py" inspect "$template" --out "$inventory"
python "$engine/texword.py" adapt "$template" "$adapter" --body-start "正文起点唯一文字" --body-end "页尾起点唯一文字"
python "$engine/texword.py" compile "$adapter" "$source" "$generator" --bibliography "$bib" --roundtrip
python "$generator" "$workingWord"
```

没有参考文献时省略 `--bibliography`。默认不覆盖已有派生输出；确需刷新本工程派生文件时用其支持的 `--overwrite`。原始模板永远不能作为生成目标。支持 `--body-start-index`／`--body-end-index` 和 `--sample body=22 --sample heading1=19`；这些值须来自实际模板。重复 `adapt` 会重建 adapter，不能覆盖已维护的专用格式代码。

若要持续桌面交互，在新建调用方中建立包含绝对路径 `source`、`adapter`、`code`、`word` 及可选 `bibliography`／`pandoc` 的 `workflow.json`，以薄 `workflow.py` 导入 `workflow_runtime.main(Path(__file__).resolve().parent)`。这是程序初始化配置的工作，不是人工修改已有工程同步身份。不要调用 `write_workflow` 去覆盖已有工程；它为 Word adoption 的默认 report.tex 布局设计。

没有固定工程时，底层 `sync` 可预览或导入，但不维护完整工程回执：

```powershell
python "$engine/texword.py" sync "$workingWord" "$source"
python "$engine/texword.py" sync "$workingWord" "$source" --apply
```

`--output <new.tex>` 与 `--apply` 互斥；输出必须为新文件。必须已有匹配的源映射，不能将未映射 Word 猜配到源文件。

## 恢复与证据

| 结果 | 行动 |
|---|---|
| `needs_write_window` | 保存并暂停的确认到达前停止回写，内容准备已完成 |
| `passed` + `source_only=true` | 导入完成，尚未交付 Word |
| `failed` / `generate` | 保留成功导入，修复编译／生成原因 |
| `failed` / `validate_candidate` | 当前 Word 保留，查看候选检查原因 |
| `failed` / `refresh_and_check` | 看 `word_replaced`；若已替换，修复桌面问题后 `check --word` |
| `conflict` / `needs_attention` | 查具体输入与对象，不强制重置基线 |
| `writeback_completed=true` | 通知已完成与可继续编辑，列出跳过建议 |

读取当前 `workflow_result.json`、`sync_check.json`、`check_result.json`，桌面证据为 `word_check.json`。哈希必须对应本次实际输入，旧 PASS 不能证明新文件。编译产出的 Python + 完整 `_assets` 可独立生成 Word，不再读取 TeX/BibTeX 或调用 Pandoc；这不等于生成阶段替代了完整交互工程的保护。

## 验证入口

从任意 cwd，以所选引擎的绝对路径运行。例如：

```powershell
python -m unittest discover -s "$engine/tests" -p test_workflow.py -v
python -m unittest discover -s "$engine/tests" -p test_tracked_reviews.py -v
python -m unittest discover -s "$engine/tests" -p test_texword.py -v
```

按改变的行为选择，不例行全跑。`test_workflow` 覆盖并行合并、等待窗口、只读准备、失败恢复；其 Word 桥接为模拟。测试生成的子程序使用当前 Python 解释器；其余环境依赖仍需根据实际结果确认。
