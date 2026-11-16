# 常代会权益提案生成系统

本项目通过 `AGENTS.md` 固化整理规则。日常使用时只需提供当月问卷 Excel；系统会识别月份、建立独立目录、整理提案、下载图片、套用固定模板，并在校验和逐页检查通过后交付 Word。

## 目录

- `AGENTS.md`：长期规则与质量门槛。
- `templates/proposal-template.docx`：从已通过成稿提取的固定版式模板。
- `reports/YYYY-MM/`：每月互相隔离的来源快照、整理配置、成稿、附件、审查记录和逐页预览。
- `scripts/monthly_report.py`：只传 Excel 的统一入口。
- `scripts/build_proposal.py`：生成 Word 和审查记录。
- `scripts/validate_proposal.py`：检查内容、来源、联系方式、禁用语言、样式、高亮和图片压缩。
- `scripts/inspect_inputs.py`：只读检查 Excel 与 Word 模板。
- `config/`、`output/`：旧版兼容目录；新月份不再写入这里。

## 日常使用

把 Excel 交给本项目中的 Agent 即可，不需要再提供模板、配置文件或输出名称。Agent 必须完整执行 `AGENTS.md` 中的整理、历史检索和逐页验收流程，最后返回 `reports/YYYY-MM/` 下的 Word。

统一入口的内部命令为：

```powershell
python scripts/monthly_report.py --xlsx "问卷.xlsx"
```

首次处理某个月份时，命令会建立目录和待整理配置。Agent 完成内容整理后再次运行同一入口，系统会生成、校验并渲染 Word。若文件名没有年月，系统会从问卷填写日期识别；两者冲突时会停止，防止错放月份。

## 月度目录

```text
reports/2026-09/
├── source/                         原始 Excel 快照
├── attachments/original/           下载的原图
├── attachments/word/               插入 Word 的压缩图
├── working/preview/                 PDF 与逐页图片
├── proposal-config.json             当月整理与来源映射
├── manifest.json                    输入、模板和成稿指纹
├── 北京大学…2026年9月提案汇总.docx
└── 北京大学…2026年9月提案汇总.audit.json
```

不同内容但同名的 Excel 不会覆盖既有快照，而会附加内容指纹另存。每条提案仍须填写 Excel 的真实 `source_rows`；联系方式由脚本直接从 Excel 读取，不手工抄写。

问卷“相关图片”列中的明文链接或单元格超链接会自动下载到当月 `attachments/original/`，压缩副本存入 `attachments/word/` 并插入 Word。有效问卷的图片无法下载时，生成过程会直接报错。

每条提案必须填写 `measure_basis`，逐项指向问卷中的问题或建议原文。生成器会在输出 Word 前拦截未登记依据、来源行重复、防御性套话以及“临时警示、故障反馈渠道、夜间转介机制、备用通行”等自行补充措施。有效但经人工决定不纳入本期成稿的问卷，必须写入 `excluded_rows` 并说明原因。

同一问卷含多个独立问题时，应按对象和措施拆成独立提案；如果因此复用 Excel 来源行，必须在 `allowed_reused_rows` 写明拆分理由。每条提案的历史检索都要记录在 `historical_search_log`，包括检索词、时间窗、结果状态和是否采用。

对常见或疑似已处理的问题，按 `AGENTS.md` 中的“微信公众号历史检索”使用仓库内 `scripts/history_search.py` 搜索“南燕研究生会”公开文章。只引用问卷月份向前12个月内且证据明确的记录；存在多个月份时，可以在 `history_reference` 中并列月份，但不写条号或具体提案名称。具体题名、公众号、检索词、证据类型和定位信息只保存在 `history_sources` 内部审查记录中。搜索摘要未展示的内容不得推断或补写。

微信公众号检索的当前命令为：先用 `python scripts/history_search.py search` 搜索，再用 `python scripts/history_search.py fetch --url ...` 抓取文章全文。脚本内部优先使用 Exa MCP，或使用 `EXA_API_KEY` 直连，不需要安装其他 skill。

重新渲染前会清理当月 `working/preview/` 的旧 PDF 和页图，防止旧页混入 manifest；逐页验收按数字页码顺序进行。
