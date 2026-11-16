# Nanyan Proposal Skill

一个自包含的 Codex/Claude skill，用一份问卷 Excel 生成北京大学深圳研究生院研究生会常代会权益提案汇总。

## 包含内容

- 固定 Word 模板和版式规则
- Excel 表头识别、联系方式准入、来源行映射
- 多问题拆分、相似问题合并、措施原文依据
- 逐条近 12 个月公众号历史检索和全文核验流程
- 图片下载、真实图片校验、压缩和 Word 插入
- Word 结构/语言/来源/图片/高亮校验
- PDF 和逐页预览
- 审查 JSON、manifest 和检索日志

## 安装

将 GitHub 仓库中的 `nanyan-proposal-skill` 目录作为 skill 安装。若在本地直接运行：

```powershell
Set-Location nanyan-proposal-skill
.\install.ps1
```

## 使用

```powershell
.\run.ps1 -Xlsx "C:\path\questionnaire.xlsx"
```

输出位于 `reports/YYYY-MM/`。技能会自动从文件名或填写日期识别月份。

首次运行可能会建立一个空的 `proposal-config.json` 并返回 `needs_editorial_config`。这是给 AI 整理来源和正文的内部中间状态，不是最终结果；使用 skill 的 AI 必须继续读取 Excel、填写配置并再次运行，直到输出 Word 和审查文件。

往期 `reports/YYYY-MM/` 只能作为版式、分类和措辞参考，不能复制其中的历史依据。每个月的历史结论必须由本月逐条独立检索重新获得；没有可靠新证据时不写历史说明。

## 最强可复现模式

不同电脑的 Word 分页可能因字体、LibreOffice 和 PDF 渲染器版本不同而变化。要尽量做到像素级一致，使用仓库内的 Dockerfile：

```powershell
docker build -t nanyan-proposal .
docker run --rm -v "C:\path\input:/input" -v "C:\path\output:/skill/reports" nanyan-proposal --xlsx /input/questionnaire.xlsx
```

技能代码、模板、规则和 Python 依赖已固定。公众号历史检索由仓库内 `scripts/history_search.py` 提供：优先使用可用的 Exa MCP 路由，或使用 `EXA_API_KEY` 直连 Exa；不需要另行安装 Agent Reach skill。网络和 Exa 凭据仍是外部事实来源，不能离线内置。

## 公开仓库隐私要求

不要提交真实问卷、联系方式、生成的报告、附件或 `reports/`。`.gitignore` 已默认排除这些文件；模板和规则可以公开。
