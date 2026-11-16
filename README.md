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

## 最强可复现模式

不同电脑的 Word 分页可能因字体、LibreOffice 和 PDF 渲染器版本不同而变化。要尽量做到像素级一致，使用仓库内的 Dockerfile：

```powershell
docker build -t nanyan-proposal .
docker run --rm -v "C:\path\input:/input" -v "C:\path\output:/skill/reports" nanyan-proposal --xlsx /input/questionnaire.xlsx
```

技能代码、模板、规则和 Python 依赖已固定；公众号历史检索仍需要网络和 Exa 路由或凭据，这是外部事实来源，不能离线内置。

## 公开仓库隐私要求

不要提交真实问卷、联系方式、生成的报告、附件或 `reports/`。`.gitignore` 已默认排除这些文件；模板和规则可以公开。
