# user-interview-analyst

用于用户研究访谈文本分析的 Codex/Agent Skill。适合处理访谈逐字稿、整理纪要、焦点小组记录，并输出结构化洞察报告。

## 功能

- 主题分析与关键词总结
- 用户画像构建
- JTBD 分析
- 普遍需求与痛点梳理
- 隐藏需求与易遗漏洞察
- 后续追问、异常值和产品设计启发
- 可选导出 `.docx` 报告

## 文件结构

```text
.
├── SKILL.md
├── evals/
│   └── evals.json
└── scripts/
    ├── build_interview_report.py
    └── sample_report.json
```

本仓库未包含生成后的 `sample_report.docx` 二进制示例文件；可用 `sample_report.json` 重新生成。

## 生成 Word 报告

```bash
pip install python-docx
python scripts/build_interview_report.py --content-file scripts/sample_report.json --output sample_report.docx
```

## 安装方式

将本仓库内容放入你的 Skill 目录，例如：

```text
C:\Users\<你的用户名>\.agents\skills\user-interview-analyst
```

然后重启 Codex/Agent 环境以加载 Skill。
