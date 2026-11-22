# 上市公司深度研报

**Equity Research Report** 是一个用于中文上市公司研究的 Codex Skill。它把选题、财报核验、盈利预测、估值复算和 Word 排版连成一套流程，适合课堂作业和投研作品集。默认交付四页以内的 A4 报告，正文与表格可在 WPS 或 Word 中编辑。

## 研究流程

| 阶段 | 产出 |
| --- | --- |
| 选题 | 选定公司、研究问题、信息截止日和估值基准日 |
| 核对资料 | 整理公告与财报，记录期间、单位、页码和计算口径 |
| 盈利预测 | 拆解销量、价格、毛利与现金流，建立悲观、基准、乐观三种情景 |
| 估值与判断 | 选择适合公司的方法，检验估值对关键假设的敏感程度 |
| 排版与复核 | 生成四页研报，复算关键数字，检查图表、来源和分页 |

求职版可加入一个具体的 AI 复核示例。例如让 AI 提取三季报和半年报的累计经营现金流，核对两份原表的期间和单位，再相减得到第三季度单季值。

## 在 Codex 中使用

将整个仓库目录作为一个 Skill 放到 `~/.codex/skills/equity-research-report/`。打开新对话，输入：

> 用 $equity-research-report 研究一家 AI 基础设施上市公司。按四页 A4 写成投研作品集，引用公司公告，给出三种盈利情景、估值敏感性和一项可复算的 AI 核验案例。输出 WPS 可编辑的 DOCX。

也可以给出公司名称、既有研报或指定数据范围，让 Skill 按相同流程完成研究或局部修订。

## 试跑四页 DOCX

在仓库根目录运行以下命令。示例使用虚构公司和模拟数据。生成脚本根据 JSON 排版，研究与计算由上游工作流完成。

```sh
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install python-docx

cat > demo.json <<'EOF'
{
  "metadata": {
    "company": "示例光通信",
    "series": "投研模拟",
    "masthead": "行业研究 · 模拟材料",
    "report_kind": "公司深度研究"
  },
  "pages": [
    [
      {"type": "title", "text": "示例光通信"},
      {"type": "subtitle", "text": "交付增长与现金回收"},
      {"type": "columns", "fractions": [0.646, 0.354], "blocks": [
        [{"type": "paragraph", "lead": "研究判断。", "text": "扩张期的现金回收速度决定盈利兑现的质量。"}],
        [{"type": "band", "text": "关注现金转化"}]
      ]}
    ],
    [
      {"type": "section", "text": "历史财务与现金流"},
      {"type": "table", "headers": ["指标（亿元）", "2025 前三季度", "2026 前三季度"], "rows": [
        ["营业收入", 72, 100], ["经营活动现金流", 8, 9]
      ]}
    ],
    [
      {"type": "section", "text": "盈利预测与三种情景"},
      {"type": "table", "headers": ["2027E", "悲观", "基准", "乐观"], "rows": [
        ["归母净利润（亿元）", 8, 10, 12]
      ]},
      {"type": "note", "text": "以上预测均为演示假设。"}
    ],
    [
      {"type": "section", "text": "估值与 AI 复核"},
      {"type": "paragraph", "lead": "倍数敏感性。", "text": "基准归母净利润为 10 亿元，采用假设市盈率 12 至 18 倍，对应股权价值 120 至 180 亿元。"},
      {"type": "note", "text": "复核示例：模拟的前三季度经营现金流 9 亿元，减去上半年 5 亿元，得到第三季度单季 4 亿元。"}
    ]
  ]
}
EOF

python3 scripts/build_report.py demo.json demo.docx
```

用 WPS 或 Word 打开 `demo.docx` 检查页面。正式版式、JSON 字段和中文字体预览方法见 [四页 WPS 版式](references/report-layout.md)。

## 仓库内容

- [SKILL.md](SKILL.md) 为 Codex 入口，包含选题、研究、排版与验收规则。
- [研究与复算](references/research-method.md) 说明财务口径、情景假设和估值检查。
- [Word 生成器](scripts/build_report.py) 将已核验的文案、财务表和图表排成可编辑的 DOCX。
- [空白 Word 版式](assets/report-shell.docx) 提供页边距、字体和版式基础。
