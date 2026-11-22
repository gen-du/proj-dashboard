# project-dashboard

把《财政项目推进总台账》Excel 一键生成「项目推进看板」的 WorkBuddy Skill。

> 读取台账 → 按「项目类别」筛选工作包 → 按「工作包类别」分类统计 → 生成单文件 HTML 看板（含三类图表）→ 可选导出 Word。

## 功能特性

- **里程碑甘特图**：只列「里程碑 = 是」的工作包，按计划完成日期升序，用工作包名称标示，颜色区分已完成 / 进行中 / 未开始，红色虚线标今日、红色描边标逾期。
- **工作包类别饼图**：按类别统计总数并标注占比。
- **完成情况直方图**：每类一根柱，柱内绿色（完成）/ 灰色（未完成）分色 + 百分比，柱顶标注总数。
- **KPI 指标卡**：总数 / 已完成 / 进行中 / 未开始 / 已逾期 / 完成率。
- **可选导出 Word**：把看板结论 + 图表截图 + 数据表整理成 `.docx`。

## 目录结构

```
project-dashboard/
├── SKILL.md                        # 触发条件 + 流程 + 全部硬性约定
├── README.md                       # 本文件
├── references/
│   └── schema.md                   # 台账列结构说明（14 列 + 工作包类别枚举）
├── assets/
│   ├── dashboard_template.html     # 看板 HTML 模板（占位符）
│   └── echarts.min.js              # 内嵌图表库（离线可用）
└── scripts/
    └── build_dashboard.py          # 一键生成脚本
```

## 安装

```bash
# 复制到 WorkBuddy 的 skills 目录
cp -r project-dashboard ~/.codebuddy/skills/
```

## 使用

**方式一：对话触发**

上传台账 Excel 后，直接说「做项目看板 / 生成推进看板 / 甘特图 / 工作包统计」，skill 会自动触发。

**方式二：命令行**

```bash
python3 scripts/build_dashboard.py <台账.xlsx> \
  --category 新系统建设 \
  --today 2026-11-14 \
  --outdir ./output \
  --word
```

| 参数 | 说明 | 默认 |
| --- | --- | --- |
| `--category` | 要筛选的项目类别 | `新系统建设` |
| `--today` | 截止日（甘特图今日线、逾期判定） | 台账「最近更新」最大日期 |
| `--outdir` | 输出目录 | 脚本同目录 |
| `--word` | 额外导出 Word 版 | 不导出 |

## 硬性约定

| 项 | 约定 |
| --- | --- |
| 甘特图 | 只列里程碑=是、按计划完成日升序、用工作包名称标示 |
| 直方图 | 柱内绿/灰分色 + 百分比、柱顶标总数 |
| 饼图 | 每类标数量 + 占比 |
| 颜色 | 完成绿 / 进行中橙 / 未开始灰 / 逾期红 |
| 逾期判定 | 计划完成日 < 今日 且 状态 ≠ 已完成 |
| 明细表 | 不生成 |
| 交付物 | 单文件 HTML + 可选 Word |

## 依赖

- Python 3.8+
- `openpyxl`（读取 Excel）
- 可选：`playwright` + Chromium（`--word` 导出时截图图表）
- 可选：WorkBuddy 本地 editor_sdk（`--word` 导出时写入 Word）
