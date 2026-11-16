---
name: attendance-stats-calc
description: "三早一晚考勤与免勤统计全流程自动化工具：提供学生考勤计算、免勤分项导出、词典规则安全维护与防篡改核对。触发: 三早一晚/免勤统计/出勤率计算/草莓小熊"
version: 1.1.0
author: Antigravity DeepMind Team
tags:
  - attendance
  - exemption
  - excel
  - sanzao-wanwai
  - strawberry-bear
  - pyinstaller
default_language: zh
languages:
  - en
  - zh
  en_doc: SKILL.en.md
  zh_doc: SKILL.zh.md
---

> **语言切换 (Language Routing)**: 本文为**中文完整规范** (`SKILL.zh.md`)。英文规范请参阅 [`SKILL.md`](SKILL.md) 与 [`SKILL.en.md`](SKILL.en.md)。

| 元数据 | 配置值 |
| :--- | :--- |
| **Skill 标识符** | `attendance-stats-calc` |
| **中文触发词** | `三早一晚`、`免勤统计`、`出勤率计算`、`草莓小熊` |
| **适用场景** | 高校本科生早晚自习考勤核算、免勤分项导出、词典规则维护与桌面端复核 |
| **所属体系** | Antigravity 全局能力库 (`.agents/skills/`) |

---

# attendance-stats-calc — 三早一晚免勤统计全流程自动化工具规范

本 Skill 面向高校本科生早晚自习考勤与免勤统计业务，深度融合自动化流水线计算、严格上下文脱敏保护（Zero PII）、词典安全生命周期治理与桌面端草莓小熊 GUI 人工复核机制。

---

## 一、三大作业轨道体系

| 轨道编号 | 轨道名称 | 适用场景 | 对应入口与工具 | 详细指引文档 |
| :--- | :--- | :--- | :--- | :--- |
| **轨道 1** | **Agent 自动化流水线** | 考勤明细计算、汇总终稿生成、历史批次查询、健康预检与防篡改校验 | `python scripts/run_calc.py` 或 `cli.py` | [`references/cli_reference.md`](references/cli_reference.md) |
| **轨道 2** | **词典安全生命周期治理** | 语义合法性前置自检、白名单核对、时间戳自动备份与别名受控写入 | `python scripts/dict_tool.py` | [`references/business_rules.md`](references/business_rules.md) |
| **轨道 3** | **桌面端 GUI 人工复核** | 重名身份冲突、新专业别名确认、黄色高亮免勤分摊抽检 | 桌面端 `三早一晚免勤自动统计.exe` | [`references/gui_workflow_guide.md`](references/gui_workflow_guide.md) |

---

## 二、关键质量闸门与白盒业务口径

1. **零隐私泄漏（Zero PII）硬约束**:
   - `scripts/run_calc.py` 与 `cli.py` 严格拦截所有输出，彻底清洗 13 位数字学号（`\d{13}`）与学生真实姓名明细。
   - 向 Agent 对话上下文仅回传学院级汇总出勤率、每周出勤率、异常分类计数、阶段耗时与输出文件路径。
   - 内部集成 13 位学号（SID）正则表达式强制断言：`re.compile(r'(?<![0-9A-Fa-f])(?!99)\d{13}(?![0-9A-Fa-f])')`，杜绝任何数据泄露。
2. **严格模式 vs 宽容模式 (`core/calc.py:291-440`)**:
   - **默认严格契约模式（`'0'` / `complete_mode=False`）**：分母 $\le 0$ 或任意排查机房实到缺失时，出勤率输出为 `rate = None`（留空 `—`）。
   - **宽容容错模式（`'1'` / `complete_mode=True`）**：分母 $\le 0$ 时按 100.00%（满勤）兜底计入，并折算有效机房。
3. **每周综合出勤率 vs 独立汇总终稿 (`core/calc.py` vs `core/pipeline.py`)**:
   - 明细表中每周出勤率是**各有效排查日出勤率的算术平均值**（任一应统计日缺失时周平均留空）。
   - 独立汇总终稿（`*_汇总终稿.xlsx`）采用**总和除总和**（$\sum \text{Actual} / \sum (\text{Expected} - \text{Exempt})$），且**恒走严格模式**；遇实到缺失、有效应到非正、免勤人数/适用范围待确认时留空并标为「待复核」。
4. **v1.1.0 身份争议口径与模板表头保留**:
   - **同号异名 / 跨学院去重计 1 人**：按 `(学号, 星期, 自习)` 去重计 1 人，**不再暂缓计入**；日出勤率与汇总终稿照常计算，来源表及行号写入「免勤核对」页、终稿「执行效果」列标注核对来源、桌面端主窗口 Header 下方固定提示条展示争议告警。
   - **仍暂不计入的情形**：无学号且无法按姓名+班级唯一对齐（`学号无法核对`）、空白或模糊校区待核对（`校区待核对`）、周次/星期/时段无法识别、集体申请缺班级（`集体申请缺班级`）及占位学号无人数（`免勤不完整（待补人数）`）。
   - **模板表头与合并区保留**：导出保留原表的星期与 `平均` / `周平均` / 预留空白子表头及首学院 `N4:N8` 合并区。
5. **机房免勤分摊与黄色高亮标示 (`ESTIMATE_RGB = 'FFFFF3CD'`)**:
   - `#FFFFF3CD`（斜体浅黄底）标记的是**按班级推断或按比例分摊的机房免勤人次（包含 0 人次）**，绝非“排课推断缺失机房”。
   - `--match-classes` 默认为 `False`。
6. **最大余额法分摊与容量上限硬约束**:
   - 未指定抽查教室的免勤通过最大余额法（Hare-Niemeyer `allocate`）分摊到各教室，单一教室免勤不得超出容量上限，严禁跨教室二次再分配。
7. **免勤理由严格白名单**:
   - 词典「免勤理由分类」严格仅允许四类：`('课务学业', '社团竞赛训练', '健康医疗', '其他')`。
8. **集体免勤覆盖个人免勤**:
   - 同班级、同日、同时段集体免勤申报自动覆盖该班学生的个人请假，不重复扣减。
9. **工作表结构不变量与环境要求**:
   - 明细结果表仅保留三张原始可见工作表：`雁山（早）`、`雁山（晚）` 与 `屏风`（除非显式 `--include-exempt-summary` 追加分项页）。其中 `屏风` 表执行逐单元格零修改。
   - **云端仓库与 Python 环境**：源码与 Skill 托管于 `https://github.com/LL-LK/jisuanqi`（Releases `latest` 含 `.sha256` 校验包）。调用时须选用已安装 `openpyxl`（及 GUI 所需的 `PySide6`）的 Python 3.10+ 环境，注意规避 LibreOffice / WPS 自带精简 Python 抢占 `PATH`；跨目录调用 `scripts/run_calc.py` 时可通过 `--app-root` 或环境变量 `ATTENDANCE_APP_ROOT` 指定仓库根目录。

---

## 三、快速上手与常用命令

### 3.1 工作簿健康预检与表型识别
```bash
python scripts/run_calc.py --preflight path/to/detail.xlsx path/to/exempt.xlsx
```

### 3.2 自动化考勤全流程计算
```bash
python scripts/run_calc.py \
  --detail path/to/detail.xlsx \
  --exempt path/to/exempt1.xlsx path/to/exempt2.xlsx \
  --week 7 \
  --complete-mode auto \
  --summary-out path/to/summary.xlsx \
  --out path/to/result_detail.xlsx
```

### 3.3 历史计算批次查询
```bash
python scripts/run_calc.py --list-runs
python scripts/run_calc.py --show-run 5536
```

### 3.4 防篡改与原表实到核对
```bash
python scripts/run_calc.py \
  --verify-tamper path/to/result_detail.xlsx \
  --against-detail path/to/detail.xlsx
```

### 3.5 词典安全维护与别名新增
```bash
# 1. 词典结构与分类白名单校验
python scripts/dict_tool.py check --dict path/to/免勤词典.xlsx

# 2. 受控新增别名（先行语义校验 -> 校验通过再建备份 -> 原子落盘）
python scripts/dict_tool.py add-alias \
  --sheet 学院别名 \
  --target "化学与生物工程学院" \
  --alias "化生院" \
  --note "口语简称"
```

---

## 四、双模式运行指南：Agent 自动化 vs 桌面端 GUI

### 4.1 自动化模式（推荐优先使用）
在常规每周考勤数据完整、无身份歧义时，由 Agent 直接调用 `scripts/run_calc.py` 自动化批量计算，毫秒级快速产出完整计算版表格。

### 4.2 桌面端 GUI 人工复核模式（降级兜底）
当计算结果返回以下异常时，Agent 应主动向用户输出 GUI 操作指引：
1. **重名身份冲突**：免勤表中学生未填学号，且全校存在同名同姓不同班级/学院学生；
2. **未知专业别名**：免勤表中出现全新缩写或错别字；
3. **回填差额人工对账**：免勤申报总人数超出排课计划容量时，需人工核对具体请假凭据。

**GUI 操作界面指引**：
- 双击启动 `三早一晚免勤自动统计.exe`（无控制台静默运行）；
- 顶部 4 卡槽流水线：`01 明细表`、`02 免勤表`、`开始计算`、`03 结果工作簿`；
- Header 固定提示条（v1.1.0）：存在同号异名/跨学院身份争议时，主窗口顶部固定条告警并指引核查「免勤核对」页，在免勤源表副本中更正后重算；
- 4 个主标签页：`概览`、`学院免勤统计`、`出勤率`（含 `每日` 与 `每周`）、`词典编辑`；
- 辅助功能：点击 `更多工具…` 可进入「教室明细」、「异常清单」、「低出勤率复核」；
- 词典操作：修改规则后点击 `保存词典` 即可自动执行白名单校验并热重载。

---

## 五、参考文档索引

- 英文主规范：[`SKILL.md`](SKILL.md) 与 [`SKILL.en.md`](SKILL.en.md)
- 考勤口径与业务规则：[`references/business_rules.md`](references/business_rules.md)
- 命令行参数与 JSON 规范：[`references/cli_reference.md`](references/cli_reference.md)
- 桌面端 GUI 操作复核指南：[`references/gui_workflow_guide.md`](references/gui_workflow_guide.md)
