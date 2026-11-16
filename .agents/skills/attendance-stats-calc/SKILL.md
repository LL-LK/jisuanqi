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
default_language: en
languages:
  - en
  - zh
en_doc: SKILL.en.md
zh_doc: SKILL.zh.md
---

> **Language Routing (Default: English)**: By default, consult this **English** specification (`SKILL.md` / [`SKILL.en.md`](SKILL.en.md)). For the Chinese edition, see [`SKILL.zh.md`](SKILL.zh.md).

| Metadata | Value |
| :--- | :--- |
| **Skill ID** | `attendance-stats-calc` |
| **Default Language** | English (`SKILL.md` / [`SKILL.en.md`](SKILL.en.md)) |
| **Chinese Edition** | [`SKILL.zh.md`](SKILL.zh.md) |
| **Triggers (Chinese)** | `三早一晚`, `免勤统计`, `出勤率计算`, `草莓小熊` |
| **Primary Domain** | Campus Attendance & Exemption Analytics Automation |

---

# attendance-stats-calc — English Specification

The **`attendance-stats-calc`** skill governs the end-to-end processing pipeline for undergraduate morning and evening study attendance ("San Zao Yi Wan" / 三早一晚). It integrates automated calculation via CLI, strict privacy preservation (zero PII leakage in Agent context), safe dictionary lifecycle management, and visual desktop GUI manual fallback.

---

## 1. Operating Tracks

| Track | Target Scenario | Tool / Command | Documentation |
| :--- | :--- | :--- | :--- |
| **Track 1: Automated Agent Pipeline** | Automated weekly calculation, preflight diagnosis, runs history query, and tamper verification | `python scripts/run_calc.py` or `python cli.py` | [`references/cli_reference.md`](references/cli_reference.md) |
| **Track 2: Dictionary Maintenance** | Table inspection, semantic validation, whitelist check, and timestamped backup | `python scripts/dict_tool.py` | [`references/business_rules.md`](references/business_rules.md) |
| **Track 3: Desktop GUI Review** | Interactive resolution for student identity conflicts, unmapped classes, or visual inspection | Desktop GUI (`三早一晚免勤自动统计.exe`) | [`references/gui_workflow_guide.md`](references/gui_workflow_guide.md) |

---

## 2. Core Quality & Privacy Gates

1. **Zero Personal Identification (PII) Leakage**:
   - The automation wrapper (`scripts/run_calc.py`) and CLI strictly sanitize all student names and 13-digit student IDs (`\d{13}`) from output emitted to Agent context.
   - Only college-level aggregations, weekly rates, anomaly counters, status codes, and artifact paths are returned.
   - Built-in internal privacy assertion validates zero unmasked student ID patterns (`(?<![0-9A-Fa-f])(?!99)\d{13}(?![0-9A-Fa-f])`) prior to printing.
2. **Strict vs. Tolerant Calculation Modes (`core/calc.py:291-440`)**:
   - In default **Strict Mode** (`complete_mode=False` / `'0'`), if denominator $\le 0$ or any classroom actual is missing, `rate = None` (leave blank `—`).
   - $\text{Denominator} \le 0 \to 100.00\%$ and effective classroom conversion applies ONLY in **Tolerant Mode** (`complete_mode=True` / `'1'`).
3. **Weekly Average vs. Standalone Summary Workbook (`core/calc.py` vs `core/pipeline.py`)**:
   - Weekly rate in detail sheets is the **arithmetic mean of daily rates** (left blank if any scheduled day is missing).
   - $\sum \text{Actual} / \sum (\text{Expected} - \text{Exempt})$ is strictly used for the standalone summary workbook (`*_汇总终稿.xlsx`), which always runs in strict audit mode.
4. **v1.1.0 Identity Dispute & Template Header Preservation**:
   - **Identity Conflict Deduplication**: Records sharing the same 13-digit SID across colleges or conflicting names in the same session deduplicate by `(SID, weekday, session)` and **count as 1 person** (no longer held from counting); daily & summary rates compute normally, while source rows are logged in `免勤核对` and flagged in summary `执行效果` + GUI fixed header alert banner. (Unresolved missing-SID rows `学号无法核对`, pending campus `校区待核对`, missing-class group rows, and placeholder SIDs without counts remain excluded/held.)
   - **Template Header Preservation**: Exports preserve original weekday and `平均` / `周平均` headers (including reserved blank subheader rows) and first-college merged region (`N4:N8`).
5. **Classroom Exemption Allocation & Amber Highlight (`#FFFFF3CD`)**:
   - `#FFFFF3CD` (italic light yellow) marks classroom exemption headcounts inferred by class or proportionally allocated via the Hare-Niemeyer largest remainder method (including 0 headcount, with classroom capacity upper bound constraint).
   - `--match-classes` defaults to `False`.
6. **Strict Reason Category Whitelist**:
   - The exemption reason classification sheet (`免勤理由分类`) strictly admits only four authorized category values:
     `('课务学业', '社团竞赛训练', '健康医疗', '其他')`.

---

## 2.5 Cloud Distribution & Environment Setup (`LL-LK/jisuanqi`)
- **Repository**: `https://github.com/LL-LK/jisuanqi` (source + `.agents/skills/attendance-stats-calc/`); prebuilt Windows/macOS/Linux packages and `.sha256` checksums are published under Releases tag `latest`.
- **Python Environment Note**: Always invoke a Python 3.10+ interpreter with `openpyxl` (and `PySide6` for GUI) installed; avoid stripped Python binaries bundled with office suites (such as LibreOffice/WPS) that may appear earlier on `PATH`. Set `ATTENDANCE_APP_ROOT` or pass `--app-root` if invoking `scripts/run_calc.py` outside the cloned repo.

---

## 3. Quick Start & Common Commands

### 3.1 Preflight Health Diagnosis
```bash
python scripts/run_calc.py --preflight path/to/detail.xlsx path/to/exempt.xlsx
```

### 3.2 Automated Full Calculation
```bash
python scripts/run_calc.py \
  --detail path/to/detail.xlsx \
  --exempt path/to/exempt.xlsx \
  --week 7 \
  --complete-mode auto \
  --summary-out path/to/summary.xlsx \
  --out path/to/result.xlsx
```

### 3.3 Query Historical Runs
```bash
python scripts/run_calc.py --list-runs
python scripts/run_calc.py --show-run 5536
```

### 3.4 Dictionary Semantic Validation & Alias Addition
```bash
# Validate dictionary structure and reason whitelist
python scripts/dict_tool.py check --dict path/to/免勤词典.xlsx

# Safely add college alias (pre-validates semantics before creating backup!)
python scripts/dict_tool.py add-alias \
  --sheet 学院别名 \
  --target "化学与生物工程学院" \
  --alias "化生院" \
  --note "口语简称"
```

---

## 4. Documentation Index

- Detailed English Documentation: [`SKILL.en.md`](SKILL.en.md)
- Complete Chinese Specification: [`SKILL.zh.md`](SKILL.zh.md)
- Attendance Calibers & Business Rules: [`references/business_rules.md`](references/business_rules.md)
- CLI Arguments & JSON Schemas: [`references/cli_reference.md`](references/cli_reference.md)
- Desktop GUI Operation Manual: [`references/gui_workflow_guide.md`](references/gui_workflow_guide.md)
