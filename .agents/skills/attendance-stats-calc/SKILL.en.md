# Attendance Statistics Automation Skill — Comprehensive English Specification (`attendance-stats-calc`)

## 1. Overview & Institutional Context

The **`attendance-stats-calc`** skill provides full-lifecycle automation for undergraduate morning and evening study attendance ("San Zao Yi Wan" / 三早一晚) across universities. It bridges programmatic batch processing and human administrative oversight through three cohesive operational tracks:
1. **Automated Calculation & Pipeline Execution**: High-throughput calculation, preflight diagnosis, and tamper-resistant exporting via CLI.
2. **Zero-PII Privacy Protection**: Complete redaction of student personal information from Agent contexts while retaining comprehensive aggregate metrics.
3. **Safe Dictionary Lifecycle Management**: Semantic pre-validation, automated timestamped backups, and strict categorical whitelist validation.
4. **Desktop GUI Manual Review Fallback**: Clear operational guidance when data ambiguities or identity conflicts require human intervention.

---

## 2. System Architecture & Component Roles

```
attendance-stats-calc/
├── SKILL.md                          # Primary specification router
├── SKILL.en.md                       # Comprehensive English documentation
├── SKILL.zh.md                       # Comprehensive Chinese documentation
├── scripts/
│   ├── __init__.py                   # Package initializer
│   ├── run_calc.py                   # Privacy-enforcing runner wrapper
│   └── dict_tool.py                  # Dictionary inspection, backup, and safe update tool
└── references/
    ├── business_rules.md             # Calibers, formulas, no-cap invariant, and FFFFF3CD styling
    ├── cli_reference.md              # Complete CLI flags, exit codes, and JSON schemas
    └── gui_workflow_guide.md         # Desktop GUI manual review & fallback guide
```

### 2.1 Component Interaction Diagram

```
[Agent Context]
       │
       ▼
[scripts/run_calc.py] ◄─── Zero-PII Enforcement (Scrub IDs/Names, Assert Privacy)
       │
       ├─► [core/excel_safety.py]  (Preflight health checks, 4-way table classification)
       ├─► [core/parsing.py]       (Classroom normalization, week number inference)
       ├─► [core/pipeline.py]      (Calculation execution & phase profiling)
       ├─► [core/export.py]        (Atomic workbook export, HMAC anti-tamper signing)
       ├─► [core/storage.py]       (Run history persistence and query)
       └─► [core/exempt_report.py] (Standalone exemption breakdown export)
```

---

## 3. Core Business Logic & Calibers

### 3.1 Strict Mode vs. Tolerant Mode (`core/calc.py:291-440`)
- **Strict Mode (`complete_mode=False` / `'0'`) [Default]**:
  $$\text{Denominator} = \text{Expected} - \text{Exempt}$$
  If $\text{Denominator} \le 0$ or any scheduled classroom headcount is missing:
  $$\text{Rate} = \text{None (leave blank '—')}$$
  This prevents distorted statistics from masking administrative issues.
- **Tolerant Mode (`complete_mode=True` / `'1'`)**:
  If $\text{Denominator} \le 0$:
  $$\text{Rate} = 1.0 \quad (100.00\%)$$
  with note: `Expected X - Exempt Y <= 0, recorded as 100.00% full attendance`. Missing classrooms are calculated based on valid classrooms.

### 3.2 Weekly Average vs. Standalone Summary Final Workbook
- **Detail Sheets (`*_第N周计算结果.xlsx`)**: Weekly rate is the **arithmetic mean of daily rates** for scheduled days (left blank if any scheduled day lacks a computed rate).
- **Standalone Summary Workbook (`*_汇总终稿.xlsx`)**: Weekly rate is evaluated as $\sum \text{Actual} / \sum (\text{Expected} - \text{Exempt})$. **Always runs in strict mode**. Missing actuals, non-positive effective denominators, incomplete placeholder counts, or rule-review scopes flag the college as "待复核" (Pending Review).

### 3.3 v1.1.0 Identity Dispute & Template Preservation Calibers
- **Same-SID Cross-College / Name Conflict**: Records sharing the same SID across colleges or different names in the same study session are deduplicated by `(SID, weekday, session)` and **counted as 1 person** (no longer held from calculation). Attendance rates and the standalone summary calculate normally, while `免勤核对` lists source workbooks/rows, summary `执行效果` notes the review reference, and the GUI shows a fixed alert banner below the Header.
- **Records Still Excluded / Held**: Unresolved missing-SID records (`学号无法核对`), ambiguous/pending campus (`校区待核对`), unrecognized week/weekday/session, group applications missing class (`集体申请缺班级`), and placeholder SIDs without headcounts (`免勤不完整（待补人数）`).
- **Template Header Preservation**: Exports preserve original weekday and `平均` / `周平均` headers (including reserved blank subheaders) as well as the first-college merged range (`N4:N8`).

### 3.4 Classroom Exemption Allocation & Amber Highlight (`#FFFFF3CD`)
- `#FFFFF3CD` (italic light yellow) marks classroom exemption headcounts inferred by class or proportionally distributed via Hare-Niemeyer largest remainder method (including 0 headcount).
- **Capacity Constraint**: Allocation per room is strictly capped at room capacity; cross-room redistribution is prohibited.
- `--match-classes` defaults to `False`.

### 3.5 Key Invariants & Cloud Setup
- **Zero Modification on `屏风` Sheet**: Pingfeng campus sheet is strictly preserved without modifying any cells.
- **Collective Exemption Priority**: Class-wide collective exemption covers individual leaves of the same class, date, and study slot.
- **13-Digit SID Zero-Leakage**: All outputs enforce regex assertion: `(?<![0-9A-Fa-f])(?!99)\d{13}(?![0-9A-Fa-f])`.
- **Cloud Repository & Python PATH**: Source and `.agents/skills/attendance-stats-calc/` are hosted at `https://github.com/LL-LK/jisuanqi` (Releases tag `latest` includes `.sha256` packages). Always select a Python environment containing `openpyxl`/`PySide6`, avoiding stripped office-suite Python binaries on `PATH`.

---

## 4. Operational Track 1: Automated Agent Pipeline (`scripts/run_calc.py`)

The runner script wraps the core engine to guarantee that **zero student identification data is returned to the Agent's conversational context**.

### 4.1 Invocation Examples

#### A. Full Weekly Attendance Calculation
```bash
python scripts/run_calc.py \
  --detail path/to/detail_sheet.xlsx \
  --exempt path/to/exempt_sheet1.xlsx path/to/exempt_sheet2.xlsx \
  --week 7 \
  --complete-mode auto \
  --summary-out path/to/summary_final.xlsx \
  --out path/to/calculated_result.xlsx
```

#### B. Preflight File Health Diagnosis & Table Classification
```bash
python scripts/run_calc.py --preflight path/to/detail.xlsx path/to/exempt.xlsx
```

#### B2. Pre-Calc Formula Flattening & Zero-Formula Output Invariant
All exported Excel files are guaranteed to contain **zero formula cells** (`data_type != 'f'` across all sheets including `屏风` and protected columns). Headcount cells (`免勤`, `实到`, `应到`) remain pure integers (e.g. `39`, `105` formatted as `'0'`), while percentage rate cells (`出勤率`, `周平均`, `巡查率`) remain floating-point ratios formatted as **`'0.00%'`** (2 decimal places). To strip formulas from an input workbook prior to calculation:
```bash
python scripts/run_calc.py --flatten-formulas path/to/detail.xlsx --out path/to/detail_pure.xlsx
# Or atomically overwrite in place:
python scripts/run_calc.py --flatten-formulas path/to/detail.xlsx --in-place
```

#### B3. Pre-Calc Privacy Stripping & Zero-PII Model Isolation
All model-facing CLI/Skill/diagnostic outputs strictly expose ONLY keyed pseudonymous SIDs (`ANON-XXXXXXXXXX`), college exemption counts, and 4-category reason breakdown counts (`college_exempt_summary`). Real student names (including surname-masked forms like `张*`), raw SIDs, raw reasons, personal classes, phones, and ID numbers are blocked by `PrivacyViolationError`. Before feeding workbooks into automated workflows, run:
```bash
python scripts/run_calc.py --strip-privacy path/to/exempt.xlsx --out path/to/exempt_sanitized.xlsx
```

#### C. Historical Runs Query
```bash
python scripts/run_calc.py --list-runs
python scripts/run_calc.py --show-run 5536
```

#### D. Standalone Exemption Report Export
```bash
python scripts/run_calc.py \
  --exempt-report path/to/exempt_report.xlsx \
  --exempt path/to/exempt.xlsx \
  --week 7
```

#### E. Anti-Tamper & Actual Headcount Verification
```bash
python scripts/run_calc.py \
  --verify-tamper path/to/result.xlsx \
  --against-detail path/to/detail.xlsx
```

---

## 5. Operational Track 2: Dictionary Maintenance (`scripts/dict_tool.py`)

### 5.1 Semantic Pre-Validation Before Backup
`dict_tool.py:cmd_add_alias` tests dictionary validity via `Dictionary(tables)` and enforces the 4-category whitelist (`'课务学业'`, `'社团竞赛训练'`, `'健康医疗'`, `'其他'`) BEFORE generating any backup file. If validation fails, it aborts immediately with exit code 1 without creating orphan backups.

```bash
# Add alias safely
python scripts/dict_tool.py add-alias \
  --sheet 学院别名 \
  --target "化学与生物工程学院" \
  --alias "化生院" \
  --note "口语简称"
```

---

## 6. Operational Track 3: Desktop GUI Review (`三早一晚免勤自动统计.exe`)

When identity conflicts, unmapped colleges, or capacity overflows occur, the desktop GUI provides an intuitive visual workflow:
- **Top 4-Slot Pipeline**: `01 明细表`, `02 免勤表`, `开始计算`, `03 结果工作簿`;
- **4 Main Tabs**: `概览`, `学院免勤统计`, `出勤率` (`每日` / `每周`), `词典编辑`;
- **Tools**: Click `更多工具…` to access `教室明细`, `异常清单`, `低出勤率复核`;
- **Dictionary**: Click `保存词典` to validate and hot-reload updated rules.
