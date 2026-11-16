# Attendance Statistics System CLI Reference Manual
# 命令行接口与参数规范速查手册

## 1. 概述与标准化退出码 (Overview & Exit Codes)

三早一晚免勤自动统计系统提供底层命令行核心（`cli.py`）与面向 Agent 的隐私安全封装运行器（`scripts/run_calc.py`）。

### 标准化退出状态码 (Standardized Exit Codes)

| 退出码 | 状态标识 | 业务与技术含义 |
| :---: | :--- | :--- |
| **`0`** | **`SUCCESS / OK`** | 考勤计算成功、前置体检通过、防篡改签名验证一致或历史批次查询成功。 |
| **`1`** | **`RUNTIME / INPUT ERROR`** | 命令行缺少必选参数、文件不存在/被独占锁定/格式损坏、或计算流水线抛出异常。 |
| **`2`** | **`TAMPER VIOLATION`** | 防篡改校验失败：表格数字签名篡改、元数据不符、或实到人数与原始明细逐格对拍存在差异。 |

---

## 2. 完整命令行参数清单 (Command-Line Arguments)

```bash
python cli.py [OPTIONS]
# 或通过封装运行器：
python scripts/run_calc.py [OPTIONS]
```

| 参数名称 | 类型 / 取值范围 | 默认值 | 详细功能说明 |
| :--- | :--- | :--- | :--- |
| **`--json`** | Flag | `False` | 启用机器可读的标准 JSON 格式输出，内部强制执行 13 位学号零泄露断言。 |
| **`--preflight FILE [FILE ...]`** | List[str] | `None` | 执行毫秒级前置 Excel 健康自检与表型自动识别（`DETAIL` / `EXEMPT` / `SUMMARY` / `DICTIONARY`）。 |
| **`--detail PATH`** | str | `None` | 查课考勤执行情况明细表路径（.xlsx）。 |
| **`--exempt PATH [PATH ...]`** | List[str] | `[]` | 一个或多个免勤审批大表路径（支持重复声明或传入多个文件）。 |
| **`--out PATH`** | str | 自动派生 | 考勤明细计算结果工作簿输出路径（默认生成至 `考勤统计输出/*_第N周计算结果.xlsx`）。 |
| **`--summary-out PATH`** | str | 自动派生 | 独立汇总终稿工作簿输出路径（默认生成为 `*_汇总终稿.xlsx`）。 |
| **`--exempt-report PATH`** | str | `None` | 另行导出独立免勤分项对账报表工作簿（.xlsx）。支持在无 `--detail` 时独立运行。 |
| **`--week N`** | int (1–30) | 自动识别或 7 | 统计周次。省略时优先从明细表文件名或 A1 标题解析。 |
| **`--complete-mode {auto,strict,tolerant}`** | Choice | `auto` | 出勤率完整计算模式：`strict`（严格契约，分母 $\le 0$ 留空）、`tolerant`（容错完整，分母 $\le 0$ 满勤 100%）、`auto`（按词典参数）。 |
| **`--campus NAME`** | str | `雁山` | 计入统计的校区名称。 |
| **`--source {全免勤大表,在线存档,未分类}`** | Choice | `未分类` | 免勤表来源分类先验。 |
| **`--dictionary PATH`** | str | 自动探测 | 自定义免勤词典文件路径（.xlsx）。省略时按目录向上爬梳就近解析。 |
| **`--match-classes`** | Flag | `False` | 课表匹配开关。默认关闭（未带教室号按已验收的比例分摊口径处理）。 |
| **`--college-exempt`** | Flag | `False` | 按学院回填免勤（在明细表缺少班级合并单元格时按学院汇总回填）。 |
| **`--include-exempt-summary`** | Flag | `False` | 将免勤统计摘要一并写入明细结果工作簿中。 |
| **`--no-storage`** | Flag | `False` | 禁用 SQLite 历史计算批次记录入库。 |
| **`--list-runs`** | Flag | `False` | 查询 SQLite 本地数据库中最近 30 条历史计算记录，返回格式化文本或 JSON 后退出。 |
| **`--show-run RUN_ID`** | int | `None` | 查询指定批次 ID 的历史计算详情，返回格式化文本或 JSON 后退出。 |
| **`--diagnostic-log PATH`** | str | `None` | 生成完全脱敏的工业级系统诊断与故障排查 Markdown 报告（.md）。 |
| **`--verify-tamper PATH`** | str | `None` | 校验指定结果 Excel 文件的防篡改元数据与 HMAC 签名完整性。 |
| **`--against-detail PATH`** | str | `None` | 配合 `--verify-tamper` 使用，与可信原始明细表逐格核对实到人数。 |

---

## 3. 标准化 JSON 响应结构 (`--json`)

所有 `--json` 输出均经过 `DiagnosticMasker` 脱敏清洗，并在输出前强制执行 13 位学号（SID）正则表达式隐私断言：
`re.compile(r'(?<![0-9A-Fa-f])(?!99)\d{13}(?![0-9A-Fa-f])')`。

### 3.1 完整计算输出 (`action: "run"`)

```json
{
  "action": "run",
  "status": "ok",
  "code": 0,
  "duration_ms": 1180,
  "phases": {
    "preflight_ms": 12,
    "parse_ms": 195,
    "calc_ms": 48,
    "export_ms": 925
  },
  "output_file": "<OUTPUT_DIR>\\detail_第7周计算结果.xlsx",
  "summary_output_file": "<OUTPUT_DIR>\\detail_第7周计算结果_汇总终稿.xlsx",
  "exempt_report": null,
  "summary": {
    "工作表": "雁山（早） / 雁山（晚） / 屏风",
    "教室数": 28,
    "学院数": 12,
    "统计周次": 7,
    "可计算出勤率的学院×天组合": "24 / 24",
    "异常条目": 0
  },
  "exempt_stats": {
    "总记录数": 450,
    "有效记录数": 442,
    "因周次不符过滤": 8,
    "因校区过滤": 0
  },
  "college_rates": [
    {
      "sheet": "雁山（早）",
      "college": "化学与生物工程学院",
      "day": "周一",
      "expected": 2100,
      "actual": 2010,
      "exempt": 87,
      "rate": 0.9985,
      "status": "正常"
    }
  ],
  "weekly_rates": [
    {
      "sheet": "雁山（早）",
      "college": "化学与生物工程学院",
      "rate": 0.9985,
      "calculated_days": 4,
      "scheduled_days": 4
    }
  ],
  "discrepancies_count": 0,
  "discrepancies": [],
  "anomalies_count": 0,
  "anomalies": []
}
```

### 3.2 前置健康体检输出 (`action: "preflight"`)
在体检输出中，原始路径 `raw_path` 已被彻底剔除，仅暴露脱敏路径 `path`：
```json
{
  "action": "preflight",
  "status": "ok",
  "code": 0,
  "duration_ms": 15,
  "reports": [
    {
      "path": "<WORKSPACE>\\ruanjian\\免勤统计系统\\detail.xlsx",
      "is_valid": true,
      "error_code": "OK",
      "user_message": "文件健康可用。",
      "tech_detail": "",
      "table_type": "DETAIL",
      "file_size_bytes": 35840,
      "file_size_formatted": "35.0 KB",
      "sha256_prefix": "7eace161",
      "sheet_names": ["雁山（早）", "雁山（晚）", "屏风"]
    }
  ]
}
```

### 3.3 历史计算批次查询 (`action: "list_runs"`)
```json
{
  "action": "list_runs",
  "status": "ok",
  "code": 0,
  "duration_ms": 12,
  "total_runs": 1,
  "runs": [
    {
      "id": 5536,
      "run_uuid": "db58c3e8-4407-4846-b483-dac161624dbd",
      "created_at": "2026-11-11T04:52:37.597943",
      "week_no": 14,
      "campus": "雁山",
      "status": "WARNING",
      "detail_file_name": "detail.xlsx",
      "detail_file_size": 7945,
      "dictionary_path": "%USERPROFILE%\\AppData\\Local\\Temp\\dict.xlsx",
      "output_path": "%USERPROFILE%\\Desktop\\out.xlsx",
      "overall_rate": 0.7657,
      "duration_ms": 170
    }
  ]
}
```
