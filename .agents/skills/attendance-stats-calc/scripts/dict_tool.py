# -*- coding: utf-8 -*-
"""Safe dictionary maintenance tool for attendance-stats-calc.

Supports:
- view: Display rule tables in clean text format.
- check: Validate dictionary tables and verify reason categories against strict whitelist.
- backup: Generate timestamped backup (免勤词典.bak.<timestamp>.xlsx).
- add-alias: Automatically backup, display differential preview, validate whitelist, and safely update tables.
"""

from __future__ import annotations

import argparse
import datetime
import os
import shutil
import sys
from typing import Dict, List, Optional, Sequence, Tuple

# Strict whitelist for exemption reason categories
ALLOWED_REASON_CATEGORIES: Tuple[str, ...] = ("课务学业", "社团竞赛训练", "健康医疗", "其他")

SHEET_HEADERS: Dict[str, Tuple[str, ...]] = {
    "学院别名": ("正式名称", "别名", "说明"),
    "班级别名": ("别名", "规范名称", "说明"),
    "校区别名": ("原始写法", "规范校区", "说明"),
    "免勤理由分类": ("理由关键词", "分类", "说明"),
    "统计参数": ("参数", "值", "说明"),
}


def discover_app_root(custom_path: Optional[str] = None) -> Optional[str]:
    """Dynamically discover the root directory containing core.dictionary."""
    candidates = []
    if custom_path:
        candidates.append(os.path.abspath(custom_path))

    env_root = os.environ.get("ATTENDANCE_APP_ROOT")
    if env_root:
        candidates.append(os.path.abspath(env_root))

    cwd = os.getcwd()
    candidates.append(cwd)

    # Check if skill resides inside <repo>/.agents/skills/attendance-stats-calc/scripts/
    script_repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
    candidates.append(script_repo)

    for depth in range(1, 4):
        up = os.path.abspath(os.path.join(cwd, *(['..'] * depth)))
        candidates.append(up)
        candidates.append(os.path.join(up, "免勤统计系统"))
        candidates.append(os.path.join(up, "ruanjian", "免勤统计系统"))

    candidates.append(os.path.join(cwd, "免勤统计系统"))
    candidates.append(os.path.join(cwd, "ruanjian", "免勤统计系统"))

    # Fallback to standard installation paths if not found via cwd climbing
    for d_letter in ("D", "d"):
        d_prefix = d_letter + ":"
        candidates.append(os.path.join(d_prefix, os.sep, "Software", "ruanjian", "免勤统计系统"))
        candidates.append(os.path.join(d_prefix, os.sep, "Software", "三早一晚免勤自动统计"))

    for cand in candidates:
        if not os.path.isdir(cand):
            continue
        core_dir = os.path.join(cand, "core")
        dict_py = os.path.join(core_dir, "dictionary.py")
        if os.path.isdir(core_dir) and os.path.isfile(dict_py):
            return cand

    return None


def resolve_dict_path(given_path: Optional[str]) -> str:
    """Resolve dictionary file path from argument, cwd, or standard locations."""
    if given_path:
        return os.path.abspath(given_path)

    cwd = os.getcwd()
    direct = os.path.join(cwd, "免勤词典.xlsx")
    if os.path.isfile(direct):
        return direct

    # Check release directory or parent directories
    for depth in range(1, 4):
        up = os.path.abspath(os.path.join(cwd, *(['..'] * depth)))
        cand = os.path.join(up, "免勤词典.xlsx")
        if os.path.isfile(cand):
            return cand
        cand_sub = os.path.join(up, "三早一晚免勤自动统计", "免勤词典.xlsx")
        if os.path.isfile(cand_sub):
            return cand_sub

    app_root = discover_app_root()
    if app_root:
        cand_root = os.path.join(app_root, "免勤词典.xlsx")
        if os.path.isfile(cand_root):
            return cand_root

    return direct


def create_timestamped_backup(dict_path: str, custom_out: Optional[str] = None) -> str:
    """Generate timestamped backup with format stem.bak.<timestamp>.xlsx."""
    if not os.path.isfile(dict_path):
        raise FileNotFoundError(f"Dictionary file not found: {dict_path}")

    if custom_out:
        backup_path = os.path.abspath(custom_out)
    else:
        dirname = os.path.dirname(os.path.abspath(dict_path))
        basename = os.path.basename(dict_path)
        stem, ext = os.path.splitext(basename)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(dirname, f"{stem}.bak.{ts}{ext}")

    os.makedirs(os.path.dirname(backup_path), exist_ok=True)
    shutil.copy2(dict_path, backup_path)
    return backup_path


def load_core_modules(app_root: Optional[str] = None):
    """Import and return core.dictionary module."""
    root = discover_app_root(app_root)
    if root and root not in sys.path:
        sys.path.insert(0, root)

    try:
        import core.dictionary as dict_mod
        return dict_mod
    except ImportError as err:
        raise RuntimeError(
            f"Failed to load core.dictionary. Ensure application root is accessible: {err}"
        )


def cmd_view(args: argparse.Namespace) -> int:
    """Display dictionary tables in clean tabular text format."""
    dict_path = resolve_dict_path(args.dict)
    if not os.path.isfile(dict_path):
        print(f"[ERROR] Dictionary file not found: {dict_path}", file=sys.stderr)
        return 1

    dict_mod = load_core_modules(args.app_root)
    try:
        tables = dict_mod.load_tables(dict_path, strict=False)
    except Exception as exc:
        print(f"[ERROR] Failed to read dictionary: {exc}", file=sys.stderr)
        return 1

    target_sheets = [args.sheet] if args.sheet else list(SHEET_HEADERS.keys())
    print(f"=== Dictionary View: {dict_path} ===")
    for sname in target_sheets:
        rows = tables.get(sname, [])
        headers = SHEET_HEADERS.get(sname, ("Col 1", "Col 2", "Col 3"))
        print(f"\n--- Sheet: {sname} ({len(rows)} rules) ---")
        print(f"  {headers[0]:<22} | {headers[1]:<22} | {headers[2]}")
        print("  " + "-" * 70)
        display_rows = rows[:args.limit] if args.limit else rows
        for r in display_rows:
            col0 = r[0] if len(r) > 0 else ""
            col1 = r[1] if len(r) > 1 else ""
            col2 = r[2] if len(r) > 2 else ""
            print(f"  {col0:<22} | {col1:<22} | {col2}")
        if args.limit and len(rows) > args.limit:
            print(f"  ... ({len(rows) - args.limit} more rows omitted)")

    return 0


def cmd_check(args: argparse.Namespace) -> int:
    """Validate dictionary tables and verify reason category whitelist."""
    dict_path = resolve_dict_path(args.dict)
    if not os.path.isfile(dict_path):
        print(f"[ERROR] Dictionary file not found: {dict_path}", file=sys.stderr)
        return 1

    dict_mod = load_core_modules(args.app_root)
    try:
        tables = dict_mod.load_tables(dict_path, strict=True)
        # Validate semantics with Dictionary object
        dict_mod.Dictionary(tables)
    except Exception as exc:
        print(f"[FAIL] Dictionary structural validation failed: {exc}", file=sys.stderr)
        return 1

    # Strict whitelist check for 免勤理由分类
    reason_rows = tables.get("免勤理由分类", [])
    violations = []
    for idx, row in enumerate(reason_rows, start=2):
        if len(row) > 1:
            category = row[1].strip()
            if category and category not in ALLOWED_REASON_CATEGORIES:
                violations.append((idx, row[0], category))

    if violations:
        print(f"[FAIL] Whitelist violation in sheet '免勤理由分类':", file=sys.stderr)
        for row_num, keyword, cat in violations:
            print(
                f"  Row {row_num}: keyword '{keyword}' has category '{cat}' which is not in whitelist {ALLOWED_REASON_CATEGORIES}",
                file=sys.stderr,
            )
        return 1

    print(f"[OK] Dictionary validation passed: {dict_path}")
    print("Rule summary:")
    for sname in SHEET_HEADERS.keys():
        rows = tables.get(sname, [])
        note = " (all categories within strict whitelist)" if sname == "免勤理由分类" else ""
        print(f"  - {sname}: {len(rows)} entries{note}")

    return 0


def cmd_backup(args: argparse.Namespace) -> int:
    """Create timestamped backup file."""
    dict_path = resolve_dict_path(args.dict)
    try:
        backup_path = create_timestamped_backup(dict_path, args.out)
        print(f"[BACKUP] Timestamped backup created successfully: {backup_path}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Backup failed: {exc}", file=sys.stderr)
        return 1


def cmd_add_alias(args: argparse.Namespace) -> int:
    """Add a new alias or reason rule with backup, diff preview, and whitelist check."""
    dict_path = resolve_dict_path(args.dict)
    if not os.path.isfile(dict_path):
        print(f"[ERROR] Dictionary file not found: {dict_path}", file=sys.stderr)
        return 1

    sheet = args.sheet
    if sheet not in SHEET_HEADERS:
        print(f"[ERROR] Invalid sheet '{sheet}'. Must be one of: {list(SHEET_HEADERS.keys())}", file=sys.stderr)
        return 1

    alias = args.alias.strip()
    target = args.target.strip()
    note = (args.note or "Manual addition via dict_tool").strip()

    if not alias or not target:
        print("[ERROR] Both --alias and --target must be non-empty strings.", file=sys.stderr)
        return 1

    # Whitelist check for 免勤理由分类
    if sheet == "免勤理由分类" and target not in ALLOWED_REASON_CATEGORIES:
        print(
            f"[ERROR] Whitelist violation: category '{target}' is not permitted. Allowed: {ALLOWED_REASON_CATEGORIES}",
            file=sys.stderr,
        )
        return 1

    dict_mod = load_core_modules(args.app_root)
    try:
        tables = dict_mod.load_tables(dict_path, strict=False)
    except Exception as exc:
        print(f"[ERROR] Failed loading existing dictionary: {exc}", file=sys.stderr)
        return 1

    # Determine row format according to sheet definition
    if sheet == "学院别名":
        # Formal name, Alias, Note
        new_row = [target, alias, note]
    elif sheet == "班级别名":
        # Alias, Formal name, Note
        new_row = [alias, target, note]
    elif sheet == "校区别名":
        # Raw alias, Canonical campus, Note
        new_row = [alias, target, note]
    elif sheet == "免勤理由分类":
        # Keyword, Category, Note
        new_row = [alias, target, note]
    else:
        new_row = [alias, target, note]

    # Pre-validation BEFORE creating backup:
    import copy
    test_tables = copy.deepcopy(tables)
    test_sheet_rows = test_tables.setdefault(sheet, [])
    test_sheet_rows.append(new_row)

    # 1. Semantic validity test by calling Dictionary(test_tables)
    try:
        _dict_instance = dict_mod.Dictionary(test_tables)
    except Exception as val_err:
        print(f"[ERROR] Dictionary semantic validation failed: {val_err}", file=sys.stderr)
        return 1

    # 2. Strict whitelist check on all rows in 免勤理由分类
    reason_rows = test_tables.get("免勤理由分类", [])
    for r in reason_rows:
        if len(r) >= 2:
            cat = str(r[1]).strip()
            if cat and cat not in ALLOWED_REASON_CATEGORIES:
                print(
                    f"[ERROR] Whitelist violation: category '{cat}' is not permitted. Allowed: {ALLOWED_REASON_CATEGORIES}",
                    file=sys.stderr,
                )
                return 1

    # Step 1: Automatic timestamped backup (ONLY when validation passes!)
    try:
        backup_path = create_timestamped_backup(dict_path)
        print(f"[BACKUP] Created safety backup: {backup_path}")
    except Exception as exc:
        print(f"[ERROR] Aborting: Failed to create safety backup: {exc}", file=sys.stderr)
        return 1

    # Step 2: Display differential preview
    headers = SHEET_HEADERS[sheet]
    print(f"[DIFF] Sheet '{sheet}':")
    print(f"  + {headers[0]}='{new_row[0]}', {headers[1]}='{new_row[1]}', {headers[2]}='{new_row[2]}'")

    # Step 3: Atomic save validated tables
    try:
        dict_mod.save_tables(dict_path, test_tables, backup=False)
        print(f"[SUCCESS] Dictionary updated successfully: {dict_path}")
        print(f"  Total rules in '{sheet}': {len(test_sheet_rows)}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to save updated dictionary: {exc}", file=sys.stderr)
        return 1


def build_parser() -> argparse.ArgumentParser:
    """Build command parser for dict_tool subcommands."""
    parser = argparse.ArgumentParser(description="Dictionary maintenance tool for attendance-stats-calc")
    parser.add_argument("--app-root", default=None, help="Explicit application root directory")
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    # Subcommand: view
    p_view = subparsers.add_parser("view", help="Display dictionary rule tables in tabular format")
    p_view.add_argument("--dict", default=None, help="Path to dictionary workbook (.xlsx)")
    p_view.add_argument("--sheet", choices=list(SHEET_HEADERS.keys()), help="Filter specific sheet")
    p_view.add_argument("--limit", type=int, default=None, help="Limit number of rows displayed")
    p_view.set_defaults(func=cmd_view)

    # Subcommand: check
    p_check = subparsers.add_parser("check", help="Validate table structure and whitelist adherence")
    p_check.add_argument("--dict", default=None, help="Path to dictionary workbook (.xlsx)")
    p_check.set_defaults(func=cmd_check)

    # Subcommand: backup
    p_backup = subparsers.add_parser("backup", help="Create timestamped backup copy")
    p_backup.add_argument("--dict", default=None, help="Path to dictionary workbook (.xlsx)")
    p_backup.add_argument("--out", default=None, help="Optional custom backup output path")
    p_backup.set_defaults(func=cmd_backup)

    # Subcommand: add-alias
    p_add = subparsers.add_parser("add-alias", help="Safely add an alias or reason rule with backup and diff preview")
    p_add.add_argument("--sheet", required=True, choices=["学院别名", "班级别名", "校区别名", "免勤理由分类"], help="Target sheet")
    p_add.add_argument("--alias", required=True, help="Alias string or keyword")
    p_add.add_argument("--target", required=True, help="Target formal name or category")
    p_add.add_argument("--note", default="Manual addition via dict_tool", help="Explanatory note")
    p_add.add_argument("--dict", default=None, help="Path to dictionary workbook (.xlsx)")
    p_add.set_defaults(func=cmd_add_alias)

    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
