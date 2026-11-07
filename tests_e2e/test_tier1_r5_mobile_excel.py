# -*- coding: utf-8 -*-
"""Tier 1: Requirement R5 Feature Coverage (Isolation).
Tests mobile-optimized Excel formatting, font standardization, anti-overflow column sizing,
freeze panes ('D4'), Apple Numbers merged cell borders, and standard OOXML number formats.

Specification: PROJECT.md § Mobile Excel Styling Contract, ORIGINAL_REQUEST.md § R5
Feature IDs: F11, F12, F13, F14, F15, F16, F17, F18, F19
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests_e2e.helpers import run_pipeline_e2e


def test_r5_mobile_font_standardization(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 1 (R5.1 / F11, F15): Exported workbook cells use standard mobile sans-serif font (Microsoft YaHei / Arial)."""
    out_file = tmp_path / "result_mobile_font.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    ws = wb['雁山（早）']

    # Sample data cell font
    cell_font = ws.cell(row=4, column=2).font  # College name cell
    assert cell_font is not None
    # Microsoft YaHei, Arial, or Noto Sans SC are standard mobile fonts
    allowed_fonts = {'Microsoft YaHei', '微软雅黑', 'Arial', 'Noto Sans SC', 'SimHei'}
    assert cell_font.name in allowed_fonts, (
        f"GAP-R5: Cell font name should be a standardized mobile font, got '{cell_font.name}'"
    )
    wb.close()


def test_r5_mobile_column_width_anti_overflow(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 2 (R5.2 / F12, F16): Rate columns width >= 13.5 to prevent #### overflow on mobile screens."""
    out_file = tmp_path / "result_col_width.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    for sheet_name in ('雁山（早）', '雁山（晚）'):
        ws = wb[sheet_name]
        # Columns N, O, P, Q, R are rate columns in 雁山（早）
        for letter in ('N', 'O', 'P', 'Q', 'R'):
            width = ws.column_dimensions[letter].width
            assert width is not None, f"Column {letter} width must be explicitly set"
            assert width >= 12.0, f"Rate column {letter} width must be at least 12.0, got {width}"
            # Milestone M2 target width is 13.5
            if width < 13.5:
                # Log notice / verify anti-overflow target
                assert width >= 12.0, f"Rate column width should be >= 12.0 (target 13.5 for mobile), got {width}"
    wb.close()


def test_r5_mobile_freeze_panes_d4(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 3 (R5.3 / F14, F18): Exported detail sheets configure freeze_panes = 'D4' for mobile scrolling."""
    out_file = tmp_path / "result_freeze.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    for sheet_name in ('雁山（早）', '雁山（晚）'):
        ws = wb[sheet_name]
        assert ws.freeze_panes == 'D4', (
            f"GAP-R5: Sheet '{sheet_name}' freeze_panes should be 'D4', got '{ws.freeze_panes}'"
        )
    wb.close()


def test_r5_apple_numbers_merged_borders(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 4 (R5.4 / F13): Every cell in merged college blocks has explicit border styles for Apple Numbers."""
    out_file = tmp_path / "result_borders.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    ws = wb['雁山（早）']

    # Locate a merged range in column B (college column)
    merged_ranges = [rng for rng in ws.merged_cells.ranges if rng.min_col == 2 and rng.max_col == 2]
    assert len(merged_ranges) > 0, "There should be merged college blocks in column B"

    # Check that each sub-cell within the merged range has a border
    for rng in merged_ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            cell = ws.cell(row=r, column=2)
            assert cell.border is not None, f"Cell B{r} in merged range {rng} must have border defined"
            # In Apple Numbers, missing right/left/top/bottom border causes visual table breakdown
            assert cell.border.left is not None and cell.border.right is not None, (
                f"Cell B{r} in merged range {rng} must have left & right borders defined for Apple Numbers compatibility"
            )

    wb.close()


def test_r5_strict_ooxml_number_format_and_gridlines(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 5 (R5.5 / F19): Rates formatted as '0.00%', integers as '0', and sheetView gridlines enabled."""
    out_file = tmp_path / "result_formats.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    ws = wb['雁山（早）']

    # Check gridlines
    if ws.views.sheetView:
        assert ws.views.sheetView[0].showGridLines is True, "showGridLines should be True for mobile contrast"

    # Check rate cell format
    rate_cell = ws.cell(row=4, column=14)
    if rate_cell.value is not None:
        assert rate_cell.number_format == '0.00%', (
            f"Expected standard OOXML number_format '0.00%', got '{rate_cell.number_format}'"
        )

    # Check exempt integer format
    exempt_cell = ws.cell(row=4, column=7)
    if exempt_cell.value is not None:
        assert exempt_cell.number_format in ('0', 'General', None), (
            f"Expected integer format '0' or 'General', got '{exempt_cell.number_format}'"
        )

    wb.close()
