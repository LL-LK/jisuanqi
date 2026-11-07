# -*- coding: utf-8 -*-
"""测试夹具与构造助手（本套件私有模块）。

为什么单独放一个模块：``tests/conftest.py`` 由测试套件统一维护，
而本模块里的夹具实现被 ``conftest.py`` 与各测试模块同时引用，
这样即使上层 conftest 被改动，本套件仍能独立跑通。

约定（与项目约束一致）：

  * 所有路径都由 ``Path(__file__).resolve().parents[1]`` 推导，测试文件里
    不出现任何绝对路径；
  * 学号 / 姓名 / 学院名 / 教室号全部是虚构占位；
  * 测试只使用 ``tmp_path`` 下的合成样本，**不读取、不复制**工作区里的
    真实 Excel（真实表含学生个人信息）。
"""

from __future__ import annotations

import hashlib
import importlib.util
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _load_module(name: str, rel_path: str):
    """按文件路径加载模块（tools/ 不是包，不改动项目结构）。"""
    spec = importlib.util.spec_from_file_location(name, ROOT / rel_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


gen_synthetic = _load_module('gen_synthetic', 'tools/gen_synthetic.py')

from core import RunConfig, run          # noqa: E402


# ------------------------------------------------------------------ 小工具


def md5(path) -> str:
    """文件 md5（用来证明输入文件没有被覆盖）。"""
    digest = hashlib.md5()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            digest.update(chunk)
    return digest.hexdigest()


def workbook_digest(path):
    """工作簿内容摘要（按单元格取值比较，忽略 zip 里的时间戳差异）。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    try:
        return tuple(
            (name, tuple(
                tuple(wb[name].cell(r, c).value
                      for c in range(1, wb[name].max_column + 1))
                for r in range(1, wb[name].max_row + 1)))
            for name in wb.sheetnames if name != '_系统校验元数据')
    finally:
        wb.close()


def run_engine(detail_path, exempt_path=None, output_path=None, week_no=7,
               campus='雁山'):
    """跑一次引擎（不传词典 → 使用内置默认词典，不依赖项目数据）。"""
    config = RunConfig(
        detail_path=str(detail_path),
        exempt_path=str(exempt_path) if exempt_path else None,
        output_path=str(output_path) if output_path else None,
        week_no=week_no,
        campus=campus,
        dictionary_path=None,
    )
    return run(config)


# ------------------------------------------------------------------ 构造助手


def make_exempt_row(sid: str = '9900000000001', college: str = '理学院',
                    name: str = '学生甲01', class_text: str = '应用统计学26-1班',
                    week_text: str = '第7周', day_text: str = '周一',
                    session: str = '早', campus: str = '雁山校区',
                    reason: str = '身体不适', room: str = ''):
    """一条虚构的免勤申请。"""
    return gen_synthetic.ExemptRow(
        seq=0, sid=sid, college=college, name=name, class_text=class_text,
        week_text=week_text, day_text=day_text, session=session,
        campus=campus, reason=reason, room=room)


def make_detail_sheet(name: str,
                      blocks: Sequence[Tuple[str, int, Sequence[Tuple[str, int, Dict]]]],
                      rate_labels=None, start_row: int = 4):
    """构造一张明细工作表。

    blocks 的每一项是 ``(学院名, 学院标注总人数, [(教室号, 应到人数, 实到字典), ...])``。
    """
    spec = gen_synthetic.SheetSpec(name=name, title='测试用·%s' % name)
    spec.rate_labels = rate_labels or gen_synthetic.RATE_LABELS_EARLY
    row = start_row
    seq = 1
    for college, total, rooms in blocks:
        blk = gen_synthetic.BlockSpec(row, row + len(rooms) - 1, college,
                                      college_total=total)
        for room, expected, actuals in rooms:
            spec.rows.append(gen_synthetic.DetailRow(
                row=row, seq=seq, college=college, college_total=total,
                room=room, classes_text='测试班26-1班', expected=expected,
                actuals=dict(actuals)))
            seq += 1
            row += 1
        spec.blocks.append(blk)
    return spec


def write_detail(path, sheets) -> Path:
    gen_synthetic.write_detail_workbook(Path(path), sheets)
    return Path(path)


def write_exempt(path, sheets: Dict[str, List], with_room: bool = False) -> Path:
    gen_synthetic.write_exempt_workbook(Path(path), sheets, with_room=with_room)
    return Path(path)


def focused_sheets() -> List:
    """最小样本：理学院 2 间教室（早/晚各一张表）+ 1 张屏风表。"""
    early = make_detail_sheet('雁山（早）', [
        ('理学院', 100, [
            ('8101', 40, {'周一': 30, '周三': 28, '周四': 29, '周五': 30}),
            ('8102', 60, {'周一': 45, '周三': 44, '周四': 46, '周五': 47}),
        ]),
    ])
    late = make_detail_sheet('雁山（晚）', [
        ('理学院', 100, [
            ('8101', 40, {'周一': 32, '周三': 31, '周四': 33, '周五': 34}),
            ('8102', 60, {'周一': 47, '周三': 46, '周四': 48, '周五': 49}),
        ]),
    ], rate_labels=gen_synthetic.RATE_LABELS_LATE)
    pingfeng = make_detail_sheet('屏风', [
        ('化工学院', 50, [
            ('9201', 50, {'周一': 45, '周三': 44, '周四': 46, '周五': 47}),
        ]),
    ])
    return [early, late, pingfeng]


def build_focused_sample(base: Path, early_rows: Optional[List] = None,
                          late_rows: Optional[List] = None,
                          with_room: bool = False) -> Tuple[Path, Path]:
    """写出一份最小样本，返回 (detail_path, exempt_path)。"""
    base = Path(base)
    detail = base / 'detail_focused.xlsx'
    exempt = base / 'exempt_focused.xlsx'
    write_detail(detail, focused_sheets())
    write_exempt(exempt, {
        '早自习免勤人员统计表': list(early_rows or []),
        '晚自习免勤人员统计表': list(late_rows or []),
    }, with_room=with_room)
    return detail, exempt


def read_exempt_rows(path, sheet: str = '早自习免勤人员统计表'):
    """读取免勤表 → ExemptRecord 列表（走完整的解析链路）。"""
    from core.io_exempt import load_exemptions
    return load_exemptions(str(path), source='未分类')


def exempt_records(base: Path, rows, sheet: str = '早自习免勤人员统计表',
                   with_room: bool = False, name: str = 'exempt_rows.xlsx'):
    """把虚构行写成免勤表再读回来（真实链路：解析 + 归一化）。"""
    path = write_exempt(Path(base) / name, {sheet: list(rows)},
                        with_room=with_room)
    return read_exempt_rows(path, sheet)


# ------------------------------------------------------------------ 夹具


@pytest.fixture(scope='session')
def generator():
    """合成数据生成器模块。"""
    return gen_synthetic


@pytest.fixture(scope='session')
def synthetic_dir(tmp_path_factory) -> Path:
    """干净的合成样本目录（tmp_path 下，不污染项目目录）。"""
    out = tmp_path_factory.mktemp('synthetic')
    gen_synthetic.generate(out, scale='small', dirty=False, seed=20261025)
    return out


@pytest.fixture(scope='session')
def detail_path(synthetic_dir) -> Path:
    return synthetic_dir / 'detail.xlsx'


@pytest.fixture(scope='session')
def exempt_path(synthetic_dir) -> Path:
    return synthetic_dir / 'exempt.xlsx'


@pytest.fixture(scope='session')
def run_result(synthetic_dir, detail_path, exempt_path):
    """对干净样本跑一次引擎的结果。"""
    result = run_engine(detail_path, exempt_path,
                        output_path=synthetic_dir / 'result.xlsx')
    assert result.ok, result.error
    return result


@pytest.fixture(scope='session')
def output_path(synthetic_dir, run_result) -> Path:
    """引擎产物的路径（依赖 run_result，保证文件已生成）。"""
    return synthetic_dir / 'result.xlsx'


@pytest.fixture(scope='session')
def output_workbook(output_path):
    from openpyxl import load_workbook
    wb = load_workbook(output_path, data_only=True)
    yield wb
    wb.close()


@pytest.fixture()
def focused_sample(tmp_path) -> Tuple[Path, Path]:
    """最小样本（明细 + 空免勤表）。"""
    return build_focused_sample(tmp_path)
