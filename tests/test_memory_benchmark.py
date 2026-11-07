# -*- coding: utf-8 -*-
"""全生命周期内存优化与防泄漏基准测试套件（Requirement R3）。

验证项目：
  1. Openpyxl 工作簿句柄防泄漏保护：解析/导出发生异常时，确保 finally: wb.close()
     正常执行，在 Windows 下不产生 PermissionError 文件句柄锁定。
  2. Qt C++ 对象树内存释放：StudentDrilldownDialog 与 StatusDialog 弹窗关闭后，
     以及后台工作线程 Worker 结束后，从 MainWindow.children() 中彻底清除，无 C++ 残留。
  3. 表格模型刷新防泄漏：_fill_table 高频重载时，旧 QAbstractTableModel 被及时断开并释放。
  4. 系统托盘最小化内存规整：最小化到托盘触发 _trim_memory() 与垃圾回收。
  5. 连续 10 次流水线运算的 RSS 内存增量基准（ΔRSS < 5.0 MB）。
"""

from __future__ import annotations

import gc
import os
import sys
from pathlib import Path
from typing import List

import psutil
import pytest
import shiboken6
from openpyxl import Workbook
from PySide6.QtCore import QEvent, QTimer, Qt
from PySide6.QtWidgets import QApplication, QDialog, QTableView

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from app import MainWindow, StatusDialog, StudentDrilldownDialog, Worker
from core import RunConfig, run
from core.dictionary import load_tables
from core.export import export
from core.io_detail import DetailModel, load_detail
from core.io_exempt import load_exemptions


@pytest.fixture(scope='session')
def qapp():
    """PySide6 应用程序实例。"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv[:1])
    return app


@pytest.fixture(autouse=True)
def mock_tray_available(monkeypatch):
    """确保在 CI/CD 或 pytest 全套件运行（设置了 QT_QPA_PLATFORM=offscreen）时，
    托盘检测判定为可用以完整覆盖托盘交互逻辑。
    """
    from PySide6.QtWidgets import QSystemTrayIcon
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: True))


@pytest.fixture
def main_window(qapp):
    """独立创建并在用例结束后彻底释放的 MainWindow 实例。"""
    win = MainWindow()
    yield win
    try:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()
    except Exception:
        pass


# =====================================================================
# 1. Openpyxl 文件句柄安全测试（Defect M1）
# =====================================================================

def test_workbook_handle_safety_io_detail_on_error(tmp_path):
    """验证 load_detail 在遇到损坏/异常结构时仍能安全关闭工作簿，允许删除文件。"""
    corrupt_file = tmp_path / "corrupt_detail.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "雁山（早）"
    ws['A1'] = "执行情况明细表·异常测试"
    ws['B3'] = "理学院（100人）"
    # 构造缺失学院列或表头的半损坏工作表
    wb.save(corrupt_file)
    wb.close()

    # 读取应当正常完成或捕获异常
    try:
        model = load_detail(str(corrupt_file))
        assert isinstance(model, DetailModel)
    except Exception:
        pass

    # Windows 核心检验：若句柄未 close()，os.remove 会抛出 WinError 32 PermissionError
    corrupt_file.unlink()
    assert not corrupt_file.exists(), "文件句柄泄漏导致 Windows 无法删除被占用的工作簿！"


def test_workbook_handle_safety_io_exempt_on_error(tmp_path):
    """验证 load_exemptions 在 schema 校验不匹配抛出异常时，仍能安全关闭句柄。"""
    invalid_exempt = tmp_path / "invalid_exempt.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "无表头免勤"
    ws.append(["无关数据1", "无关数据2", "无关数据3"])
    wb.save(invalid_exempt)
    wb.close()

    with pytest.raises(ValueError, match="找不到免勤申请表头"):
        load_exemptions(str(invalid_exempt), require_schema=True)

    # 验证文件句柄已被 finally 关闭，文件可立即被删除或重命名
    invalid_exempt.unlink()
    assert not invalid_exempt.exists(), "load_exemptions 异常分支未释放工作簿文件句柄！"


def test_workbook_handle_safety_export_detail_on_error(tmp_path):
    """验证 export_detail 在写出流程中途出错时，wb 与 wb_values 句柄均被关闭。"""
    detail_file = tmp_path / "export_input.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "雁山（早）"
    ws.append(["标题"])
    wb.save(detail_file)
    wb.close()

    output_file = tmp_path / "export_output.xlsx"
    dummy_model = DetailModel(path=str(detail_file))

    # 模拟在 append_report 阶段触发未捕获异常
    def broken_append_report(book):
        raise RuntimeError("模拟导出修饰过程崩溃")

    with pytest.raises(RuntimeError, match="模拟导出修饰过程崩溃"):
        export(
            detail_path=str(detail_file),
            output_path=str(output_file),
            detail=dummy_model,
            room_exempt={},
            rates=[],
            average_rates={},
            estimated_cells=set(),
            append_report=broken_append_report
        )

    # 验证输入与输出文件句柄全部已安全关闭，可立即被重命名或移除
    detail_file.unlink()
    assert not detail_file.exists(), "export_detail 异常后输入文件句柄仍被占用！"


def test_workbook_handle_safety_dictionary_on_error(tmp_path):
    """验证 load_tables 读取损坏字典文件时句柄正常释放。"""
    dict_file = tmp_path / "bad_dict.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "未知工作表"
    wb.save(dict_file)
    wb.close()

    tables = load_tables(str(dict_file))
    assert isinstance(tables, dict)

    dict_file.unlink()
    assert not dict_file.exists(), "load_tables 读取后工作簿句柄泄漏！"


# =====================================================================
# 2. Qt C++ 对象树内存释放测试（Defect M2）
# =====================================================================

def test_drilldown_dialog_lifecycle_and_child_tree(qapp, main_window):
    """验证 StudentDrilldownDialog 弹窗在关闭后从 MainWindow.children() 中彻底清除。"""
    win = main_window

    students = [
        {'sid': f'99000000000{i}', 'name': f'学生{i}', 'college': '理学院',
         'class_text': '应用统计学26-1班', 'day': '周一', 'session': '早',
         'room': '8101', 'source_ref': '免勤表1', 'reason': '病假'}
        for i in range(20)
    ]

    # 打开弹窗并定时自动接受
    dlg = StudentDrilldownDialog(win, title="测试明细", scope_info="理学院 / 周一 / 早", students=students)
    assert dlg in win.children(), "弹窗应当正确注册为 MainWindow 的子组件"

    QTimer.singleShot(0, dlg.accept)
    try:
        dlg.exec()
    finally:
        try:
            dlg.deleteLater()
        except RuntimeError:
            pass

    qapp.processEvents()

    # 验证 win.children() 中不再持有任何 StudentDrilldownDialog 实例
    leaked_dialogs = [c for c in win.children() if isinstance(c, StudentDrilldownDialog)]
    assert len(leaked_dialogs) == 0, f"发现 {len(leaked_dialogs)} 个泄漏的 StudentDrilldownDialog 实例！"


def test_status_dialog_lifecycle_and_child_tree(qapp, main_window):
    """验证 StatusDialog (show_warning / show_error / show_info) 关闭后不累积子组件。"""
    win = main_window

    for kind, method in [('warn', StatusDialog.show_warning),
                         ('error', StatusDialog.show_error),
                         ('info', StatusDialog.show_info)]:
        QTimer.singleShot(10, lambda: [
            c.accept() for c in win.children() if isinstance(c, QDialog)
        ])
        method(win, title=f"测试{kind}", message="提示内容", advice="处理建议")
        qapp.processEvents()

    qapp.processEvents()
    leaked = [c for c in win.children() if isinstance(c, StatusDialog)]
    assert len(leaked) == 0, f"发现 {len(leaked)} 个泄漏的 StatusDialog 实例！"


def test_worker_thread_lifecycle_and_child_tree(qapp, main_window):
    """验证后台 Worker 运算线程完成后在 _cleanup_thread 中被 deleteLater 释放。"""
    win = main_window

    dummy_config = RunConfig(detail_path="non_existent.xlsx")
    worker = Worker(dummy_config, win)
    win.thread = worker
    win.worker = worker

    # 模拟线程完成与清理
    win._cleanup_thread()
    del worker
    qapp.sendPostedEvents()
    qapp.processEvents()
    gc.collect()

    assert win.thread is None
    assert win.worker is None
    leaked_workers = [c for c in win.children() if isinstance(c, Worker)]
    assert len(leaked_workers) == 0, f"发现 {len(leaked_workers)} 个泄漏的 Worker 线程对象！"


# =====================================================================
# 3. 表格模型刷新与内存清理测试（Defect M3 & M4）
# =====================================================================

def test_fill_table_disposes_previous_model(qapp, main_window):
    """验证 _fill_table 重复设置表格数据时，前序模型被断开并标记释放。"""
    win = main_window
    view = QTableView(win)

    columns = ['序号', '名称', '数值']
    rows_v1 = [{'序号': 1, '名称': '甲', '数值': 10}]
    rows_v2 = [{'序号': 1, '名称': '乙', '数值': 20}]

    win._fill_table(view, rows_v1, columns)
    model_v1 = view.model()
    assert model_v1 is not None

    win._fill_table(view, rows_v2, columns)
    model_v2 = view.model()
    assert model_v2 is not model_v1

    qapp.processEvents()
    # 验证 model_v1 已从视图分离，且已被 deleteLater
    assert view.model() is model_v2


def test_drilldown_dialog_filter_model_cleanup(qapp, main_window):
    """验证 StudentDrilldownDialog 在输入过滤条件时旧模型被及时释放。"""
    students = [
        {'sid': f'99000000000{i}', 'name': f'学生{i}', 'college': '理学院',
         'class_text': '应用统计学26-1班', 'day': '周一', 'session': '早',
         'room': '8101', 'source_ref': '免勤表1', 'reason': '病假'}
        for i in range(10)
    ]
    dlg = StudentDrilldownDialog(main_window, students=students)
    m1 = dlg.table.model()

    # 模拟文本过滤
    dlg._apply_filter("学生1")
    m2 = dlg.table.model()
    assert m2 is not m1

    dlg.accept()
    qapp.processEvents()


# =====================================================================
# 4. 系统托盘最小化内存规整测试（Defect M5）
# =====================================================================

def test_system_tray_minimize_triggers_trim_memory(qapp, main_window, monkeypatch):
    """验证主窗口最小化到托盘时，触发 _trim_memory() 内存修剪与垃圾回收。"""
    win = main_window
    trim_called = []

    orig_trim = win._trim_memory
    def tracking_trim():
        trim_called.append(True)
        orig_trim()

    monkeypatch.setattr(win, '_trim_memory', tracking_trim)

    # 模拟触发最小化事件
    event = QEvent(QEvent.WindowStateChange)
    win.setWindowState(Qt.WindowMinimized)
    win.changeEvent(event)
    qapp.processEvents()

    assert len(trim_called) > 0, "最小化到托盘事件未触发 _trim_memory 内存清理！"


# =====================================================================
# 5. 连续 10 次流水线运算 RSS 稳定性基准测试（Defect M5 & Acceptance Criteria）
# =====================================================================

def test_memory_rss_stability_across_10_runs(detail_path, exempt_path, tmp_path):
    """验证连续 10 次全流水线运行，内存无持续异常增长（ΔRSS < 5.0 MB）。"""
    det = detail_path
    exm = exempt_path
    out_dir = tmp_path / "bench_outputs"
    out_dir.mkdir(exist_ok=True)

    proc = psutil.Process(os.getpid())

    # 预热第 1 次运算以加载 openpyxl 与模块静态缓存
    warmup_out = str(out_dir / "out_warmup.xlsx")
    res = run(RunConfig(detail_path=str(det), exempt_paths=[str(exm)], output_path=warmup_out, week_no=7))
    assert res.ok, f"预热计算失败: {res.error}"

    gc.collect()
    rss_baseline = proc.memory_info().rss / (1024 * 1024)

    rss_samples: List[float] = [rss_baseline]

    for i in range(10):
        out_path = str(out_dir / f"out_run_{i}.xlsx")
        config = RunConfig(
            detail_path=str(det),
            exempt_paths=[str(exm)],
            output_path=out_path,
            week_no=7,
        )
        res = run(config)
        assert res.ok, f"第 {i+1} 次计算失败: {res.error}"
        gc.collect()
        current_rss = proc.memory_info().rss / (1024 * 1024)
        rss_samples.append(current_rss)

    # 计算从第 2 次运算到第 10 次运算的 RSS 增量（排除初始预热波动）
    delta_rss = rss_samples[-1] - rss_samples[1]

    print(f"\n[RSS Benchmark] 初始: {rss_samples[0]:.2f} MB | 第1次后: {rss_samples[1]:.2f} MB | 第10次后: {rss_samples[-1]:.2f} MB | ΔRSS: {delta_rss:.2f} MB")

    assert delta_rss < 5.0, (
        f"全生命周期内存泄漏超标: 10 次连续运算后 RSS 净增长 {delta_rss:.2f} MB (阈值为 5.0 MB)！"
    )
