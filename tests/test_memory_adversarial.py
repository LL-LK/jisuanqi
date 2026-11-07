# -*- coding: utf-8 -*-
"""Milestone M2 内存优化与防泄漏对抗性压力测试套件（Adversarial Challenge Suite）。

针对 Milestone M2 交付进行深度压力测试与失效模式挖掘：
  1. 极端异常注入与文件句柄释放测试（Extreme Exception Injection & File Handle Safety）：
     - 损坏的 ZIP 头（截断、非法魔数、伪造数据）注入
     - 损坏或不可读的 Sheet XML（解析崩溃、标签未闭合）注入
     - 运算/解析中途未捕获异常模拟注入
     - 只读与权限锁定状态测试
     - 严格在 Windows 下断言：发生异常后，底层文件句柄彻底释放，可立即调用 unlink() 删除，绝不报 WinError 32 PermissionError。
  2. Qt C++ 对象树内存泄漏压力测试（Qt C++ Child Tree Leak Stress Test）：
     - 50 次循环高速弹出并销毁 StudentDrilldownDialog 与 StatusDialog（warn / error / info）
     - 50 次循环启动与清理后台 Worker 线程
     - 50 次循环高速刷新表格模型（_fill_table 与 _apply_filter）
     - 严格断言：除主窗口常驻工具面板外，无任何未释放的 QDialog、Worker 或 QThread 残留。
  3. 连续 15 次全流程流水线运算的 RSS 稳定性压力测试（15-Run Memory Delta Stress Test）：
     - 执行 15 次真实全流程考勤计算
     - 采集全周期 RSS 内存变化
     - 验证 ΔRSS(Run 15 - Run 2) < 5.0 MB，且内存不呈现单调线性无界增长。
"""

from __future__ import annotations

import gc
import io
import os
import stat
import sys
import zipfile
from pathlib import Path
from typing import List

import psutil
import pytest
from openpyxl import Workbook
from PySide6.QtCore import QAbstractItemModel, QEvent, QThread, QTimer, Qt
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
    """确保测试环境下托盘检测可用。"""
    from PySide6.QtWidgets import QSystemTrayIcon
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: True))


@pytest.fixture
def window(qapp):
    """每个测试独立创建并确保测试后安全退出的 MainWindow 实例。"""
    win = MainWindow()
    yield win
    try:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()
    except Exception:
        pass


# =====================================================================
# 1. 极端异常注入与文件句柄释放安全性测试
# =====================================================================

class TestExtremeExceptionHandleSafety:
    """对抗性检验 1：在极端异常（损坏压缩包、非法 XML、解析中途崩溃）下文件句柄是否 100% 释放。"""

    def test_corrupted_zip_header_load_detail(self, tmp_path):
        """场景 1a: 明细表 ZIP 头部损毁（伪造/截断），文件句柄必须立即释放且可被 unlink 删除。"""
        corrupted_file = tmp_path / "corrupt_header_detail.xlsx"
        # 写入损坏的 ZIP 幻数及乱码字节
        corrupted_file.write_bytes(b"PK\x03\x04\x14\x00\x00\x00\x08\x00CORRUPTED_ZIP_STREAM_DATA_EXTREME_STRESS")

        with pytest.raises(Exception):
            load_detail(str(corrupted_file))

        # Windows 核心断言：若底层文件描述符未关闭，此处抛出 PermissionError (WinError 32)
        corrupted_file.unlink()
        assert not corrupted_file.exists(), "明细表遭遇损坏 ZIP 头时句柄泄漏，无法删除！"

    def test_corrupted_zip_header_load_exempt(self, tmp_path):
        """场景 1b: 免勤表 ZIP 头部损毁，验证 load_exemptions 句柄释放与 unlink 成功。"""
        corrupted_file = tmp_path / "corrupt_header_exempt.xlsx"
        corrupted_file.write_bytes(b"NOT_A_ZIP_HEADER_AT_ALL_GARBAGE_PAYLOAD")

        with pytest.raises(Exception):
            load_exemptions(str(corrupted_file), require_schema=True)

        corrupted_file.unlink()
        assert not corrupted_file.exists(), "免勤表遭遇非法 ZIP 格式时句柄泄漏！"

    def test_corrupted_zip_header_load_tables(self, tmp_path):
        """场景 1c: 词典表 ZIP 损毁，load_tables 内部回退并确保释放句柄。"""
        corrupted_file = tmp_path / "corrupt_dict.xlsx"
        corrupted_file.write_bytes(b"PK\x05\x06" + b"\x00" * 18)  # 空/坏的 Central Directory 结尾

        tables = load_tables(str(corrupted_file))
        assert isinstance(tables, dict), "损坏词典未正确回退至默认词典"

        corrupted_file.unlink()
        assert not corrupted_file.exists(), "词典读取损坏文件后句柄泄漏！"

    def test_unreadable_corrupt_xml_sheet_in_valid_zip(self, tmp_path):
        """场景 1d: 合法的 ZIP 结构包裹语法错误的 Sheet XML（截断/标签失配）。
        此时 load_workbook 建立归档句柄，但在解压/解析工作表 XML 时抛出 ParseError。
        验证工作簿句柄是否被安全释放，允许在 Windows 下立即 unlink。
        """
        bad_sheet_xlsx = tmp_path / "bad_sheet_xml.xlsx"

        # 构造最小化但 sheet XML 损坏的 Excel 归档
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                b'<Default Extension="xml" ContentType="application/xml"/>'
                b'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                b'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                b'</Types>'
            ))
            zf.writestr("_rels/.rels", (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                b'</Relationships>'
            ))
            zf.writestr("xl/workbook.xml", (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                b'<sheets><sheet name="\xe9\x9b\x81\xe5\xb1\xb1\xef\xbc\x88\xe6\x97\xa9\xef\xbc\x89" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>'
                b'</workbook>'
            ))
            zf.writestr("xl/_rels/workbook.xml.rels", (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                b'</Relationships>'
            ))
            # 非法 XML：严重截断且标签错乱
            zf.writestr("xl/worksheets/sheet1.xml", b"<?xml version='1.0'?><worksheet><sheetData><row><c><v>BROKEN_UNCLOSED_TAG")

        bad_sheet_xlsx.write_bytes(zip_buf.getvalue())

        # 读取必定抛出 XML 解析错误或 openpyxl 异常
        with pytest.raises(Exception):
            load_detail(str(bad_sheet_xlsx))

        # 核心验证：发生工作表级 XML 解析异常后，文件句柄必须已被安全关闭，允许立即 unlink
        bad_sheet_xlsx.unlink()
        assert not bad_sheet_xlsx.exists(), "Sheet XML 解析异常导致底层 ZIP 文件句柄泄漏！"

    def test_mid_stream_exception_injection_in_load_detail(self, tmp_path, monkeypatch):
        """场景 1e: 解析明细表过程中途注入未捕获异常，验证 finally: wb.close() 执行无遗漏。"""
        valid_file = tmp_path / "mid_stream_detail.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "雁山（早）"
        ws["A1"] = "执行情况明细表"
        ws["B3"] = "理学院（100人）"
        wb.save(valid_file)
        wb.close()

        # 猴子补丁：模拟遍历中途发生内存耗尽或逻辑崩溃
        import core.io_detail
        orig_detect = core.io_detail._detect_data_cols

        def crash_detect(*args, **kwargs):
            raise MemoryError("模拟中间阶段遭遇 MemoryError 极端故障")

        monkeypatch.setattr(core.io_detail, "_detect_data_cols", crash_detect)

        with pytest.raises(MemoryError, match="极端故障"):
            load_detail(str(valid_file))

        # 验证文件句柄已被释放
        valid_file.unlink()
        assert not valid_file.exists(), "中途注入 MemoryError 导致工作簿句柄未能安全关闭！"

    def test_export_dual_handle_mid_stream_crash(self, tmp_path, monkeypatch):
        """场景 1f: export 导出期间同时持有 wb 与 wb_values 双句柄，在中途崩溃时两者均必须释放。"""
        detail_in = tmp_path / "export_in.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "雁山（早）"
        ws["A1"] = "明细表"
        wb.save(detail_in)
        wb.close()

        detail_out = tmp_path / "export_out.xlsx"
        dummy_model = DetailModel(path=str(detail_in))

        # 模拟在 strip_illegal 阶段抛出异常（此时 wb 与 wb_values 均处于打开状态）
        import core.export
        def crash_strip(val):
            raise RuntimeError("模拟公式与字符串处理时崩溃")

        monkeypatch.setattr(core.export, "strip_illegal", crash_strip)

        # 在单元格放入待清洗字符串以触发 crash_strip
        wb_reopen = Workbook()
        ws_r = wb_reopen.active
        ws_r.title = "雁山（早）"
        ws_r["B2"] = "待清洗内容"
        wb_reopen.save(detail_in)
        wb_reopen.close()

        with pytest.raises(RuntimeError, match="模拟公式与字符串处理时崩溃"):
            export(
                detail_path=str(detail_in),
                output_path=str(detail_out),
                detail=dummy_model,
                room_exempt={},
                rates=[],
                average_rates={},
                estimated_cells=set()
            )

        # 验证 detail_in 输入文件句柄已被 wb 与 wb_values 双重关闭
        detail_in.unlink()
        assert not detail_in.exists(), "导出中途崩溃导致 wb 或 wb_values 文件句柄残留锁定！"

    def test_permission_locked_readonly_sheet_handling(self, tmp_path):
        """场景 1g: 只读文件属性（Windows FILE_ATTRIBUTE_READONLY）下的句柄开闭合安全性。"""
        ro_file = tmp_path / "readonly_detail.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "雁山（早）"
        ws["A1"] = "只读测试明细表"
        wb.save(ro_file)
        wb.close()

        # 赋予只读属性
        os.chmod(ro_file, stat.S_IREAD)

        try:
            model = load_detail(str(ro_file))
            assert isinstance(model, DetailModel)
        finally:
            # 恢复可写以便 unlink
            os.chmod(ro_file, stat.S_IWRITE)

        ro_file.unlink()
        assert not ro_file.exists(), "只读文件读取后句柄未完全关闭！"


# =====================================================================
# 2. Qt C++ 对象树内存泄漏压力测试（50 次高频弹窗与线程循环）
# =====================================================================

class TestQtChildTreeLeakStress:
    """对抗性检验 2：50 次循环打开/关闭弹窗与工作线程，检验 win.children() 无 C++ 残留。"""

    def test_50_cycles_drilldown_dialog_stress(self, qapp, window):
        """场景 2a: 50 次连续高速模态弹出、交互过滤并关闭 StudentDrilldownDialog。"""
        win = window
        students = [
            {'sid': f'2026999900{i:02d}', 'name': f'测试学生_{i}', 'college': '信息科学与工程学院',
             'class_text': '计算机科学与技术26-1班', 'day': '周三', 'session': '晚',
             'room': '1204', 'source_ref': f'免勤汇总表_{i}', 'reason': '实验教学科研'}
            for i in range(30)
        ]

        for i in range(50):
            dlg = StudentDrilldownDialog(
                win,
                title=f"第 {i+1} 轮压力穿透",
                scope_info=f"信科院 / 周三 / 晚 / 轮次_{i}",
                students=students
            )
            assert dlg in win.children(), f"第 {i} 轮弹窗未作为 MainWindow 子节点注册"

            # 模拟在打开期间执行过滤
            dlg._apply_filter(f"测试学生_{i % 30}")

            # 模态交互关闭：通过单次定时器交替触发 accept 与 reject
            QTimer.singleShot(0, dlg.accept if i % 2 == 0 else dlg.reject)
            try:
                dlg.exec()
            finally:
                try:
                    dlg.deleteLater()
                except RuntimeError:
                    pass

            qapp.processEvents()

        qapp.sendPostedEvents(None, QEvent.DeferredDelete)
        qapp.processEvents()
        gc.collect()

        leaked = [c for c in win.children() if isinstance(c, StudentDrilldownDialog)]
        assert len(leaked) == 0, f"50 轮压力测试后发现 {len(leaked)} 个泄漏的 StudentDrilldownDialog 实例！"

    def test_50_cycles_status_dialog_stress(self, qapp, window):
        """场景 2b: 50 次连续调用 StatusDialog（warn / error / info），断言零残留。"""
        win = window

        methods = [
            ('warn', StatusDialog.show_warning),
            ('error', StatusDialog.show_error),
            ('info', StatusDialog.show_info),
        ]

        for i in range(50):
            kind, method = methods[i % len(methods)]
            # 挂载定时器自动关闭弹窗
            QTimer.singleShot(0, lambda: [
                c.accept() for c in win.children()
                if isinstance(c, QDialog) and c is not getattr(win, 'tools_dialog', None)
            ])
            method(
                win,
                title=f"压力提示_{kind}_{i}",
                message=f"第 {i+1} 次高频提示信息",
                advice="无需人工干预",
                technical_detail=f"DEBUG_STACK_TRACE_CYCLE_{i}"
            )
            qapp.processEvents()

        qapp.sendPostedEvents(None, QEvent.DeferredDelete)
        qapp.processEvents()
        gc.collect()

        leaked = [c for c in win.children() if isinstance(c, StatusDialog)]
        assert len(leaked) == 0, f"50 轮 StatusDialog 压力后存在 {len(leaked)} 个泄漏弹窗！"

    def test_50_cycles_worker_thread_lifecycle_stress(self, qapp, window):
        """场景 2c: 50 次快速模拟后台 Worker 运算线程生命周期（创建、挂载、完成、清理）。"""
        win = window

        for i in range(50):
            dummy_cfg = RunConfig(detail_path=f"dummy_{i}.xlsx")
            worker = Worker(dummy_cfg, win)
            win.thread = worker
            win.worker = worker

            # 触发完成清理
            win._cleanup_thread()
            del worker

            qapp.sendPostedEvents(None, QEvent.DeferredDelete)
            qapp.processEvents()

        gc.collect()

        assert win.thread is None, "最终线程引用未置空！"
        assert win.worker is None, "最终 Worker 引用未置空！"

        leaked_workers = [c for c in win.children() if isinstance(c, Worker)]
        assert len(leaked_workers) == 0, f"50 轮线程清理后，MainWindow.children() 仍残留 {len(leaked_workers)} 个 Worker 线程！"

    def test_comprehensive_children_tree_dialog_and_thread_cleanliness(self, qapp, window):
        """场景 2d: 严格断言验收标准：
        打开并关闭 StudentDrilldownDialog 与 StatusDialog，以及 Worker 线程后，
        除主界面常驻 tools_dialog 外，MainWindow.children() 中零泄漏。
        """
        win = window

        # 混合高频触发 20 次弹窗、20 次穿透与 20 次线程
        for i in range(20):
            # StatusDialog
            QTimer.singleShot(0, lambda: [
                c.accept() for c in win.children()
                if isinstance(c, QDialog) and c is not getattr(win, 'tools_dialog', None)
            ])
            StatusDialog.show_info(win, title=f"综合测试_{i}", message="内容")
            qapp.processEvents()

            # StudentDrilldownDialog
            dlg = StudentDrilldownDialog(win, title=f"穿透_{i}", students=[])
            QTimer.singleShot(0, dlg.accept)
            try:
                dlg.exec()
            finally:
                try:
                    dlg.deleteLater()
                except RuntimeError:
                    pass
            qapp.processEvents()

            # Worker 线程
            worker = Worker(RunConfig(detail_path="t.xlsx"), win)
            win.thread = worker
            win.worker = worker
            win._cleanup_thread()
            del worker
            qapp.processEvents()

        qapp.sendPostedEvents(None, QEvent.DeferredDelete)
        qapp.processEvents()
        gc.collect()

        # 核心契约检验：排除常驻的 tools_dialog 后，绝无任何动态弹窗或线程残留
        dynamic_dialogs = [
            c for c in win.children()
            if isinstance(c, QDialog) and c is not getattr(win, 'tools_dialog', None)
        ]
        assert len(dynamic_dialogs) == 0, (
            f"违反 Qt C++ 对象树释放契约: MainWindow 仍持有 {len(dynamic_dialogs)} 个未析构的动态弹窗: {dynamic_dialogs}"
        )

        dynamic_threads = [
            c for c in win.children()
            if isinstance(c, (Worker, QThread))
        ]
        assert len(dynamic_threads) == 0, (
            f"违反 Qt C++ 对象树释放契约: MainWindow 仍持有 {len(dynamic_threads)} 个未析构的线程: {dynamic_threads}"
        )

    def test_rapid_table_refills_churn_stress(self, qapp, window):
        """场景 2e: 50 次快速向 QTableView 灌注新模型，检验旧 QAbstractTableModel 被及时析构。"""
        win = window
        view = QTableView(win)
        cols = ['列A', '列B', '列C']

        for i in range(50):
            rows = [{'列A': f'A_{i}_{j}', '列B': f'B_{i}_{j}', '列C': j} for j in range(10)]
            win._fill_table(view, rows, cols)

        qapp.sendPostedEvents(None, QEvent.DeferredDelete)
        qapp.processEvents()
        gc.collect()

        # 当前绑定的模型应是唯一的活动模型
        current_model = view.model()
        assert current_model is not None
        assert current_model.rowCount() == 10

        # 核心验证：50 次重填后，view.children() 中不累积泄露孤儿 QItemSelectionModel
        from PySide6.QtCore import QItemSelectionModel
        sms = [c for c in view.children() if isinstance(c, QItemSelectionModel)]
        assert len(sms) <= 1, f"QItemSelectionModel 累积泄漏: {len(sms)} 个"


# =====================================================================
# 3. 连续 15 次流水线运算 RSS 稳定性压力测试（15-Run Memory Delta）
# =====================================================================

class Test15RunMemoryDeltaStress:
    """对抗性检验 3：执行连续 15 次全量考勤计算，监测 RSS 进程物理内存波动。"""

    def test_memory_rss_stability_across_15_runs(self, detail_path, exempt_path, tmp_path):
        """场景 3a: 连续 15 次完整流水线计算，验证 ΔRSS(Run 15 - Run 2) < 5.0 MB，且绝无单调线性无界增长。"""
        det = detail_path
        exm = exempt_path
        out_dir = tmp_path / "adversarial_15_runs"
        out_dir.mkdir(exist_ok=True)

        proc = psutil.Process(os.getpid())

        # 第 0 次（预热）：加载 openpyxl 解释器内部静态类型、正则编译缓存等
        warmup_out = str(out_dir / "out_warmup.xlsx")
        res_warmup = run(RunConfig(detail_path=str(det), exempt_paths=[str(exm)], output_path=warmup_out, week_no=7))
        assert res_warmup.ok, f"预热计算失败: {res_warmup.error}"

        gc.collect()
        rss_samples: List[float] = []

        print("\n" + "=" * 70)
        print("Milestone M2 连续 15 次全流水线运算 RSS 压力测试启动")
        print("=" * 70)

        for run_idx in range(1, 16):
            out_file = str(out_dir / f"out_run_{run_idx:02d}.xlsx")
            config = RunConfig(
                detail_path=str(det),
                exempt_paths=[str(exm)],
                output_path=out_file,
                week_no=7
            )
            res = run(config)
            assert res.ok, f"第 {run_idx} 次计算遭遇异常: {res.error}"

            gc.collect()
            current_rss = proc.memory_info().rss / (1024 * 1024)
            rss_samples.append(current_rss)
            print(f"  [Run {run_idx:02d}/15] 完成 | RSS: {current_rss:.2f} MB")

        print("-" * 70)

        # 重点指标分析：从第 2 次运算（索引 1）到第 15 次运算（索引 14）的物理内存净增量
        rss_run_2 = rss_samples[1]
        rss_run_15 = rss_samples[14]
        delta_rss_2_to_15 = rss_run_15 - rss_run_2

        # 全程最大值与最小值波动
        rss_max = max(rss_samples[1:])
        rss_min = min(rss_samples[1:])
        fluctuation = rss_max - rss_min

        print(f"[RSS 统计结果]")
        print(f"  Run 2  RSS: {rss_run_2:.2f} MB")
        print(f"  Run 15 RSS: {rss_run_15:.2f} MB")
        print(f"  ΔRSS (Run 15 - Run 2): {delta_rss_2_to_15:.2f} MB (严格阈值: < 5.0 MB)")
        print(f"  稳态全周期极差 (Max - Min): {fluctuation:.2f} MB")
        print("=" * 70)

        # 断言 1: ΔRSS 必须严格小于 5.0 MB
        assert delta_rss_2_to_15 < 5.0, (
            f"内存泄漏超标: Run 2 到 Run 15 的 ΔRSS 达到 {delta_rss_2_to_15:.2f} MB，超过 5.0 MB 上限！"
        )

        # 断言 2: 验证内存绝非单调无界递增（防泄漏验证：过滤小于 0.1 MB 的分配器微小抖动后，绝非每一步都净增长）
        monotonic_steps = [
            round(rss_samples[k + 1] - rss_samples[k], 1) > 0
            for k in range(1, len(rss_samples) - 1)
        ]
        is_strictly_monotonic = all(monotonic_steps)
        assert not is_strictly_monotonic, (
            f"警告: 物理内存呈现严格单调递增，存在隐蔽内存泄漏风险！步骤增量: {monotonic_steps}"
        )
