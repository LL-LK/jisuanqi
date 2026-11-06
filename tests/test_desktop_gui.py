# -*- coding: utf-8 -*-
"""桌面界面冒烟与不变量（PySide6，offscreen 模式）。

界面层只验证「四个能力面板在位 + 关键硬规则不被界面破坏」，
不做像素级/布局级断言——那是发布工程师与需求方的验收范围。
"""

from __future__ import annotations

import os
import sys

import pytest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

pytest.importorskip('PySide6', reason='未安装 PySide6，跳过桌面界面用例')

from PySide6.QtWidgets import QApplication       # noqa: E402

import app as gui                                 # noqa: E402


@pytest.fixture(scope='session')
def qapp():
    app = QApplication.instance() or QApplication(sys.argv[:1])
    yield app


@pytest.fixture()
def window(qapp):
    win = gui.MainWindow()
    yield win
    win.close()


def test_main_window_exposes_four_capability_panels(window):
    """主页只展示四项，辅助功能保留在概览工具窗口。"""
    titles = [window.tabs.tabText(i) for i in range(window.tabs.count())]
    assert titles == ['概览', '学院免勤统计', '出勤率', '词典编辑']
    assert [window.rate_views.tabText(i) for i in range(window.rate_views.count())] == ['每日', '每周']
    tool_titles = [window.tools_tabs.tabText(i) for i in range(window.tools_tabs.count())]
    assert tool_titles == ['教室明细', '异常清单', '低出勤率复核', '合成数据自测', '可选设置']


def test_window_title_has_name_and_version(window):
    assert gui.APP_NAME in window.windowTitle()
    assert gui.APP_VERSION in window.windowTitle()


def test_dictionary_panel_loads_tables(window):
    """词典编辑面板必须加载出 5 张词表。"""
    from core.dictionary import SHEET_NAMES
    assert set(window.dict_tables) >= set(SHEET_NAMES)


def test_dictionary_model_append_and_edit():
    """词典行可新增、可编辑，改动能读回来。"""
    model = gui._DictModel('学院别名', [['测试学院甲', '别名甲', '']])
    model.append_blank()
    assert model.rowCount() == 2
    index = model.index(1, 0)
    assert model.setData(index, '新增正式名称') is True
    assert model.data(index) == '新增正式名称'


def test_resource_path_falls_back_to_project_dir():
    """打包后 resource_path 必须仍能定位到词典文件。"""
    path = gui.resource_path(gui.DICT_NAME)
    assert os.path.exists(path)


def test_selfcheck_command_is_wellformed(monkeypatch, window):
    """「合成数据自测」按钮必须调用 pytest tests -q，且工作目录是项目根。

    v1.0.3 起自测改在后台线程执行（避免冻结界面），命令构造在 _SelfCheckThread 里。
    """
    captured = {}

    class _Proc:
        stdout = 'ok'
        stderr = ''
        returncode = 0

    def _fake_run(cmd, **kwargs):
        captured['cmd'] = cmd
        captured['cwd'] = kwargs.get('cwd')
        return _Proc()

    monkeypatch.setattr(gui.subprocess, 'run', _fake_run)
    thread = gui._SelfCheckThread()
    thread.run()                      # 直接跑线程体，避免真的起线程
    assert 'pytest' in captured['cmd']
    assert captured['cmd'][-2:] == ['tests', '-q']
    assert os.path.isdir(captured['cwd'])


def test_low_rate_threshold_is_in_range():
    """低出勤率复核阈值必须落在 (0, 1] 内。"""
    assert 0 < gui.LOW_RATE_THRESHOLD <= 1
