# -*- coding: utf-8 -*-
"""离屏渲染桌面界面并保存 PNG，便于人工核对排版。

路径全部由项目根推导，不写死任何本机绝对路径。
"""

import glob
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from PySide6.QtWidgets import QApplication        # noqa: E402

OUT = os.path.join(ROOT, 'docs', '界面截图.png')
WORKSPACE = os.path.dirname(ROOT)


def _find(pattern: str):
    for path in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if pattern in os.path.basename(path):
            return path
    return None


def main() -> int:
    app = QApplication(sys.argv)
    import app as gui
    w = gui.MainWindow()
    w.resize(1440, 900)
    w.show()

    detail = _find('明细表')
    exempt = _find('免勤人员统计表')
    if detail:
        w.detail_slot.set_path(detail)
    if exempt:
        w.exempt_slot.set_path(exempt)

    from core import RunConfig, run
    if detail and exempt:
        res = run(RunConfig(detail_path=detail, exempt_path=exempt, week_no=7,
                            dictionary_path=w.dict_edit.text().strip() or None))
        w._on_done(res)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    app.processEvents()
    w.grab().save(OUT)
    print('已保存界面截图：%s' % OUT)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
