# -*- coding: utf-8 -*-
"""组装发布目录：复制 EXE / 词典 / 文档，并生成 SHA-256 校验文件。

发布目录纪律：
  1. 组装前先清空旧文件——否则上一版的 EXE / 发布包会残留在目录里，
     出现「清单里是新版、目录里混着旧版」的静默不一致。
  2. 清单覆盖目录内**全部**文件（ZIP 发布包放在同级的 `发布包/`，不混进来）。
  3. 清单用 UTF-8 写；文件名是中文，必须由 tools/verify_release.py 以 UTF-8 读回校验，
     手工在 GBK 控制台比对必然对不上。
"""

from __future__ import annotations

import hashlib
import os
import shutil

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REL = os.path.join(ROOT, '发布版')
MANIFEST = 'SHA256.txt'

ITEMS = [
    (os.path.join(ROOT, 'dist', 'AttendanceStats.exe'), '三早一晚免勤自动统计.exe'),
    (os.path.join(ROOT, '免勤词典.xlsx'), '免勤词典.xlsx'),
    (os.path.join(ROOT, '使用说明.md'), '使用说明.md'),
    (os.path.join(ROOT, 'README.md'), 'README.md'),
    (os.path.join(ROOT, '隐私扫描报告.md'), '隐私扫描报告.md'),
]


def clean_release_dir() -> None:
    """清空发布目录里的旧文件（不递归删子目录，避免误删）。"""
    if not os.path.isdir(REL):
        os.makedirs(REL, exist_ok=True)
        return
    for name in os.listdir(REL):
        path = os.path.join(REL, name)
        if os.path.isfile(path):
            os.remove(path)
            print('清理旧文件：%s' % name)


def main() -> int:
    clean_release_dir()

    for src, name in ITEMS:
        if not os.path.exists(src):
            print('跳过（不存在）：%s' % os.path.basename(src))
            continue
        shutil.copy2(src, os.path.join(REL, name))
        print('已复制：%s' % name)

    lines = []
    for name in sorted(os.listdir(REL)):
        path = os.path.join(REL, name)
        if not os.path.isfile(path) or name == MANIFEST:
            continue
        digest = hashlib.sha256(open(path, 'rb').read()).hexdigest()
        lines.append('%s  %s' % (digest, name))
    with open(os.path.join(REL, MANIFEST), 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines) + '\n')
    print('\n=== SHA256 ===')
    print('\n'.join(lines))
    print('\n发布目录：%s' % REL)
    print('校验命令：python -X utf8 tools/verify_release.py 发布版')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
