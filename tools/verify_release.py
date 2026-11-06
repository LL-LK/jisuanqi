# -*- coding: utf-8 -*-
"""校验发布目录的 SHA-256 清单。

为什么需要它：SHA256.txt 用 UTF-8 写、文件名是中文，在 GBK 控制台里
`Get-Content` 显示为乱码，手工比对必然对不上（实测会出现 `涓夋棭...exe`）。
所以校验必须由程序以 UTF-8 读清单来做，用户侧只需跑一个命令。

用法：
    python -X utf8 tools/verify_release.py 发布版
退出码：0 = 全部一致；1 = 不一致或清单缺项/多项。
"""
from __future__ import annotations

import hashlib
import os
import sys


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def verify_package(pkg_dir: str) -> int:
    """校验发布包目录：每个 .zip 旁边必须有同名 .sha256 且哈希一致。"""
    if not os.path.isdir(pkg_dir):
        return 0
    bad = 0
    zips = [n for n in sorted(os.listdir(pkg_dir)) if n.lower().endswith('.zip')]
    for name in zips:
        side = os.path.join(pkg_dir, name + '.sha256')
        path = os.path.join(pkg_dir, name)
        if not os.path.exists(side):
            print('[MISS ] %s  (no .sha256 sidecar)' % name)
            bad += 1
            continue
        with open(side, encoding='utf-8') as fh:
            expected = fh.read().split()[0].strip().lower()
        actual = sha256_file(path)
        if expected != actual:
            print('[BAD  ] %s' % name)
            print('        sidecar: %s' % expected)
            print('        actual : %s' % actual)
            bad += 1
        else:
            print('[OK   ] %s  %s' % (name, actual[:16]))
    if not zips:
        print('[NOTE ] 发布包目录内没有 zip（尚未执行 make_zip）')
    return bad


def main(argv=None) -> int:
    argv = argv or sys.argv[1:]
    if not argv:
        print('usage: verify_release.py <release_dir> [package_dir]')
        return 2
    rel = argv[0]
    pkg = argv[1] if len(argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(rel)), '发布包')
    manifest = os.path.join(rel, 'SHA256.txt')
    if not os.path.exists(manifest):
        print('[ERROR] manifest not found: SHA256.txt')
        return 1

    with open(manifest, encoding='utf-8') as fh:
        entries = {}
        for line in fh:
            line = line.strip()
            if not line:
                continue
            digest, _, name = line.partition('  ')
            entries[name.strip()] = digest.strip().lower()

    on_disk = {n for n in os.listdir(rel)
               if os.path.isfile(os.path.join(rel, n)) and n != 'SHA256.txt'}

    bad = 0
    for name, digest in sorted(entries.items()):
        path = os.path.join(rel, name)
        if not os.path.exists(path):
            print('[MISS ] %s  (listed in manifest, not on disk)' % name)
            bad += 1
            continue
        actual = sha256_file(path)
        if actual != digest:
            print('[BAD  ] %s' % name)
            print('        manifest: %s' % digest)
            print('        actual  : %s' % actual)
            bad += 1
        else:
            print('[OK   ] %s  %s' % (name, digest[:16]))

    missing = sorted(on_disk - set(entries))
    for name in missing:
        print('[EXTRA] %s  (on disk, not in manifest)' % name)
        bad += 1

    bad_pkg = verify_package(pkg)
    bad += bad_pkg

    if bad:
        print('\n[FAIL] %d problem(s). Release must not ship.' % bad)
        return 1
    print('\n[OK] %d file(s) verified against SHA256.txt' % len(entries))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
