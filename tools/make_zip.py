# -*- coding: utf-8 -*-
"""把发布目录打包成 ZIP，供分发与 GitHub Release 使用。

为什么放在 `发布包/` 而不是 `发布版/`：
  发布版/ 内的每个文件都必须被 SHA256.txt 覆盖（verify_release.py 会逐个校验，
  出现「目录里有、清单里没有」的文件即判定失败）。ZIP 是发布版的**容器**，
  它在发布版组装之后才生成，放进发布版/ 会造成清单与目录不一致。
  因此 ZIP 连同它自己的 .sha256 放在同级的 `发布包/`。
"""

from __future__ import annotations

import hashlib
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from core import __version__ as _VERSION      # noqa: E402  版本号单一事实源

REL = os.path.join(ROOT, '发布版')
PKG = os.path.join(ROOT, '发布包')
VERSION = 'v' + _VERSION
ZIP_NAME = '三早一晚免勤自动统计_%s_Windows发布包.zip' % VERSION


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def clean_package_dir() -> None:
    """清掉上一版遗留的发布包。

    与 make_release.py 同理：旧包残留会让「目录里有两个版本号的包」，
    校验时只认清单内的那一个，极易把旧包当成新包发出去。
    """
    if not os.path.isdir(PKG):
        os.makedirs(PKG, exist_ok=True)
        return
    for name in os.listdir(PKG):
        if name.lower().endswith(('.zip', '.sha256')):
            os.remove(os.path.join(PKG, name))
            print('清理旧发布包：%s' % name)


def main() -> int:
    if not os.path.isdir(REL):
        print('发布目录不存在，先跑 build_windows.bat')
        return 1
    clean_package_dir()
    os.makedirs(PKG, exist_ok=True)
    out = os.path.join(PKG, ZIP_NAME)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for name in sorted(os.listdir(REL)):
            path = os.path.join(REL, name)
            if not os.path.isfile(path) or name.lower().endswith('.zip'):
                continue          # 不把旧的发布包再打进新包里
            zf.write(path, arcname='三早一晚免勤自动统计/' + name)
    print('已生成：%s（%.1f MB）' % (out, os.path.getsize(out) / 1048576))
    with zipfile.ZipFile(out) as zf:
        for n in zf.namelist():
            print('   %s  (%.1f MB)' % (n, zf.getinfo(n).file_size / 1048576))
    digest = sha256_file(out)
    with open(out + '.sha256', 'w', encoding='utf-8') as fh:
        fh.write('%s  %s\n' % (digest, ZIP_NAME))
    print('\n发布包校验值：%s' % digest)
    print('已写出：%s.sha256' % os.path.basename(out))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
