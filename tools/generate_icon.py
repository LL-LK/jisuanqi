# -*- coding: utf-8 -*-
"""Multi-resolution Vector Bear Icon Generator.

Extracts vector drawing routines from BearPainter in app.py, renders 7 standard
resolutions (16x16, 24x24, 32x32, 48x48, 64x64, 128x128, 256x256), and packs
them into a genuine Windows ICO container at assets/bear.ico.

Supports headless CI environments via QT_QPA_PLATFORM=offscreen.
"""

from __future__ import annotations

import argparse
import os
import struct
import sys

# Ensure headless execution in CI or non-GUI environments
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

STANDARD_SIZES = (16, 24, 32, 48, 64, 128, 256)


def get_paths() -> tuple[str, str]:
    """Resolve project desktop root and default icon path."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    default_icon = os.path.join(project_dir, "assets", "bear.ico")
    return project_dir, default_icon


def render_png_frames(bear_painter_cls, sizes: tuple[int, ...] = STANDARD_SIZES) -> dict[int, bytes]:
    """Render PNG binary stream for each requested resolution using QBuffer."""
    from PySide6.QtCore import QBuffer, QIODevice

    frames: dict[int, bytes] = {}
    for size in sizes:
        if hasattr(bear_painter_cls, "pixmap"):
            pixmap = bear_painter_cls.pixmap(size)
        elif hasattr(bear_painter_cls, "create_bear_avatar"):
            pixmap = bear_painter_cls.create_bear_avatar(size)
        else:
            raise AttributeError(f"{bear_painter_cls} has neither pixmap nor create_bear_avatar")
        buffer = QBuffer()
        buffer.open(QIODevice.WriteOnly)
        success = pixmap.save(buffer, "PNG")
        if not success:
            raise RuntimeError(f"Failed to render QPixmap to PNG for size {size}x{size}")
        frames[size] = bytes(buffer.data())
    return frames


def assemble_ico(frames: dict[int, bytes]) -> bytes:
    """Pack multi-resolution PNG frames into standard Windows ICO format."""
    sorted_sizes = sorted(frames.keys())
    count = len(sorted_sizes)
    # ICONDIR: idReserved (WORD=0), idType (WORD=1 for ICO), idCount (WORD)
    header = struct.pack("<HHH", 0, 1, count)

    entries: list[bytes] = []
    # 6 bytes header + 16 bytes per entry
    offset = 6 + count * 16
    data_blobs: list[bytes] = []

    for size in sorted_sizes:
        data = frames[size]
        # In ICO spec, width/height of 256 is encoded as 0
        w = size if size < 256 else 0
        h = size if size < 256 else 0
        # ICONDIRENTRY:
        # bWidth(B), bHeight(B), bColorCount(B), bReserved(B),
        # wPlanes(H=1), wBitCount(H=32), dwBytesInRes(I), dwImageOffset(I)
        entry = struct.pack("<BBBBHHII", w, h, 0, 0, 1, 32, len(data), offset)
        entries.append(entry)
        data_blobs.append(data)
        offset += len(data)

    return header + b"".join(entries) + b"".join(data_blobs)


def generate_ico(output_path: str | None = None, force: bool = False) -> str:
    """Generate and write bear.ico file."""
    project_dir, default_icon = get_paths()
    target_path = str(output_path) if output_path else default_icon

    if os.path.exists(target_path) and not force:
        print(f"[Icon] Icon already exists: {target_path} (use --force to overwrite)")
        return target_path

    # Initialize headless Qt application
    try:
        from PySide6.QtGui import QGuiApplication
    except ImportError:
        if os.path.exists(target_path):
            print(f"[Icon] PySide6 unavailable, keeping existing icon: {target_path}")
            return target_path
        raise RuntimeError("PySide6 is required to generate bear.ico from vector graphics.")

    app = QGuiApplication.instance()
    if app is None:
        app = QGuiApplication([sys.argv[0], "-platform", "offscreen"])

    if project_dir not in sys.path:
        sys.path.insert(0, project_dir)

    try:
        from app import AttendanceIcon as icon_provider
    except ImportError:
        try:
            from app import BearPainter as icon_provider
        except ImportError as exc:
            raise ImportError(f"Could not import AttendanceIcon or BearPainter from {project_dir}/app.py") from exc

    frames = render_png_frames(icon_provider, STANDARD_SIZES)
    ico_bytes = assemble_ico(frames)

    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    with open(target_path, "wb") as f:
        f.write(ico_bytes)

    print(f"[Icon] Successfully generated Windows ICO: {target_path} ({len(ico_bytes)} bytes, {len(frames)} sizes)")
    return target_path


# Convenient alias for tests and callers
generate_bear_ico = generate_ico


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Multi-resolution Vector Bear Icon Generator")
    parser.add_argument("--output", "-o", default=None, help="Output path for .ico file (default: assets/bear.ico)")
    parser.add_argument("--force", "-f", action="store_true", help="Force regenerate even if file already exists")
    args = parser.parse_args(argv)

    try:
        generate_ico(args.output, force=args.force)
        return 0
    except Exception as exc:
        print(f"[Icon] Error generating icon: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
