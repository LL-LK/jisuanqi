# -*- coding: utf-8 -*-
"""动态高清晰度矢量图标、草莓小熊头像与 ActionIcon 资产引擎。"""

from __future__ import annotations

import os
import sys
from functools import lru_cache

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (
    QBrush,
    QColor,
    QGuiApplication,
    QIcon,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
)
from PySide6.QtWidgets import QApplication

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class AttendanceIcon:
    """High-DPI retina sharpness asset provider with fallback vector engine."""

    @staticmethod
    def pixmap(size=48, target_widget=None) -> QPixmap:
        if isinstance(target_widget, (int, float)):
            dpr = float(target_widget)
        elif target_widget is not None and hasattr(target_widget, 'devicePixelRatio'):
            try:
                dpr = float(target_widget.devicePixelRatio())
            except Exception:
                dpr = 1.0
        else:
            screen = QGuiApplication.primaryScreen()
            dpr = float(screen.devicePixelRatio()) if screen else 1.0
        if dpr <= 0:
            dpr = 1.0
        return AttendanceIcon._render_cached(int(size), float(dpr))

    @staticmethod
    @lru_cache(maxsize=32)
    def _render_cached(size: int, dpr: float) -> QPixmap:
        phys_size = max(1, int(round(size * dpr)))
        image_path = os.path.join(getattr(sys, '_MEIPASS', PROJECT_DIR), 'assets', 'strawberry-bear.png')
        artwork = QPixmap(image_path)
        if not artwork.isNull():
            scaled = artwork.scaled(phys_size, phys_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            scaled.setDevicePixelRatio(dpr)
            return scaled
        icon = QPixmap(phys_size, phys_size)
        icon.fill(Qt.transparent)
        painter = QPainter(icon)
        painter.setRenderHint(QPainter.Antialiasing)
        scale_factor = phys_size / 48.0
        painter.scale(scale_factor, scale_factor)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor('#F8E7EC'))
        painter.drawRoundedRect(QRectF(1, 1, 46, 46), 10, 10)
        pen = QPen(QColor('#781236'), 2.6)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(QRectF(11, 13, 26, 25), 4, 4)
        painter.drawLine(QPointF(11, 21), QPointF(37, 21))
        painter.drawLine(QPointF(18, 10), QPointF(18, 16))
        painter.drawLine(QPointF(30, 10), QPointF(30, 16))
        painter.drawLine(QPointF(18, 29), QPointF(22, 33))
        painter.drawLine(QPointF(22, 33), QPointF(30, 26))
        painter.end()
        icon.setDevicePixelRatio(dpr)
        return icon

    @staticmethod
    def icon(size=48, target_widget=None) -> QIcon:
        ico = QIcon()
        for s in (16, 24, 32, 48, 64, 128, 256):
            ico.addPixmap(AttendanceIcon.pixmap(s, target_widget=target_widget))
        return ico

    # Backward compatibility for icon generators and legacy tools
    create_bear_avatar = pixmap


AttendanceIcon.pixmap.cache_clear = AttendanceIcon._render_cached.cache_clear
AttendanceIcon.pixmap.cache_info = AttendanceIcon._render_cached.cache_info
AttendanceIcon.create_bear_avatar.cache_clear = AttendanceIcon._render_cached.cache_clear
AttendanceIcon.create_bear_avatar.cache_info = AttendanceIcon._render_cached.cache_info

# Backward compatibility alias
BearPainter = AttendanceIcon


class ActionIcon:
    """Multi-resolution vector action icon provider for QPushButton and UI actions."""

    _CACHE: dict[tuple[str, str], QIcon] = {}

    @classmethod
    def get(cls, kind: str, color: str = '#781236') -> QIcon:
        key = (kind, color.lower())
        if key in cls._CACHE:
            return cls._CACHE[key]
        ico = QIcon()
        for size in (16, 20, 24, 32, 48):
            px = cls._render_pixmap(kind, size, color)
            ico.addPixmap(px)
        cls._CACHE[key] = ico
        return ico

    @staticmethod
    def _render_pixmap(kind: str, size: int, color_hex: str) -> QPixmap:
        dpr = 1.0
        app = QApplication.instance()
        if app:
            try:
                dpr = float(app.devicePixelRatio())
            except Exception:
                dpr = 1.0
        if dpr <= 0:
            dpr = 1.0
        phys_size = max(1, int(round(size * dpr)))
        px = QPixmap(phys_size, phys_size)
        px.fill(Qt.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        scale_factor = phys_size / 24.0
        p.scale(scale_factor, scale_factor)
        color = QColor(color_hex)
        pen = QPen(color, 2.0, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)

        if kind in ('document', 'file'):
            path = QPainterPath()
            path.moveTo(5, 3)
            path.lineTo(14, 3)
            path.lineTo(19, 8)
            path.lineTo(19, 21)
            path.lineTo(5, 21)
            path.closeSubpath()
            p.drawPath(path)
            p.drawLine(QPointF(14, 3), QPointF(14, 8))
            p.drawLine(QPointF(14, 8), QPointF(19, 8))
            p.drawLine(QPointF(8, 12), QPointF(16, 12))
            p.drawLine(QPointF(8, 16), QPointF(13, 16))
        elif kind == 'table':
            p.drawRoundedRect(QRectF(3.5, 3.5, 17, 17), 2.0, 2.0)
            p.drawLine(QPointF(3.5, 8.5), QPointF(20.5, 8.5))
            p.drawLine(QPointF(9.5, 8.5), QPointF(9.5, 20.5))
            p.drawLine(QPointF(3.5, 14.5), QPointF(20.5, 14.5))
        elif kind in ('play', 'compute', 'calculate'):
            path = QPainterPath()
            path.moveTo(8.5, 6.0)
            path.lineTo(18.0, 12.0)
            path.lineTo(8.5, 18.0)
            path.closeSubpath()
            p.setBrush(QBrush(color))
            p.drawPath(path)
        elif kind in ('sparkle', 'magic'):
            path = QPainterPath()
            path.moveTo(12, 3)
            path.quadTo(12, 12, 21, 12)
            path.quadTo(12, 12, 12, 21)
            path.quadTo(12, 12, 3, 12)
            path.quadTo(12, 12, 12, 3)
            path.closeSubpath()
            p.setBrush(QBrush(color))
            p.drawPath(path)
            p.drawEllipse(QRectF(17, 4, 2.5, 2.5))
        elif kind in ('export', 'save_as', 'upload', 'save'):
            tray = QPainterPath()
            tray.moveTo(4, 14)
            tray.lineTo(4, 19)
            tray.quadTo(4, 20.5, 5.5, 20.5)
            tray.lineTo(18.5, 20.5)
            tray.quadTo(20, 20.5, 20, 19)
            tray.lineTo(20, 14)
            p.drawPath(tray)
            p.drawLine(QPointF(12, 15), QPointF(12, 4.5))
            head = QPainterPath()
            head.moveTo(7.5, 9)
            head.lineTo(12, 4.5)
            head.lineTo(16.5, 9)
            p.drawPath(head)
        elif kind in ('folder', 'folder_open', 'dir'):
            back = QPainterPath()
            back.moveTo(3, 7)
            back.lineTo(8, 7)
            back.lineTo(10, 9)
            back.lineTo(20, 9)
            back.quadTo(21, 9, 21, 10)
            back.lineTo(21, 18)
            back.lineTo(3, 18)
            back.closeSubpath()
            p.drawPath(back)

            flap = QPainterPath()
            flap.moveTo(2.5, 19)
            flap.lineTo(5.5, 12)
            flap.lineTo(21.5, 12)
            flap.lineTo(18.5, 19)
            flap.closeSubpath()
            p.setBrush(QColor('#FFF8FA'))
            p.drawPath(flap)
            p.setBrush(Qt.NoBrush)
        elif kind in ('tools', 'settings', 'gear'):
            for i in range(8):
                angle = i * 45.0
                p.save()
                p.translate(12, 12)
                p.rotate(angle)
                p.drawRoundedRect(QRectF(-1.8, -9.5, 3.6, 3.2), 0.8, 0.8)
                p.restore()
            p.drawEllipse(QRectF(4.5, 4.5, 15.0, 15.0))
            p.drawEllipse(QRectF(9.0, 9.0, 6.0, 6.0))
        elif kind in ('search', 'magnifier'):
            p.drawEllipse(QRectF(5.5, 5.5, 9.0, 9.0))
            p.drawLine(QPointF(13.18, 13.18), QPointF(19.0, 19.0))
        elif kind in ('save_disk', 'disk'):
            path = QPainterPath()
            path.moveTo(4, 4)
            path.lineTo(17, 4)
            path.lineTo(20, 7)
            path.lineTo(20, 20)
            path.lineTo(4, 20)
            path.closeSubpath()
            p.drawPath(path)
            p.drawRect(QRectF(7, 4, 10, 5))
            p.drawRect(QRectF(6, 12, 12, 6))
        elif kind in ('refresh', 'reset'):
            p.drawArc(QRectF(5.5, 5.5, 13.0, 13.0), 45 * 16, 270 * 16)
            head = QPainterPath()
            head.moveTo(13.0, 6.0)
            head.lineTo(17.5, 6.5)
            head.lineTo(16.5, 11.0)
            p.drawPath(head)
        elif kind in ('loading', 'spinner'):
            for i in range(8):
                angle = i * 45
                p.save()
                p.translate(12, 12)
                p.rotate(angle)
                dot_color = QColor(color)
                dot_color.setAlphaF(max(0.18, (i + 1) / 8.0))
                p.setPen(QPen(dot_color, 2.2, Qt.SolidLine, Qt.RoundCap))
                p.drawLine(QPointF(0, -6), QPointF(0, -9))
                p.restore()
        elif kind in ('dashboard', 'overview'):
            p.drawRoundedRect(QRectF(3.5, 3.5, 7.0, 7.0), 2, 2)
            p.drawRoundedRect(QRectF(13.5, 3.5, 7.0, 5.0), 2, 2)
            p.drawRoundedRect(QRectF(13.5, 11.5, 7.0, 9.0), 2, 2)
            p.drawRoundedRect(QRectF(3.5, 13.5, 7.0, 7.0), 2, 2)
        elif kind in ('percent', 'rate'):
            p.setBrush(QBrush(color))
            p.drawEllipse(QPointF(7.5, 7.5), 2.5, 2.5)
            p.drawEllipse(QPointF(16.5, 16.5), 2.5, 2.5)
            p.setBrush(Qt.NoBrush)
            p.drawLine(QPointF(17.5, 6.5), QPointF(6.5, 17.5))
        elif kind in ('dictionary', 'book'):
            left = QPainterPath()
            left.moveTo(11.5, 5)
            left.quadTo(8.0, 4.0, 4.0, 5.0)
            left.lineTo(4.0, 19.0)
            left.quadTo(8.0, 18.0, 11.5, 19.0)
            left.closeSubpath()
            p.drawPath(left)
            right = QPainterPath()
            right.moveTo(12.5, 5)
            right.quadTo(16.0, 4.0, 20.0, 5.0)
            right.lineTo(20.0, 19.0)
            right.quadTo(16.0, 18.0, 12.5, 19.0)
            right.closeSubpath()
            p.drawPath(right)
        elif kind in ('database', 'db', 'storage'):
            p.drawEllipse(QRectF(4, 4, 16, 5))
            p.drawArc(QRectF(4, 9, 16, 5), 180 * 16, 180 * 16)
            p.drawArc(QRectF(4, 14, 16, 5), 180 * 16, 180 * 16)
            p.drawLine(QPointF(4, 6.5), QPointF(4, 16.5))
            p.drawLine(QPointF(20, 6.5), QPointF(20, 16.5))
        elif kind in ('chart', 'chart_bar'):
            p.drawLine(QPointF(3, 20), QPointF(21, 20))
            p.drawLine(QPointF(3, 4), QPointF(3, 20))
            p.drawRect(QRectF(5.5, 12, 3.5, 8))
            p.drawRect(QRectF(10.5, 8, 3.5, 12))
            p.drawRect(QRectF(15.5, 5, 3.5, 15))
        elif kind in ('check_circle', 'check', 'success'):
            p.drawEllipse(QRectF(3.5, 3.5, 17.0, 17.0))
            chk = QPainterPath()
            chk.moveTo(7.5, 12.0)
            chk.lineTo(10.5, 15.0)
            chk.lineTo(16.5, 9.0)
            p.drawPath(chk)
        elif kind in ('alert_circle', 'alert', 'warn', 'warning'):
            tri = QPainterPath()
            tri.moveTo(12, 3.5)
            tri.lineTo(20.5, 18.5)
            tri.quadTo(21.0, 19.5, 20.0, 20.0)
            tri.lineTo(4.0, 20.0)
            tri.quadTo(3.0, 19.5, 3.5, 18.5)
            tri.closeSubpath()
            p.drawPath(tri)
            p.drawLine(QPointF(12, 8.5), QPointF(12, 13.5))
            p.drawEllipse(QRectF(11.2, 15.5, 1.6, 1.6))
        elif kind in ('info_circle', 'info'):
            p.drawEllipse(QRectF(3.5, 3.5, 17.0, 17.0))
            p.drawEllipse(QRectF(11.2, 7.5, 1.6, 1.6))
            p.drawLine(QPointF(12, 11.0), QPointF(12, 16.5))
        elif kind in ('lightbulb', 'idea', 'tip'):
            bulb = QPainterPath()
            bulb.moveTo(9, 16)
            bulb.lineTo(9, 15)
            bulb.quadTo(5.5, 12, 5.5, 8.5)
            bulb.arcTo(QRectF(5.5, 3.5, 13, 13), 180, -180)
            bulb.quadTo(18.5, 12, 15, 15)
            bulb.lineTo(15, 16)
            bulb.closeSubpath()
            p.drawPath(bulb)
            p.drawLine(QPointF(9, 18.5), QPointF(15, 18.5))
            p.drawLine(QPointF(10.5, 20.5), QPointF(13.5, 20.5))
        elif kind in ('lightning', 'flash', 'quick'):
            bolt = QPainterPath()
            bolt.moveTo(13, 3)
            bolt.lineTo(6, 13)
            bolt.lineTo(12, 13)
            bolt.lineTo(11, 21)
            bolt.lineTo(18, 11)
            bolt.lineTo(12, 11)
            bolt.closeSubpath()
            p.setBrush(QBrush(color))
            p.drawPath(bolt)
        p.end()
        px.setDevicePixelRatio(dpr)
        return px
