# -*- coding: utf-8 -*-
"""生成 app.ico（打包用的程序图标）。

为什么用代码画而不是放一张图：图标要 16/24/32/48/64/128/256 这么多尺寸，
手工切图容易糊；这里用 QPainter 按尺寸重绘，每个尺寸都清晰。
Qt 的 ico 写入器一次只写一张图，所以这里自己拼 ICO 容器（PNG 压缩）。

用法：
    python make_icon.py            # 生成 app.ico
    python make_icon.py out.ico    # 指定输出名
"""
import os
import struct
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QBuffer, QByteArray, QRectF, Qt
from PySide6.QtGui import (QColor, QImage, QLinearGradient, QPainter,
                           QPainterPath, QPen)

SIZES = (16, 24, 32, 48, 64, 128, 256)

BG_TOP = QColor("#4F8DF7")
BG_BOTTOM = QColor("#1D4ED8")
CELL = QColor(255, 255, 255, 235)
CELL_DIM = QColor(255, 255, 255, 120)
HILIGHT = QColor("#FFD451")


def render(size: int) -> QImage:
    """画一枚：圆角蓝底 + 白色表格格子 + 一个黄色高亮格。"""
    img = QImage(size, size, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing, True)
    p.setRenderHint(QPainter.SmoothPixmapTransform, True)

    s = float(size)
    # 背景：圆角方块 + 竖向渐变
    grad = QLinearGradient(0, 0, 0, s)
    grad.setColorAt(0.0, BG_TOP)
    grad.setColorAt(1.0, BG_BOTTOM)
    path = QPainterPath()
    path.addRoundedRect(QRectF(0.5, 0.5, s - 1, s - 1), s * 0.22, s * 0.22)
    p.fillPath(path, grad)

    # 表格：3 列 × 3 行，留边距
    pad = s * 0.20
    gap = max(s * 0.035, 1.0)
    w = s - pad * 2
    cell = (w - gap * 2) / 3.0
    r = max(cell * 0.18, s * 0.02)
    for row in range(3):
        for col in range(3):
            x = pad + col * (cell + gap)
            y = pad + row * (cell + gap)
            box = QPainterPath()
            box.addRoundedRect(QRectF(x, y, cell, cell), r, r)
            if row == 0:
                color = CELL          # 表头一行最亮
            elif row == 1 and col == 2:
                color = HILIGHT       # 高亮格：对应「多列黄色合并」
            else:
                color = CELL_DIM
            p.fillPath(box, color)

    # 大尺寸时补一条细描边，缩小后更容易看清轮廓
    if size >= 48:
        p.setPen(QPen(QColor(255, 255, 255, 60), max(s * 0.012, 1.0)))
        p.setBrush(Qt.NoBrush)
        p.drawPath(path)
    p.end()
    return img


def png_bytes(img: QImage) -> bytes:
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QBuffer.WriteOnly)
    img.save(buf, "PNG")
    buf.close()
    return bytes(ba)


def write_ico(path: str, sizes=SIZES) -> str:
    """把多个尺寸的 PNG 拼成一个 ICO 文件。"""
    blobs = [png_bytes(render(s)) for s in sizes]
    n = len(sizes)
    header = struct.pack("<HHH", 0, 1, n)          # reserved, type=icon, count
    offset = 6 + 16 * n
    entries, data = b"", b""
    for s, blob in zip(sizes, blobs):
        dim = 0 if s >= 256 else s                  # 256 在 ICO 里写作 0
        entries += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32,
                               len(blob), offset)
        data += blob
        offset += len(blob)
    with open(path, "wb") as f:
        f.write(header + entries + data)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
    write_ico(out)
    print("saved %s (%.1f KB)" % (out, os.path.getsize(out) / 1024.0))
