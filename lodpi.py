"""
lodpi.py

Dpi scaling helpers for Library Organizer WinForms UI (96-DPI design baseline).
"""

import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import Control
from System.Drawing import Size

BASE_DPI = 96.0

# Baseline control metrics at 96 DPI (slightly wider than original 58px fields).
TEXTBOX_WIDTH = 80
TEXTBOX_HEIGHT = 22
SMALL_TEXTBOX_WIDTH = 72
NUMERIC_WIDTH = 40


def get_scale():
    g = Control.CreateGraphics()
    try:
        scale = g.DpiX / BASE_DPI
        if scale < 1.0:
            return 1.0
        return scale
    finally:
        g.Dispose()


def scale_int(value, scale=None):
    if scale is None:
        scale = get_scale()
    return int(round(value * scale))


def scaled_size(width, height, scale=None):
    return Size(scale_int(width, scale), scale_int(height, scale))


def textbox_size(width=TEXTBOX_WIDTH, height=TEXTBOX_HEIGHT):
    return scaled_size(width, height)


def needs_hidpi_layout():
    return get_scale() > 1.01
