"""
lodpi.py

Dpi scaling helpers for Library Organizer WinForms UI (96-DPI design baseline).
"""

import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System import IntPtr
from System.Drawing import Graphics, Size, Point

BASE_DPI = 96.0

# Baseline control metrics at 96 DPI (slightly wider than original 58px fields).
TEXTBOX_WIDTH = 80
TEXTBOX_HEIGHT = 22
SMALL_TEXTBOX_WIDTH = 72
NUMERIC_WIDTH = 40

# Two-column insert grids (configureform.py design coordinates at 96 DPI).
BASE_COL1 = 4
BASE_COL2 = 240
BASE_ROW0 = 10
BASE_ROW_STEP = 45
WIDE_ROW_WIDTH = 280
NARROW_ROW_WIDTH = 220


def get_scale(owner=None):
    """Return display scale factor (1.0 at 96 DPI). owner may be any Control with a handle."""
    g = None
    try:
        if owner is not None:
            g = owner.CreateGraphics()
        else:
            # IronPython: Control.CreateGraphics() is instance-only; desktop DC works at load time.
            g = Graphics.FromHwnd(IntPtr.Zero)
        scale = g.DpiX / BASE_DPI
        if scale < 1.0:
            return 1.0
        return scale
    finally:
        if g is not None:
            g.Dispose()


def scale_int(value, scale=None, owner=None):
    if scale is None:
        scale = get_scale(owner)
    return int(round(value * scale))


def scaled_size(width, height, scale=None, owner=None):
    return Size(scale_int(width, scale, owner), scale_int(height, scale, owner))


def textbox_size(width=TEXTBOX_WIDTH, height=TEXTBOX_HEIGHT, owner=None):
    return scaled_size(width, height, owner=owner)


def needs_hidpi_layout(owner=None):
    return get_scale(owner) > 1.01


def scale_insert_point(baseline, wide=False, owner=None):
    """Map a 96-DPI grid Location to HiDPI coordinates (two-column insert tabs)."""
    scale = get_scale(owner)
    if scale <= 1.01:
        return Point(baseline.X, baseline.Y)

    col1 = scale_int(BASE_COL1, scale, owner)
    row_step = scale_int(BASE_ROW_STEP, scale, owner)
    row0 = scale_int(BASE_ROW0, scale, owner)
    margin = scale_int(16, scale, owner)

    if baseline.X >= BASE_COL2:
        row_width = WIDE_ROW_WIDTH if wide else NARROW_ROW_WIDTH
        new_x = col1 + scale_int(row_width, scale, owner) + margin
    else:
        new_x = col1

    row = 0
    if baseline.Y > BASE_ROW0:
        row = int(round((baseline.Y - BASE_ROW0) / float(BASE_ROW_STEP)))
    new_y = row0 + row * row_step
    return Point(new_x, new_y)


def layout_row(controls, start_x, y, owner=None, gap=8, coords_scaled=False):
    """Place controls left-to-right; avoids fixed-X overlap when fonts grow at HiDPI.

    coords_scaled: when True, start_x, y, and gap are already in device pixels (for vertical flow chains).
    """
    scale = get_scale(owner)
    if scale <= 1.01:
        return
    if coords_scaled:
        x = start_x
        row_y = y
        g = gap
        nudge = max(1, int(round(2 * scale)))
    else:
        x = scale_int(start_x, scale, owner)
        row_y = scale_int(y, scale, owner)
        g = scale_int(gap, scale, owner)
        nudge = scale_int(2, scale, owner)
    for control in controls:
        type_name = control.GetType().Name
        cy = row_y
        if type_name in ("TextBox", "ComboBox", "NumericUpDown"):
            cy = row_y - nudge
        elif type_name == "Button":
            cy = row_y - nudge
        control.Location = Point(x, cy)
        x = control.Right + g

