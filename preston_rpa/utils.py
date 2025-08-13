"""Utility helpers for ROI (Region of Interest) management."""
from __future__ import annotations

from typing import Tuple


def xywh_to_ltrb(xywh: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
    """Convert an ``(x, y, width, height)`` box to ``(left, top, right, bottom)``.

    Parameters
    ----------
    xywh:
        Tuple containing ``(x, y, width, height)``.

    Returns
    -------
    Tuple[int, int, int, int]
        Converted ``(left, top, right, bottom)`` tuple suitable for APIs that
        expect the LTRB format.
    """
    x, y, w, h = xywh
    return x, y, x + w, y + h
