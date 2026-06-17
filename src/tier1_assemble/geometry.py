from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple


@dataclass(frozen=True)
class Box:
    left: int
    top: int
    width: int
    height: int


def contain_box(slot: Box, image_size: Tuple[int, int], anchor: str, margin: int = 0) -> Box:
    work = _inset(slot, margin)
    img_w, img_h = image_size
    if img_w <= 0 or img_h <= 0 or work.width <= 0 or work.height <= 0:
        return work
    scale = min(work.width / img_w, work.height / img_h)
    width = max(1, int(img_w * scale))
    height = max(1, int(img_h * scale))
    return _anchor_box(work, width, height, anchor)


def _inset(slot: Box, margin: int) -> Box:
    margin = max(0, margin)
    width = max(1, slot.width - margin * 2)
    height = max(1, slot.height - margin * 2)
    return Box(slot.left + margin, slot.top + margin, width, height)


def _anchor_box(slot: Box, width: int, height: int, anchor: str) -> Box:
    if anchor in {"top-left", "left", "bottom-left"}:
        left = slot.left
    elif anchor in {"top-right", "right", "bottom-right"}:
        left = slot.left + slot.width - width
    else:
        left = slot.left + (slot.width - width) // 2

    if anchor in {"top-left", "top", "top-right"}:
        top = slot.top
    elif anchor in {"bottom-left", "bottom", "bottom-right"}:
        top = slot.top + slot.height - height
    else:
        top = slot.top + (slot.height - height) // 2

    return Box(left, top, width, height)
