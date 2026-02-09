from __future__ import annotations

import os
from datetime import datetime
from typing import List, Tuple

from PIL import Image, ImageDraw, ImageFont


def _load_font(paths: List[str], size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in paths:
        if os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def _text_size(font: ImageFont.ImageFont, text: str) -> Tuple[int, int]:
    if hasattr(font, "getbbox"):
        box = font.getbbox(text)
        return box[2] - box[0], box[3] - box[1]
    return font.getsize(text)


def _wrap_text(text: str, font: ImageFont.ImageFont, max_width: int) -> List[str]:
    words = text.split()
    lines: List[str] = []
    current: List[str] = []
    for word in words:
        current.append(word)
        w, _ = _text_size(font, " ".join(current))
        if w > max_width:
            current.pop()
            if current:
                lines.append(" ".join(current))
            current = [word]
    if current:
        lines.append(" ".join(current))
    return lines


def _fit_text(
    text: str,
    font_paths: List[str],
    max_width: int,
    max_height: int,
    start_size: int,
    min_size: int = 26,
) -> Tuple[ImageFont.ImageFont, List[str]]:
    size = start_size
    while size >= min_size:
        font = _load_font(font_paths, size)
        lines = _wrap_text(text, font, max_width)
        line_height = _text_size(font, "Ay")[1] + int(size * 0.2)
        if line_height * len(lines) <= max_height and len(lines) <= 4:
            return font, lines
        size -= 2
    font = _load_font(font_paths, min_size)
    return font, _wrap_text(text, font, max_width)


def generate_image(
    title: str,
    template_path: str,
    output_dir: str,
    size: Tuple[int, int],
    title_box: Tuple[int, int, int, int],
    title_color: str,
    title_font_size: int,
    title_font_paths: List[str],
    stroke_width: int,
    stroke_fill: str,
    background_color: str,
    uppercase: bool = True,
) -> str:
    if os.path.exists(template_path):
        base = Image.open(template_path).convert("RGBA")
    else:
        base = Image.new("RGBA", size, background_color)

    draw = ImageDraw.Draw(base)
    text = title.upper() if uppercase else title
    x, y, w, h = title_box
    font, lines = _fit_text(text, title_font_paths, w, h, title_font_size)
    line_height = _text_size(font, "Ay")[1] + int(font.size * 0.2)
    total_height = line_height * len(lines)
    y_offset = y + max(0, (h - total_height) // 2)
    for line in lines:
        line_width = _text_size(font, line)[0]
        x_offset = x + max(0, (w - line_width) // 2)
        draw.text(
            (x_offset, y_offset),
            line,
            font=font,
            fill=title_color,
            stroke_width=stroke_width,
            stroke_fill=stroke_fill,
        )
        y_offset += line_height

    os.makedirs(output_dir, exist_ok=True)
    filename = f"post_{datetime.utcnow().strftime('%Y%m%d_%H%M%S_%f')}.png"
    output_path = os.path.join(output_dir, filename)
    base.convert("RGB").save(output_path, format="PNG")
    return output_path
