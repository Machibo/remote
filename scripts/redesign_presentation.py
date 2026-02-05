from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from math import ceil
import os
import re

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


INPUT_PATH = os.path.join("input", "presentation.pptx")
OUTPUT_PATH = os.path.join("output", "redesign.pptx")


# Slide geometry (16:9)
SLIDE_W = Inches(10)
SLIDE_H = Inches(5.625)

# Layout constants
MARGIN_X = Inches(0.6)
TITLE_Y = Inches(0.25)
TITLE_H = Inches(0.7)
TITLE_LINE_Y = Inches(0.98)
TITLE_LINE_W = Inches(2.6)
TITLE_LINE_H = Inches(0.04)
CONTENT_Y = Inches(1.25)
FOOTER_H = Inches(0.5)
GAP = Inches(0.35)

# Colors
COLOR_BG = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_TEXT = RGBColor(0x0A, 0x0A, 0x0A)
COLOR_ACCENT = RGBColor(0x5E, 0xC0, 0x62)
COLOR_ACCENT_2 = RGBColor(0x00, 0x97, 0xA7)
COLOR_MUTED = RGBColor(0x6B, 0x72, 0x80)
COLOR_CALLOUT_BG = RGBColor(0xE8, 0xF4, 0xEC)
COLOR_FOOTER_BG = RGBColor(0xF1, 0xF3, 0xF6)

# Fonts
FONT_TITLE = "Montserrat"
FONT_BODY = "Inter"


@dataclass
class SlideData:
    title: str
    lead: list[str]
    bullets: list[str]
    callouts: list[str]
    cta: list[str]
    contacts: list[str]
    images: list[dict]


def _get_max_font_size(shape) -> float:
    max_size = 0.0
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size:
                max_size = max(max_size, run.font.size.pt)
    return max_size


def _classify_contact(text: str) -> str | None:
    if "Москва" in text:
        return "address"
    if text.strip().startswith("@"):
        return "handle"
    if "+7" in text:
        return "phone"
    if "@" in text and "." in text:
        return "email"
    if ".ru" in text or "asg-" in text:
        return "site"
    return None


def _is_callout(text: str) -> bool:
    stripped = text.strip()
    if stripped.startswith("Вывод:"):
        return True
    if stripped.startswith("ЭТО ") or stripped.startswith("ВЫ "):
        return True
    letters = [c for c in stripped if c.isalpha()]
    if not letters:
        return False
    upper = [c for c in letters if c.upper() == c]
    if len(upper) / len(letters) > 0.85 and len(stripped) > 20:
        return True
    if "→" in stripped and len(stripped) > 10:
        return True
    return False


def _is_cta(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    if len(stripped) > 30:
        return False
    if stripped.upper() != stripped:
        return False
    return "ЗАЯВ" in stripped or "ВСТРЕЧ" in stripped


def _split_items(text: str) -> list[str]:
    items: list[str] = []
    for chunk in text.split("|"):
        for part in chunk.split("\n"):
            item = part.strip()
            if item:
                items.append(item)
    return items


def _extract_slide_data(slide) -> SlideData:
    texts = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text.strip()
        if not text:
            continue
        texts.append(
            {
                "text": text,
                "top": shape.top,
                "left": shape.left,
                "size": _get_max_font_size(shape),
            }
        )
    texts.sort(key=lambda t: (t["top"], t["left"]))

    images = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            images.append(
                {
                    "blob": shape.image.blob,
                    "width": int(shape.width),
                    "height": int(shape.height),
                    "area": int(shape.width) * int(shape.height),
                }
            )

    title = ""
    if texts:
        texts_by_size = sorted(texts, key=lambda t: (-t["size"], t["top"], t["left"]))
        title = texts_by_size[0]["text"]

    contacts = []
    content_texts = []
    for item in texts:
        if item["text"] == title:
            continue
        if _classify_contact(item["text"]):
            contacts.append(item["text"])
        else:
            content_texts.append(item["text"])

    lead: list[str] = []
    bullets: list[str] = []
    callouts: list[str] = []
    cta: list[str] = []

    for text in content_texts:
        if _is_cta(text):
            cta.append(text)
            continue
        if _is_callout(text):
            callouts.append(text)
            continue
        if text.strip().endswith(":"):
            lead.append(text)
            continue
        if text.strip().endswith(".") and len(text) > 40 and not bullets:
            lead.append(text)
            continue
        bullets.extend(_split_items(text))

    return SlideData(
        title=title,
        lead=lead,
        bullets=bullets,
        callouts=callouts,
        cta=cta,
        contacts=contacts,
        images=images,
    )


def _add_rect(slide, left, top, width, height, fill_color, line_color=None, radius=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def _add_text_box(
    slide,
    left,
    top,
    width,
    height,
    lines,
    font_name,
    font_size,
    bold=False,
    color=COLOR_TEXT,
    align=PP_ALIGN.LEFT,
    line_spacing=1.1,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    for idx, line in enumerate(lines):
        paragraph = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        paragraph.text = line
        paragraph.font.name = font_name
        paragraph.font.size = Pt(font_size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color
        paragraph.alignment = align
        paragraph.line_spacing = line_spacing
    return box


def _apply_text_frame(tf, lines, font_name, font_size, bold=False, color=COLOR_TEXT, align=PP_ALIGN.LEFT):
    tf.clear()
    tf.word_wrap = True
    for idx, line in enumerate(lines):
        paragraph = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        paragraph.text = line
        paragraph.font.name = font_name
        paragraph.font.size = Pt(font_size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color
        paragraph.alignment = align


def _add_picture_contain(slide, image_blob, box_left, box_top, box_width, box_height):
    stream = BytesIO(image_blob)
    image = Image.open(stream)
    img_w, img_h = image.size
    stream.seek(0)
    if img_w == 0 or img_h == 0:
        return
    scale = min(box_width / img_w, box_height / img_h)
    width = int(img_w * scale)
    height = int(img_h * scale)
    left = int(box_left + (box_width - width) / 2)
    top = int(box_top + (box_height - height) / 2)
    slide.shapes.add_picture(stream, left, top, width=width, height=height)


def _add_title(slide, title: str):
    _add_rect(slide, 0, 0, SLIDE_W, Inches(0.08), COLOR_ACCENT)
    _add_text_box(
        slide,
        MARGIN_X,
        TITLE_Y,
        SLIDE_W - 2 * MARGIN_X,
        TITLE_H,
        [title],
        FONT_TITLE,
        28,
        bold=True,
        color=COLOR_TEXT,
    )
    _add_rect(slide, MARGIN_X, TITLE_LINE_Y, TITLE_LINE_W, TITLE_LINE_H, COLOR_ACCENT_2)


def _format_contacts(contacts: list[str]) -> str:
    ordered = {"address": [], "handle": [], "phone": [], "email": [], "site": []}
    for text in contacts:
        kind = _classify_contact(text)
        if not kind:
            continue
        ordered[kind].append(text)
    combined = []
    for key in ["address", "handle", "phone", "email", "site"]:
        combined.extend(ordered[key])
    if not combined:
        return ""
    return " | ".join(combined)


def _add_footer(slide, contacts: list[str]):
    text = _format_contacts(contacts)
    if not text:
        return
    _add_rect(slide, 0, SLIDE_H - FOOTER_H, SLIDE_W, FOOTER_H, COLOR_FOOTER_BG)
    _add_text_box(
        slide,
        MARGIN_X,
        SLIDE_H - FOOTER_H + Inches(0.08),
        SLIDE_W - 2 * MARGIN_X,
        FOOTER_H - Inches(0.12),
        [text],
        FONT_BODY,
        10,
        color=COLOR_MUTED,
    )


def _largest_image(images: list[dict], min_ratio: float = 0.1) -> dict | None:
    if not images:
        return None
    slide_area = int(SLIDE_W) * int(SLIDE_H)
    images_sorted = sorted(images, key=lambda img: img["area"], reverse=True)
    if images_sorted[0]["area"] >= slide_area * min_ratio:
        return images_sorted[0]
    return None


def _build_cover_slide(slide, data: SlideData):
    _add_title(slide, data.title.replace(" | ", "\n"))
    subtitle_source = data.lead if data.lead else data.callouts
    subtitle = [t.strip("-") for t in subtitle_source] if subtitle_source else []
    if subtitle:
        _add_text_box(
            slide,
            MARGIN_X,
            CONTENT_Y,
            Inches(5.2),
            Inches(1.2),
            subtitle,
            FONT_BODY,
            16,
            color=COLOR_TEXT,
        )

    logo = _largest_image(data.images, min_ratio=0.02)
    if logo:
        _add_picture_contain(
            slide,
            logo["blob"],
            Inches(6.0),
            Inches(1.5),
            Inches(3.3),
            Inches(2.4),
        )

    _add_footer(slide, data.contacts)


def _build_text_slide(slide, data: SlideData):
    _add_title(slide, data.title)
    content_h = int(SLIDE_H - CONTENT_Y - FOOTER_H - Inches(0.2))

    cursor_y = CONTENT_Y
    if data.lead:
        lead_h = Inches(0.35 * len(data.lead) + 0.1)
        _add_text_box(
            slide,
            MARGIN_X,
            cursor_y,
            SLIDE_W - 2 * MARGIN_X,
            lead_h,
            data.lead,
            FONT_BODY,
            16,
            color=COLOR_TEXT,
        )
        cursor_y += lead_h + Inches(0.1)

    callout_h = Inches(0.7) if data.callouts else Inches(0)
    bullets_h = content_h - int(callout_h) - int(cursor_y - CONTENT_Y)
    bullets = data.bullets

    if bullets:
        if len(bullets) > 6:
            mid = ceil(len(bullets) / 2)
            left_items = [f"• {b}" for b in bullets[:mid]]
            right_items = [f"• {b}" for b in bullets[mid:]]
            left_w = (SLIDE_W - 2 * MARGIN_X - GAP) / 2
            right_x = MARGIN_X + left_w + GAP
            _add_text_box(
                slide,
                MARGIN_X,
                cursor_y,
                left_w,
                bullets_h,
                left_items,
                FONT_BODY,
                16,
                color=COLOR_TEXT,
            )
            _add_text_box(
                slide,
                right_x,
                cursor_y,
                left_w,
                bullets_h,
                right_items,
                FONT_BODY,
                16,
                color=COLOR_TEXT,
            )
        else:
            _add_text_box(
                slide,
                MARGIN_X,
                cursor_y,
                SLIDE_W - 2 * MARGIN_X,
                bullets_h,
                [f"• {b}" for b in bullets],
                FONT_BODY,
                16,
                color=COLOR_TEXT,
            )

    if data.callouts:
        callout = _add_rect(
            slide,
            MARGIN_X,
            SLIDE_H - FOOTER_H - Inches(0.9),
            SLIDE_W - 2 * MARGIN_X,
            Inches(0.7),
            COLOR_CALLOUT_BG,
            line_color=COLOR_ACCENT,
            radius=True,
        )
        _apply_text_frame(
            callout.text_frame,
            data.callouts,
            FONT_TITLE,
            16,
            bold=True,
            color=COLOR_ACCENT,
        )

    _add_footer(slide, data.contacts)


def _build_split_slide(slide, data: SlideData, image: dict | None):
    _add_title(slide, data.title)
    content_h = SLIDE_H - CONTENT_Y - FOOTER_H - Inches(0.2)
    left_w = Inches(5.6)
    right_w = SLIDE_W - 2 * MARGIN_X - left_w - GAP
    right_x = MARGIN_X + left_w + GAP

    cursor_y = CONTENT_Y
    if data.lead:
        lead_h = Inches(0.35 * len(data.lead) + 0.1)
        _add_text_box(
            slide,
            MARGIN_X,
            cursor_y,
            left_w,
            lead_h,
            data.lead,
            FONT_BODY,
            16,
            color=COLOR_TEXT,
        )
        cursor_y += lead_h + Inches(0.1)

    if data.bullets:
        _add_text_box(
            slide,
            MARGIN_X,
            cursor_y,
            left_w,
            content_h - (cursor_y - CONTENT_Y),
            [f"• {b}" for b in data.bullets],
            FONT_BODY,
            16,
            color=COLOR_TEXT,
        )

    if data.callouts:
        callout = _add_rect(
            slide,
            MARGIN_X,
            SLIDE_H - FOOTER_H - Inches(0.9),
            left_w,
            Inches(0.7),
            COLOR_CALLOUT_BG,
            line_color=COLOR_ACCENT,
            radius=True,
        )
        _apply_text_frame(
            callout.text_frame,
            data.callouts,
            FONT_TITLE,
            16,
            bold=True,
            color=COLOR_ACCENT,
        )

    if image:
        _add_picture_contain(
            slide,
            image["blob"],
            right_x,
            CONTENT_Y,
            right_w,
            content_h,
        )

    _add_footer(slide, data.contacts)


def _build_image_slide(slide, data: SlideData, image: dict | None):
    _add_title(slide, data.title)
    content_h = SLIDE_H - CONTENT_Y - FOOTER_H - Inches(0.2)
    cursor_y = CONTENT_Y
    if data.lead:
        lead_h = Inches(0.35 * len(data.lead) + 0.1)
        _add_text_box(
            slide,
            MARGIN_X,
            cursor_y,
            SLIDE_W - 2 * MARGIN_X,
            lead_h,
            data.lead,
            FONT_BODY,
            16,
            color=COLOR_TEXT,
        )
        cursor_y += lead_h + Inches(0.1)

    if data.callouts:
        callout = _add_rect(
            slide,
            MARGIN_X,
            cursor_y,
            SLIDE_W - 2 * MARGIN_X,
            Inches(0.7),
            COLOR_CALLOUT_BG,
            line_color=COLOR_ACCENT,
            radius=True,
        )
        _apply_text_frame(
            callout.text_frame,
            data.callouts,
            FONT_TITLE,
            16,
            bold=True,
            color=COLOR_ACCENT,
        )
        cursor_y += Inches(0.8)

    if image:
        _add_picture_contain(
            slide,
            image["blob"],
            MARGIN_X,
            cursor_y,
            SLIDE_W - 2 * MARGIN_X,
            content_h - (cursor_y - CONTENT_Y),
        )

    _add_footer(slide, data.contacts)


def _build_cta_slide(slide, data: SlideData):
    _add_title(slide, data.title)
    content_h = SLIDE_H - CONTENT_Y - FOOTER_H - Inches(0.2)
    cursor_y = CONTENT_Y

    if data.lead:
        lead_h = Inches(0.35 * len(data.lead) + 0.1)
        _add_text_box(
            slide,
            MARGIN_X,
            cursor_y,
            SLIDE_W - 2 * MARGIN_X,
            lead_h,
            data.lead,
            FONT_BODY,
            16,
            color=COLOR_TEXT,
        )
        cursor_y += lead_h + Inches(0.1)

    if data.bullets:
        _add_text_box(
            slide,
            MARGIN_X,
            cursor_y,
            SLIDE_W - 2 * MARGIN_X,
            content_h - Inches(1.5),
            [f"• {b}" for b in data.bullets],
            FONT_BODY,
            16,
            color=COLOR_TEXT,
        )
        cursor_y += Inches(1.3)

    if data.cta:
        cta_w = Inches(3.2)
        cta_h = Inches(0.6)
        start_x = MARGIN_X
        for idx, text in enumerate(data.cta[:2]):
            rect = _add_rect(
                slide,
                start_x + idx * (cta_w + Inches(0.3)),
                SLIDE_H - FOOTER_H - Inches(1.0),
                cta_w,
                cta_h,
                COLOR_ACCENT,
                line_color=COLOR_ACCENT,
                radius=True,
            )
            tf = rect.text_frame
            tf.text = text
            tf.paragraphs[0].font.name = FONT_TITLE
            tf.paragraphs[0].font.size = Pt(14)
            tf.paragraphs[0].font.bold = True
            tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    _add_footer(slide, data.contacts)


def main():
    prs = Presentation(INPUT_PATH)
    new_prs = Presentation()
    new_prs.slide_width = SLIDE_W
    new_prs.slide_height = SLIDE_H
    blank = new_prs.slide_layouts[6]

    for index, slide in enumerate(prs.slides):
        data = _extract_slide_data(slide)
        new_slide = new_prs.slides.add_slide(blank)

        if index == 0:
            _build_cover_slide(new_slide, data)
            continue

        if data.cta:
            _build_cta_slide(new_slide, data)
            continue

        large_image = _largest_image(data.images, min_ratio=0.12)

        if large_image and len(data.bullets) <= 2 and len(data.lead) <= 1:
            _build_image_slide(new_slide, data, large_image)
        elif large_image:
            _build_split_slide(new_slide, data, large_image)
        else:
            _build_text_slide(new_slide, data)

    new_prs.save(OUTPUT_PATH)


if __name__ == "__main__":
    main()
