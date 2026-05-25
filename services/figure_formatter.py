from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.enum.text import WD_ALIGN_PARAGRAPH

from app.models import FigureProfile, StyleProfile
from app.services.ooxml_utils import name_to_alignment, normalize_text
from app.services.table_formatter import _apply_style_to_paragraph, _profile_from_paragraph

FIGURE_CAPTION_RE = re.compile(
    r"^(图\s*\d+(?:[-.－—]\d+)?\s*\S+|图\s*[一二三四五六七八九十]+\s*\S+|Figure\s+\d+(?:[-.]\d+)?\s*\S+|Fig\.\s*\d+(?:[-.]\d+)?\s*\S+)",
    re.I,
)


@dataclass(slots=True)
class FigureFormatResult:
    figures_detected: int = 0
    figures_centered: int = 0
    figure_captions_detected: int = 0
    figure_captions_formatted: int = 0
    uncertain_links: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def is_figure_caption_text(text: str) -> bool:
    return bool(FIGURE_CAPTION_RE.match(normalize_text(text)))


def _paragraph_has_image(paragraph: object) -> bool:
    for node in paragraph._p.iter():
        if node.tag.endswith("}drawing") or node.tag.endswith("}pict"):
            return True
    return False


def find_figures(document: object) -> list[object]:
    return [paragraph for paragraph in document.paragraphs if _paragraph_has_image(paragraph)]


def find_figure_captions(document: object) -> list[object]:
    return [paragraph for paragraph in document.paragraphs if is_figure_caption_text(paragraph.text)]


def detect_figure_caption_style_from_template(template_doc: object) -> FigureProfile:
    profile = FigureProfile()
    captions = find_figure_captions(template_doc)
    if captions:
        profile.caption = _profile_from_paragraph(captions[0], profile.caption)
        profile.used_default = False
    figures = find_figures(template_doc)
    if figures:
        alignment = figures[0].alignment
        if alignment is not None:
            profile.image_alignment = alignment.name
        else:
            profile.image_alignment = "CENTER"
        profile.used_default = False
    return profile


def format_figure_caption(paragraph: object, style_profile: StyleProfile) -> None:
    _apply_style_to_paragraph(paragraph, style_profile)


def center_images(document: object, profile: FigureProfile | None = None) -> int:
    count = 0
    alignment = WD_ALIGN_PARAGRAPH.CENTER
    if profile is not None:
        alignment = name_to_alignment(profile.image_alignment) or WD_ALIGN_PARAGRAPH.CENTER
    for paragraph in find_figures(document):
        paragraph.alignment = alignment
        count += 1
    return count


def format_all_figures(document: object, template_profile: object) -> FigureFormatResult:
    result = FigureFormatResult()
    figure_profile: FigureProfile = template_profile.figure
    figures = find_figures(document)
    captions = find_figure_captions(document)
    result.figures_detected = len(figures)
    result.figure_captions_detected = len(captions)
    result.figures_centered = center_images(document, figure_profile)
    for paragraph in captions:
        format_figure_caption(paragraph, figure_profile.caption)
        result.figure_captions_formatted += 1
    if result.figures_detected and not result.figure_captions_detected:
        result.uncertain_links.append("检测到图片，但未识别到明确图题。")
    if result.figure_captions_detected and not result.figures_detected:
        result.uncertain_links.append("检测到图题，但未检测到内嵌图片对象。")
    if figure_profile.used_default:
        result.warnings.append("模板未识别到明确图题样式，系统使用默认中文论文图题格式。")
    return result
