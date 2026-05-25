from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Protocol

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from app.models import ParagraphPrediction
from app.services.ooxml_utils import normalize_text

CHINESE_NUM = "一二三四五六七八九十百千万零〇两"

SPECIAL_PATTERNS: list[tuple[str, str, float]] = [
    (r"^摘\s*要$", "abstract_title", 0.96),
    (r"^abstract$", "abstract_title", 0.96),
    (r"^(关键词|关键字)\s*[:：]?", "keywords", 0.94),
    (r"^(keywords?|key\s+words?)\s*[:：]?", "keywords", 0.94),
    (r"^目\s*录$", "contents_title", 0.97),
    (r"^(参考文献|references)$", "references_title", 0.96),
    (r"^(致谢|acknowledg(e)?ments?)$", "acknowledgements_title", 0.95),
    (r"^(附录|appendix)(\s*[A-Z0-9一二三四五六七八九十]+)?$", "appendix_title", 0.95),
]

HEADING_PATTERNS: list[tuple[str, str, float, str]] = [
    (rf"^第[{CHINESE_NUM}0-9]+章(\s+|[:：])?.+", "heading_1", 0.96, "chapter"),
    (rf"^[{CHINESE_NUM}]+、\s*\S+", "heading_1", 0.9, "chinese-list"),
    (r"^Chapter\s+[0-9]+(\s+|[:：])+\S+", "heading_1", 0.94, "chapter-en"),
    (r"^[1-9][0-9]?\s+[^.。；;，,]{1,40}$", "heading_1", 0.78, "numeric-title"),
    (r"^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+(\s+|[:：])+\S+", "heading_4", 0.92, "decimal-4"),
    (r"^[0-9]+\.[0-9]+\.[0-9]+(\s+|[:：])+\S+", "heading_3", 0.94, "decimal-3"),
    (r"^（[0-9]+）\s*\S+", "heading_3", 0.83, "paren-digit-cn"),
    (r"^\([0-9]+\)\s*\S+", "heading_3", 0.83, "paren-digit"),
    (r"^[0-9]+\.[0-9]+(\s+|[:：])+\S+", "heading_2", 0.92, "decimal-2"),
    (rf"^（[{CHINESE_NUM}]+）\s*\S+", "heading_2", 0.84, "paren-chinese-cn"),
    (rf"^\([{CHINESE_NUM}]+\)\s*\S+", "heading_2", 0.84, "paren-chinese"),
]


class AIParagraphClassifier(Protocol):
    def __call__(self, text: str, features: dict[str, object]) -> ParagraphPrediction | None:
        """Return a model-assisted prediction, or None to keep rule output."""


@dataclass(slots=True)
class ParagraphFeatures:
    text: str
    style_name: str
    is_centered: bool
    has_bold_run: bool
    max_font_size_pt: float | None
    length: int

    def to_dict(self) -> dict[str, object]:
        return {
            "text": self.text,
            "style_name": self.style_name,
            "is_centered": self.is_centered,
            "has_bold_run": self.has_bold_run,
            "max_font_size_pt": self.max_font_size_pt,
            "length": self.length,
        }


def _max_font_size(paragraph: object) -> float | None:
    sizes: list[float] = []
    for run in paragraph.runs:
        if run.font.size is not None:
            sizes.append(float(run.font.size.pt))
    return max(sizes) if sizes else None


def _has_bold_run(paragraph: object) -> bool:
    bold_runs = [run for run in paragraph.runs if run.text.strip() and run.bold]
    return bool(bold_runs)


def extract_features(paragraph: object) -> ParagraphFeatures:
    text = normalize_text(paragraph.text)
    style_name = getattr(getattr(paragraph, "style", None), "name", "") or ""
    alignment = paragraph.alignment
    is_centered = alignment == WD_ALIGN_PARAGRAPH.CENTER
    return ParagraphFeatures(
        text=text,
        style_name=style_name,
        is_centered=is_centered,
        has_bold_run=_has_bold_run(paragraph),
        max_font_size_pt=_max_font_size(paragraph),
        length=len(text),
    )


def _looks_like_sentence(text: str) -> bool:
    if len(text) > 70:
        return True
    punctuation_count = len(re.findall(r"[，,。；;：:]", text))
    return punctuation_count >= 2 or text.endswith(("。", ".", "；", ";"))


def _clamp_confidence(value: float) -> float:
    return round(max(0.0, min(0.99, value)), 2)


def classify_text(
    text: str,
    *,
    style_name: str = "",
    is_centered: bool = False,
    has_bold_run: bool = False,
    max_font_size_pt: float | None = None,
    index: int = 0,
) -> ParagraphPrediction:
    normalized = normalize_text(text)
    if not normalized:
        return ParagraphPrediction(index=index, text="", predicted_type="body", confidence=0.6, reasons=["empty"])

    for pattern, predicted_type, confidence in SPECIAL_PATTERNS:
        if re.match(pattern, normalized, re.I):
            return ParagraphPrediction(
                index=index,
                text=normalized,
                predicted_type=predicted_type,  # type: ignore[arg-type]
                confidence=confidence,
                reasons=["special-section"],
            )

    style_lower = style_name.lower()
    if style_lower.startswith("toc") or _looks_like_toc_entry(normalized):
        return ParagraphPrediction(index, normalized, "contents_entry", 0.98, ["toc-entry"])

    for pattern, predicted_type, base_confidence, reason in HEADING_PATTERNS:
        if not re.match(pattern, normalized, re.I):
            continue
        if reason in {"paren-digit", "paren-digit-cn", "paren-chinese", "paren-chinese-cn"}:
            if not has_bold_run and not is_centered:
                return ParagraphPrediction(index, normalized, "body", 0.8, ["list-marker-body"])
        confidence = base_confidence
        reasons = [reason]
        if has_bold_run:
            confidence += 0.04
            reasons.append("bold")
        if is_centered:
            confidence += 0.03
            reasons.append("centered")
        if max_font_size_pt and max_font_size_pt >= 14:
            confidence += 0.03
            reasons.append("large-font")
        if _looks_like_sentence(normalized):
            confidence -= 0.18
            reasons.append("sentence-like")
        if len(normalized) > 45 and reason in {"numeric-title", "paren-digit", "paren-digit-cn"}:
            confidence -= 0.12
            reasons.append("long-title-risk")
        return ParagraphPrediction(
            index=index,
            text=normalized,
            predicted_type=predicted_type,  # type: ignore[arg-type]
            confidence=_clamp_confidence(confidence),
            reasons=reasons,
        )

    if "heading 1" in style_lower or style_name in {"标题 1", "标题1"}:
        return ParagraphPrediction(index, normalized, "heading_1", 0.9, ["word-style"])
    if "heading 2" in style_lower or style_name in {"标题 2", "标题2"}:
        return ParagraphPrediction(index, normalized, "heading_2", 0.88, ["word-style"])
    if "heading 3" in style_lower or style_name in {"标题 3", "标题3"}:
        return ParagraphPrediction(index, normalized, "heading_3", 0.86, ["word-style"])

    body_confidence = 0.88
    if has_bold_run and len(normalized) <= 35:
        body_confidence = 0.74
    return ParagraphPrediction(index, normalized, "body", body_confidence, ["default-body"])


def _looks_like_toc_entry(text: str) -> bool:
    if "\t" not in text:
        return False
    return bool(re.search(r"\t\s*(\d+|[ivxlcdm]+|n)\s*$", text, re.I))


class StudentAnalyzer:
    def __init__(self, ai_classifier: AIParagraphClassifier | None = None) -> None:
        self.ai_classifier = ai_classifier

    def classify_paragraph(self, paragraph: object, index: int) -> ParagraphPrediction:
        features = extract_features(paragraph)
        prediction = classify_text(
            features.text,
            style_name=features.style_name,
            is_centered=features.is_centered,
            has_bold_run=features.has_bold_run,
            max_font_size_pt=features.max_font_size_pt,
            index=index,
        )
        if self.ai_classifier and prediction.confidence < 0.75:
            ai_prediction = self.ai_classifier(features.text, features.to_dict())
            if ai_prediction is not None:
                return ai_prediction
        return prediction

    def analyze(self, docx_path: str | Path) -> list[ParagraphPrediction]:
        document = Document(str(docx_path))
        return [
            self.classify_paragraph(paragraph, index)
            for index, paragraph in enumerate(document.paragraphs)
        ]


def analyze_student_document(docx_path: str | Path) -> list[ParagraphPrediction]:
    return StudentAnalyzer().analyze(docx_path)
