from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Literal

ParagraphType = Literal[
    "body",
    "heading_1",
    "heading_2",
    "heading_3",
    "heading_4",
    "abstract_title",
    "keywords",
    "contents_title",
    "contents_entry",
    "references_title",
    "acknowledgements_title",
    "appendix_title",
]


@dataclass(slots=True)
class StyleProfile:
    font_name: str | None = None
    latin_font_name: str | None = None
    font_size_pt: float | None = None
    bold: bool | None = None
    line_spacing: float | None = None
    first_line_indent: float | None = None
    alignment: str | None = None
    space_before: float | None = None
    space_after: float | None = None
    outline_level: int | None = None
    numbering_pattern: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class PageProfile:
    margin_top: float | None = None
    margin_bottom: float | None = None
    margin_left: float | None = None
    margin_right: float | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class HeaderFooterProfile:
    header_text: str | None = None
    footer_text: str | None = None
    has_page_number: bool | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class TableBorderProfile:
    top_size: int = 12
    header_bottom_size: int = 8
    bottom_size: int = 12
    color: str = "000000"
    vertical: bool = False
    inside_horizontal: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class TableProfile:
    is_three_line: bool = True
    caption: StyleProfile = field(
        default_factory=lambda: StyleProfile(
            font_name="黑体",
            latin_font_name="Times New Roman",
            font_size_pt=11,
            bold=False,
            alignment="CENTER",
            space_before=12,
            space_after=6,
        )
    )
    body: StyleProfile = field(
        default_factory=lambda: StyleProfile(
            font_name="宋体",
            latin_font_name="Times New Roman",
            font_size_pt=11,
            alignment="CENTER",
            line_spacing=1.0,
            space_before=3,
            space_after=3,
        )
    )
    header: StyleProfile = field(
        default_factory=lambda: StyleProfile(
            font_name="宋体",
            latin_font_name="Times New Roman",
            font_size_pt=11,
            bold=True,
            alignment="CENTER",
            line_spacing=1.0,
            space_before=3,
            space_after=3,
        )
    )
    borders: TableBorderProfile = field(default_factory=TableBorderProfile)
    used_default: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class FigureProfile:
    image_alignment: str = "CENTER"
    caption: StyleProfile = field(
        default_factory=lambda: StyleProfile(
            font_name="宋体",
            font_size_pt=10.5,
            bold=False,
            alignment="CENTER",
            space_before=6,
            space_after=6,
        )
    )
    used_default: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class ReferenceProfile:
    title: StyleProfile = field(default_factory=StyleProfile)
    body: StyleProfile = field(default_factory=StyleProfile)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class TemplateStyleProfile:
    body: StyleProfile = field(default_factory=StyleProfile)
    heading_1: StyleProfile = field(
        default_factory=lambda: StyleProfile(outline_level=1)
    )
    heading_2: StyleProfile = field(
        default_factory=lambda: StyleProfile(outline_level=2)
    )
    heading_3: StyleProfile = field(
        default_factory=lambda: StyleProfile(outline_level=3)
    )
    heading_4: StyleProfile = field(
        default_factory=lambda: StyleProfile(outline_level=4)
    )
    abstract_title: StyleProfile = field(default_factory=StyleProfile)
    english_abstract_title: StyleProfile = field(default_factory=StyleProfile)
    keywords: StyleProfile = field(default_factory=StyleProfile)
    english_keywords: StyleProfile = field(default_factory=StyleProfile)
    page: PageProfile = field(default_factory=PageProfile)
    header_footer: HeaderFooterProfile = field(default_factory=HeaderFooterProfile)
    table: TableProfile = field(default_factory=TableProfile)
    figure: FigureProfile = field(default_factory=FigureProfile)
    reference: ReferenceProfile = field(default_factory=ReferenceProfile)
    profile_name: str = "template"
    missing_fields: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class ParagraphPrediction:
    index: int
    text: str
    predicted_type: ParagraphType
    confidence: float
    reasons: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class TocResult:
    inserted: bool = False
    existing_toc_detected: bool = False
    message: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class FormatSummary:
    body_paragraphs_modified: int = 0
    heading_1_modified: int = 0
    heading_2_modified: int = 0
    heading_3_modified: int = 0
    heading_4_modified: int = 0
    special_paragraphs_modified: int = 0
    tables_detected: int = 0
    tables_formatted: int = 0
    three_line_tables_applied: int = 0
    table_captions_detected: int = 0
    table_captions_formatted: int = 0
    figures_detected: int = 0
    figures_centered: int = 0
    figure_captions_detected: int = 0
    figure_captions_formatted: int = 0
    cover_heading_formatted: bool = False
    cover_title_formatted: bool = False
    uncertain_paragraphs: int = 0
    toc_inserted: bool = False
    header_footer_copied: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class FormatRunResult:
    summary: FormatSummary
    toc: TocResult
    uncertain_predictions: list[ParagraphPrediction] = field(default_factory=list)
    unsupported_objects: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "summary": self.summary.to_dict(),
            "toc": self.toc.to_dict(),
            "uncertain_predictions": [item.to_dict() for item in self.uncertain_predictions],
            "unsupported_objects": self.unsupported_objects,
            "warnings": self.warnings,
        }
