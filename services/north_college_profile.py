from __future__ import annotations

from app.models import ReferenceProfile, StyleProfile, TableBorderProfile, TableProfile, TemplateStyleProfile


def build_north_college_profile() -> TemplateStyleProfile:
    """河北北方学院论文固定格式规则。

    该 Profile 不依赖学校模板识别，适合普通用户直接上传学生论文后一键套用。
    """

    profile = TemplateStyleProfile(profile_name="hebei_north_college_fixed")
    latin = "Times New Roman"

    profile.body = StyleProfile(
        font_name="宋体",
        latin_font_name=latin,
        font_size_pt=12,
        alignment="JUSTIFY",
        first_line_indent=24,
        line_spacing=20,
        space_before=0,
        space_after=0,
    )
    profile.heading_1 = StyleProfile(
        font_name="黑体",
        latin_font_name=latin,
        font_size_pt=15,
        bold=False,
        alignment="CENTER",
        line_spacing=20,
        space_before=40,
        space_after=20,
        outline_level=1,
    )
    profile.heading_2 = StyleProfile(
        font_name="黑体",
        latin_font_name=latin,
        font_size_pt=14,
        bold=False,
        alignment="LEFT",
        line_spacing=20,
        space_before=24,
        space_after=6,
        outline_level=2,
    )
    profile.heading_3 = StyleProfile(
        font_name="黑体",
        latin_font_name=latin,
        font_size_pt=13,
        bold=False,
        alignment="LEFT",
        line_spacing=20,
        space_before=12,
        space_after=6,
        outline_level=3,
    )
    profile.heading_4 = StyleProfile(
        font_name="黑体",
        latin_font_name=latin,
        font_size_pt=12,
        bold=False,
        alignment="LEFT",
        line_spacing=20,
        space_before=12,
        space_after=6,
        outline_level=4,
    )
    profile.abstract_title = StyleProfile(
        font_name="黑体",
        latin_font_name=latin,
        font_size_pt=15,
        bold=False,
        alignment="CENTER",
        line_spacing=20,
        space_before=40,
        space_after=20,
    )
    profile.english_abstract_title = StyleProfile(
        font_name="Arial",
        latin_font_name="Arial",
        font_size_pt=15,
        bold=False,
        alignment="CENTER",
        line_spacing=20,
        space_before=40,
        space_after=20,
    )
    profile.keywords = StyleProfile(
        font_name="宋体",
        latin_font_name=latin,
        font_size_pt=12,
        bold=False,
        alignment="LEFT",
        first_line_indent=0,
        line_spacing=20,
        space_before=0,
        space_after=0,
    )
    profile.english_keywords = StyleProfile(
        font_name=latin,
        latin_font_name=latin,
        font_size_pt=12,
        bold=True,
        alignment="LEFT",
        first_line_indent=0,
        line_spacing=20,
        space_before=0,
        space_after=0,
    )
    profile.table = TableProfile(
        is_three_line=True,
        caption=StyleProfile(
            font_name="黑体",
            latin_font_name=latin,
            font_size_pt=11,
            bold=False,
            alignment="CENTER",
            space_before=12,
            space_after=6,
        ),
        body=StyleProfile(
            font_name="宋体",
            latin_font_name=latin,
            font_size_pt=11,
            alignment="CENTER",
            line_spacing=1.0,
            space_before=3,
            space_after=3,
        ),
        header=StyleProfile(
            font_name="宋体",
            latin_font_name=latin,
            font_size_pt=11,
            bold=True,
            alignment="CENTER",
            line_spacing=1.0,
            space_before=3,
            space_after=3,
        ),
        borders=TableBorderProfile(
            top_size=12,
            header_bottom_size=8,
            bottom_size=12,
            color="000000",
            vertical=False,
            inside_horizontal=False,
        ),
        used_default=False,
    )
    profile.figure.caption = StyleProfile(
        font_name="黑体",
        latin_font_name=latin,
        font_size_pt=11,
        bold=False,
        alignment="CENTER",
        line_spacing=1.0,
        space_before=6,
        space_after=12,
    )
    profile.figure.image_alignment = "CENTER"
    profile.figure.used_default = False
    profile.reference = ReferenceProfile(
        title=StyleProfile(
            font_name="黑体",
            latin_font_name=latin,
            font_size_pt=15,
            bold=False,
            alignment="CENTER",
            line_spacing=20,
            space_before=40,
            space_after=20,
        ),
        body=StyleProfile(
            font_name="宋体",
            latin_font_name=latin,
            font_size_pt=10.5,
            alignment="LEFT",
            line_spacing=16,
            space_before=3,
            space_after=0,
        ),
    )
    profile.warnings.append("已使用河北北方学院固定格式模块，未依赖模板自动推断正文、标题、图表和参考文献样式。")
    return profile
