from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import streamlit as st
from docx.opc.exceptions import PackageNotFoundError

ROOT_DIR = Path(__file__).resolve().parent
BACKEND_DIR = ROOT_DIR / "backend"
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from services.formatter import format_document
from services.north_college_profile import build_north_college_profile
from services.report_builder import write_reports
from services.student_analyzer import analyze_student_document
from services.template_analyzer import analyze_template


MAX_UPLOAD_MB = 50
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024


@dataclass(slots=True)
class ProcessedResult:
    processed_docx: bytes
    report_docx: bytes
    report_json: dict[str, Any]
    summary: dict[str, Any]
    warnings: list[str]


def setup_page() -> None:
    st.set_page_config(
        page_title="论文格式自动标准化助手",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    st.markdown(
        """
        <style>
        .main .block-container {
            max-width: 1120px;
            padding-top: 2.2rem;
            padding-bottom: 3rem;
        }
        .hero {
            padding: 2rem 2.25rem;
            border: 1px solid #d8e2dc;
            border-radius: 14px;
            background: linear-gradient(135deg, #f8fbfa 0%, #eef6f3 100%);
            margin-bottom: 1.4rem;
        }
        .hero h1 {
            margin: 0 0 .5rem 0;
            font-size: 2.4rem;
            line-height: 1.12;
            color: #17342f;
            letter-spacing: 0;
        }
        .hero p {
            color: #4c635e;
            font-size: 1.05rem;
            margin: 0;
        }
        .notice {
            border-left: 4px solid #1d6f5f;
            background: #f4faf8;
            padding: .9rem 1rem;
            border-radius: 8px;
            color: #344c47;
            margin: .75rem 0 1.25rem;
        }
        .small-muted {
            color: #66736f;
            font-size: .92rem;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.55rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        """
        <section class="hero">
          <h1>论文格式自动标准化助手</h1>
          <p>上传学生论文，系统会按学校论文规范自动修正文档格式、表格三线表、图题表题与目录字段，并生成可下载的 Word 文件。</p>
        </section>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="notice">
        <strong>隐私提示：</strong>
        本演示版使用服务器临时目录处理文件。处理完成后，程序会立即读取结果到当前会话并删除临时目录，
        不会在服务器磁盘长期保存论文原文或处理后的文档。请勿上传涉密、未授权或不适合境外云服务处理的论文。
        </div>
        """,
        unsafe_allow_html=True,
    )


def validate_uploaded_docx(file_bytes: bytes, label: str) -> None:
    if len(file_bytes) > MAX_UPLOAD_BYTES:
        raise ValueError(f"{label} 超过 {MAX_UPLOAD_MB}MB 上传限制。")
    try:
        with zipfile.ZipFile(BytesIO(file_bytes)) as archive:
            names = set(archive.namelist())
    except Exception as exc:
        raise ValueError(f"{label} 不是有效的 .docx 文件，或文件已经损坏。") from exc
    if "[Content_Types].xml" not in names or "word/document.xml" not in names:
        raise ValueError(f"{label} 缺少 Word 文档必要结构，请确认上传的是 .docx 文件。")


def safe_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def process_document(
    *,
    student_bytes: bytes,
    template_bytes: bytes | None,
    mode: str,
) -> ProcessedResult:
    validate_uploaded_docx(student_bytes, "学生论文")
    if mode == "template":
        if template_bytes is None:
            raise ValueError("请上传学校论文模板，或切换为河北北方学院固定格式。")
        validate_uploaded_docx(template_bytes, "学校论文模板")

    temp_root = Path(tempfile.mkdtemp(prefix="thesis_formatter_streamlit_"))
    try:
        student_path = temp_root / "student.docx"
        template_path = temp_root / "template.docx"
        processed_path = temp_root / "processed.docx"
        report_json_path = temp_root / "format_report.json"
        report_docx_path = temp_root / "report.docx"
        report_md_path = temp_root / "report.md"

        safe_write_bytes(student_path, student_bytes)
        if template_bytes is not None:
            safe_write_bytes(template_path, template_bytes)

        if mode == "north_college":
            template_profile = build_north_college_profile()
        else:
            template_profile = analyze_template(template_path)

        predictions = analyze_student_document(student_path)
        format_result = format_document(
            student_docx_path=student_path,
            output_docx_path=processed_path,
            template_profile=template_profile,
            predictions=predictions,
            insert_toc=True,
        )
        if not any(item.predicted_type.startswith("heading") for item in predictions):
            format_result.warnings.append("学生论文中没有检测到明显标题，请人工复核论文结构。")

        report_data = write_reports(
            template_profile,
            predictions,
            format_result,
            report_json_path,
            report_docx_path,
            report_md_path,
        )

        return ProcessedResult(
            processed_docx=processed_path.read_bytes(),
            report_docx=report_docx_path.read_bytes(),
            report_json=report_data,
            summary=format_result.summary.to_dict(),
            warnings=template_profile.warnings + format_result.warnings,
        )
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


def render_summary(summary: dict[str, Any]) -> None:
    st.subheader("处理摘要")
    cols = st.columns(4)
    metrics = [
        ("正文段落", summary.get("body_paragraphs_modified", 0)),
        ("一级标题", summary.get("heading_1_modified", 0)),
        ("二级标题", summary.get("heading_2_modified", 0)),
        ("三级标题", summary.get("heading_3_modified", 0)),
        ("表格", summary.get("tables_formatted", 0)),
        ("表题", summary.get("table_captions_formatted", 0)),
        ("图片", summary.get("figures_centered", 0)),
        ("图题", summary.get("figure_captions_formatted", 0)),
    ]
    for index, (label, value) in enumerate(metrics):
        cols[index % 4].metric(label, value)

    status_cols = st.columns(3)
    status_cols[0].metric("目录", "已处理" if summary.get("toc_inserted") else "未插入")
    status_cols[1].metric("页眉页脚", "已处理" if summary.get("header_footer_copied") else "未处理")
    status_cols[2].metric("低置信标题", summary.get("uncertain_paragraphs", 0))


def clear_result_state() -> None:
    for key in ("processed_result", "last_student_name"):
        st.session_state.pop(key, None)


def main() -> None:
    setup_page()
    render_header()

    with st.sidebar:
        st.header("使用范围")
        st.caption("当前演示版仅支持 .docx。目录页码可能需要在 Word/WPS 中右键更新域。")
        st.divider()
        st.caption("建议先用非涉密论文试用；正式提交前请人工复核封面、目录、图表、参考文献和学校特殊要求。")

    st.subheader("选择格式化模式")
    mode_label = st.radio(
        "学校模板来源",
        options=("河北北方学院固定格式", "上传学校论文模板自动分析"),
        horizontal=True,
    )
    mode = "north_college" if mode_label == "河北北方学院固定格式" else "template"

    left, right = st.columns(2)
    with left:
        st.markdown("#### 学生论文")
        student_file = st.file_uploader(
            "上传学生已写好的论文（.docx）",
            type=["docx"],
            key="student_docx",
            on_change=clear_result_state,
        )
        st.caption("系统会保留文字内容，并重建网页粘贴表格为干净三线表。")

    template_file = None
    with right:
        st.markdown("#### 学校论文模板")
        if mode == "template":
            template_file = st.file_uploader(
                "上传学校标准论文模板（.docx）",
                type=["docx"],
                key="template_docx",
                on_change=clear_result_state,
            )
            st.caption("如果模板样式不规范，系统会使用启发式规则推断。")
        else:
            st.info("已选择河北北方学院固定格式，无需上传学校模板。")

    st.markdown(
        '<p class="small-muted">处理逻辑：上传文件 → 临时目录处理 → 生成 Word → 读入会话内存 → 立即删除服务器临时文件。</p>',
        unsafe_allow_html=True,
    )

    can_process = student_file is not None and (mode == "north_college" or template_file is not None)
    if st.button("开始格式化", type="primary", disabled=not can_process, use_container_width=True):
        clear_result_state()
        try:
            student_bytes = student_file.getvalue() if student_file is not None else b""
            template_bytes = template_file.getvalue() if template_file is not None else None
            with st.status("正在处理论文，请稍候...", expanded=True) as status:
                st.write("正在校验 .docx 文件结构")
                st.write("正在分析论文结构和标题层级")
                st.write("正在清洗表格并套用论文三线表")
                st.write("正在生成格式化后的 Word 文档和报告")
                result = process_document(
                    student_bytes=student_bytes,
                    template_bytes=template_bytes,
                    mode=mode,
                )
                status.update(label="处理完成", state="complete", expanded=False)
            st.session_state["processed_result"] = result
            st.session_state["last_student_name"] = student_file.name if student_file else "processed.docx"
            st.success("格式化完成。服务器临时目录已删除，结果仅保留在当前浏览器会话中供下载。")
        except ValueError as exc:
            st.error(str(exc))
        except PackageNotFoundError:
            st.error("Word 文档损坏或不是有效的 .docx 文件，请重新另存为 .docx 后再上传。")
        except zipfile.BadZipFile:
            st.error("Word 文档压缩结构损坏，无法处理。请用 Word/WPS 打开后另存为新的 .docx。")
        except Exception as exc:
            st.error(f"处理失败：{exc}")

    result = st.session_state.get("processed_result")
    if isinstance(result, ProcessedResult):
        st.divider()
        render_summary(result.summary)

        st.subheader("下载结果")
        col_a, col_b, col_c = st.columns(3)
        col_a.download_button(
            "下载格式化后的论文",
            data=result.processed_docx,
            file_name="processed.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        col_b.download_button(
            "下载格式修正报告",
            data=result.report_docx,
            file_name="format_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        col_c.download_button(
            "下载 JSON 明细",
            data=json.dumps(result.report_json, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name="format_report.json",
            mime="application/json",
            use_container_width=True,
        )

        if result.warnings:
            with st.expander("处理提示与需要人工复核的内容"):
                for item in result.warnings[:30]:
                    st.write(f"- {item}")

        if st.button("清除当前会话中的结果文件", use_container_width=True):
            clear_result_state()
            st.success("当前会话中的结果文件已清除。")
            st.rerun()


if __name__ == "__main__":
    main()
