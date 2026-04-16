from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from thesis_checker import analyze_docx

ALLOWED_EXTENSIONS = {".docx"}


def save_uploaded_file(uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix.lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        return Path(temp_file.name)


def severity_label(value: str) -> str:
    mapping = {
        "error": "\u932f\u8aa4",
        "warning": "\u8b66\u544a",
        "info": "\u63d0\u793a",
    }
    return mapping.get(value, value)


def render_summary(report: dict) -> None:
    st.subheader("\u6aa2\u67e5\u7e3d\u89bd")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("\u932f\u8aa4", report["summary"]["errors"])
    col2.metric("\u8b66\u544a", report["summary"]["warnings"])
    col3.metric("\u63d0\u793a", report["summary"]["infos"])
    col4.metric("\u7bc0\u6578", report["coverage"]["section_count"])
    st.caption(f"\u6a94\u540d\uff1a{report['file_name']}")


def render_properties(report: dict) -> None:
    st.subheader("\u6587\u4ef6\u5c6c\u6027\u8207\u7248\u9762\u8cc7\u8a0a")
    page_number = report["coverage"]["page_number"]
    protection = report["coverage"]["protection"]
    st.write(f"\u6bb5\u843d\u6578\uff1a{report['coverage']['paragraph_count']}")
    st.write(
        "\u9801\u78bc\u6b04\u4f4d\uff1a"
        + (f"\u5df2\u5075\u6e2c\u5230 ({page_number['format']})" if page_number["present"] else "\u672a\u5075\u6e2c\u5230")
    )
    st.write("\u6d6e\u6c34\u5370\uff1a" + ("\u5df2\u5075\u6e2c\u5230\u76f8\u95dc\u7269\u4ef6" if report["coverage"]["watermark"] else "\u672a\u5075\u6e2c\u5230"))
    st.write(
        "Word \u4fdd\u8b77\uff1a"
        + (f"\u5df2\u555f\u7528 ({protection['mode']})" if protection["enabled"] else "\u672a\u555f\u7528")
    )
    st.dataframe(report["section_results"], use_container_width=True, hide_index=True)


def render_issues(report: dict) -> None:
    st.subheader("\u554f\u984c\u8207\u4fee\u6b63\u5efa\u8b70")
    if not report["issues"]:
        st.success("\u76ee\u524d\u898f\u5247\u4e0b\u672a\u5075\u6e2c\u5230\u660e\u986f\u7684\u683c\u5f0f\u554f\u984c\u3002")
        return

    for issue in report["issues"]:
        title = f"[{severity_label(issue['severity'])}] {issue['title']}"
        with st.expander(title, expanded=issue["severity"] == "error"):
            st.write(f"\u5206\u985e\uff1a{issue['category']}")
            if issue.get("location"):
                st.write(f"\u4f4d\u7f6e\uff1a{issue['location']}")
            st.write(issue["details"])
            if issue.get("suggestion"):
                st.info(f"\u5efa\u8b70\uff1a{issue['suggestion']}")


def render_limits(report: dict) -> None:
    st.subheader("\u5de5\u5177\u9650\u5236")
    for item in report["summary"]["limitations"]:
        st.write(f"- {item}")


def main() -> None:
    st.set_page_config(
        page_title="\u8ad6\u6587 Word \u683c\u5f0f\u6aa2\u67e5\u5668",
        layout="wide",
    )

    st.title("\u8ad6\u6587 Word \u683c\u5f0f\u6aa2\u67e5\u5668")
    st.write(
        "\u4e0a\u50b3 `.docx` \u8ad6\u6587\u6a94\uff0c\u7cfb\u7d71\u6703\u4f9d\u64da\u5716\u66f8\u9928\u898f\u7bc4\u81ea\u52d5\u7522\u51fa\u8a73\u7d30\u6aa2\u67e5\u5831\u544a\u3002"
    )

    uploaded_file = st.file_uploader(
        "\u9078\u64c7 Word \u6a94\u6848",
        type=["docx"],
        accept_multiple_files=False,
    )

    if uploaded_file is None:
        st.caption("\u76ee\u524d\u53ea\u652f\u63f4 `.docx` \u683c\u5f0f\u3002")
        return

    if Path(uploaded_file.name).suffix.lower() not in ALLOWED_EXTENSIONS:
        st.error("\u76ee\u524d\u53ea\u652f\u63f4 .docx \u683c\u5f0f\u3002")
        return

    temp_path = save_uploaded_file(uploaded_file)
    try:
        with st.spinner("\u6b63\u5728\u6aa2\u67e5\u6587\u4ef6\u683c\u5f0f..."):
            report = analyze_docx(temp_path)
    finally:
        temp_path.unlink(missing_ok=True)

    render_summary(report)
    st.subheader("\u672c\u6b21\u6aa2\u67e5\u9805\u76ee")
    st.write(" | ".join(report["summary"]["checked_items"]))
    render_properties(report)
    render_issues(report)
    render_limits(report)


if __name__ == "__main__":
    main()
