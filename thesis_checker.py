from __future__ import annotations

import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Length, Pt
from lxml import etree

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

ALLOWED_CHINESE_FONTS = {"DFKai-SB", "\u6a19\u6977\u9ad4"}
ALLOWED_ENGLISH_FONTS = {"Times New Roman"}

CHAPTER_PATTERN = re.compile(r"^\u7b2c[\u4e00-\u9fff0-9]+\u7ae0")
SECTION_PATTERN = re.compile(r"^(?:\u7b2c[\u4e00-\u9fff0-9]+\u7bc0|[0-9]+-[0-9]+(?:-[0-9]+)?)$")
ABSTRACT_HEADINGS = {"ABSTRACT", "\u6458\u8981"}
FRONT_HEADINGS = {
    "\u81f4\u8b1d",
    "\u8b1d\u8a8c",
    "\u76ee\u9304",
    "\u8868\u76ee\u9304",
    "\u5716\u76ee\u9304",
    "\u53c3\u8003\u6587\u737b",
}
KEYWORD_PREFIXES = ("Keywords", "\u95dc\u9375\u5b57")


@dataclass
class Issue:
    severity: str
    category: str
    title: str
    details: str
    location: str | None = None
    suggestion: str | None = None


def length_to_cm(length: Length | None) -> float | None:
    return None if length is None else round(length.cm, 2)


def pt_value(value) -> float | None:
    if value is None:
        return None
    try:
        return round(value.pt, 1)
    except Exception:
        return None


def alignment_name(value: int | None) -> str:
    mapping = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
        WD_ALIGN_PARAGRAPH.DISTRIBUTE: "distribute",
    }
    return mapping.get(value, "unspecified")


def paragraph_text(paragraph) -> str:
    return re.sub(r"\s+", " ", paragraph.text or "").strip()


def visible_paragraphs(document: Document):
    for index, paragraph in enumerate(document.paragraphs, start=1):
        text = paragraph_text(paragraph)
        if text:
            yield index, paragraph, text


def run_font_name(run) -> str | None:
    if run.font.name:
        return run.font.name
    rfonts = getattr(getattr(run._element, "rPr", None), "rFonts", None)
    if rfonts is None:
        return None
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        value = rfonts.get(f"{{{WORD_NS['w']}}}{attr}")
        if value:
            return value
    return None


def iter_runs_with_text(paragraph) -> Iterable:
    for run in paragraph.runs:
        if run.text and run.text.strip():
            yield run


def paragraph_fonts(paragraph) -> set[str]:
    return {
        font
        for run in iter_runs_with_text(paragraph)
        for font in [run_font_name(run)]
        if font
    }


def paragraph_sizes(paragraph) -> set[float]:
    return {
        size
        for run in iter_runs_with_text(paragraph)
        for size in [pt_value(run.font.size)]
        if size is not None
    }


def contains_cjk(text: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def contains_ascii_letters_or_digits(text: str) -> bool:
    return bool(re.search(r"[A-Za-z0-9]", text))


def paragraph_line_spacing(paragraph) -> tuple[str, float | None]:
    fmt = paragraph.paragraph_format
    if fmt.line_spacing_rule == WD_LINE_SPACING.ONE_POINT_FIVE:
        return "1.5 lines", 1.5
    if isinstance(fmt.line_spacing, Pt):
        return "exact", round(fmt.line_spacing.pt, 1)
    if isinstance(fmt.line_spacing, float):
        return "multiple", round(float(fmt.line_spacing), 2)
    if fmt.line_spacing is not None:
        try:
            return "multiple", round(float(fmt.line_spacing), 2)
        except Exception:
            return "set", None
    return "unspecified", None


def is_bold_paragraph(paragraph) -> bool:
    runs = list(iter_runs_with_text(paragraph))
    if not runs:
        return False
    return sum(1 for run in runs if run.bold) >= max(1, len(runs) // 2)


def extract_docx_xml(docx_path: Path, member: str) -> bytes | None:
    try:
        with zipfile.ZipFile(docx_path) as archive:
            return archive.read(member)
    except KeyError:
        return None


def list_docx_members(docx_path: Path, prefix: str) -> list[str]:
    with zipfile.ZipFile(docx_path) as archive:
        return [name for name in archive.namelist() if name.startswith(prefix)]


def parse_xml(data: bytes | None):
    return None if not data else etree.fromstring(data)


def page_number_info(docx_path: Path) -> dict[str, str | bool]:
    footer_files = list_docx_members(docx_path, "word/footer")
    if not footer_files:
        return {"present": False, "format": "none"}
    found = False
    page_format = "arabic"
    for member in footer_files:
        text = (extract_docx_xml(docx_path, member) or b"").decode("utf-8", errors="ignore")
        if "PAGE" in text:
            found = True
            if "ROMAN" in text:
                page_format = "roman"
    return {"present": found, "format": page_format if found else "none"}


def has_watermark(docx_path: Path) -> bool:
    for member in list_docx_members(docx_path, "word/header"):
        text = (extract_docx_xml(docx_path, member) or b"").decode("utf-8", errors="ignore")
        if any(marker in text for marker in ("PowerPlusWaterMarkObject", "w:pict", "v:shape", "w:txbxContent", "o:spid")):
            return True
    return False


def document_protection(docx_path: Path) -> dict[str, str | bool]:
    settings_xml = parse_xml(extract_docx_xml(docx_path, "word/settings.xml"))
    if settings_xml is None:
        return {"enabled": False, "mode": "unknown"}
    protection = settings_xml.find("w:documentProtection", namespaces=WORD_NS)
    if protection is None:
        return {"enabled": False, "mode": "none"}
    edit_mode = protection.get(f"{{{WORD_NS['w']}}}edit", "unknown")
    enforcement = protection.get(f"{{{WORD_NS['w']}}}enforcement", "0") == "1"
    return {"enabled": enforcement, "mode": edit_mode}


def classify_paragraph(text: str) -> str:
    if text in ABSTRACT_HEADINGS:
        return "abstract_heading"
    if text in FRONT_HEADINGS:
        return "front_heading"
    if CHAPTER_PATTERN.match(text):
        return "chapter_heading"
    if SECTION_PATTERN.match(text):
        return "section_heading"
    if any(text.startswith(prefix) for prefix in KEYWORD_PREFIXES):
        return "keywords"
    return "body"


def add_issue(
    issues: list[Issue],
    severity: str,
    category: str,
    title: str,
    details: str,
    location: str | None = None,
    suggestion: str | None = None,
) -> None:
    issues.append(Issue(severity, category, title, details, location, suggestion))


def analyze_sections(document: Document, issues: list[Issue]) -> list[dict]:
    results = []
    for idx, section in enumerate(document.sections, start=1):
        page_width = length_to_cm(section.page_width)
        page_height = length_to_cm(section.page_height)
        top = length_to_cm(section.top_margin)
        bottom = length_to_cm(section.bottom_margin)
        left = length_to_cm(section.left_margin)
        right = length_to_cm(section.right_margin)
        results.append(
            {
                "index": idx,
                "page_width_cm": page_width,
                "page_height_cm": page_height,
                "top_cm": top,
                "bottom_cm": bottom,
                "left_cm": left,
                "right_cm": right,
                "start_type": str(section.start_type),
            }
        )
        if (page_width, page_height) != (21.0, 29.7):
            add_issue(issues, "error", "Page layout", f"Section {idx} is not A4 size", f"Detected {page_width} x {page_height} cm. The spec requires A4 21.0 x 29.7 cm.", f"Section {idx}", "Change the section page size to A4 in Word.")
        for side, expected_value, current in (("top", 2.5, top), ("bottom", 2.5, bottom), ("left", 3.0, left), ("right", 2.0, right)):
            if current is None or abs(current - expected_value) > 0.11:
                add_issue(issues, "error", "Page layout", f"Section {idx} {side} margin is out of spec", f"Detected {current} cm. The spec requires {expected_value} cm.", f"Section {idx}", "Adjust the page margins for this section.")
        if idx > 1 and section.start_type != WD_SECTION_START.NEW_PAGE:
            add_issue(issues, "warning", "Section breaks", f"Section {idx} does not clearly start on a new page", "The spec expects major sections and chapters to begin on a new page.", f"Section {idx}", "Use a next-page section break for each chapter.")
    return results


def check_paragraphs(document: Document, issues: list[Issue]) -> list[dict]:
    paragraph_summaries = []
    in_abstract = False
    for index, paragraph, text in visible_paragraphs(document):
        kind = classify_paragraph(text)
        if kind == "abstract_heading":
            in_abstract = True
        elif kind in {"chapter_heading", "front_heading"} and text != "\u6458\u8981":
            in_abstract = False
        fonts = paragraph_fonts(paragraph)
        sizes = paragraph_sizes(paragraph)
        alignment = alignment_name(paragraph.alignment)
        line_spacing_label, line_spacing_value = paragraph_line_spacing(paragraph)
        location = f"Paragraph {index}"
        paragraph_summaries.append({"index": index, "text": text[:120], "kind": kind, "fonts": sorted(fonts), "sizes": sorted(sizes), "alignment": alignment, "line_spacing": line_spacing_label})
        if kind == "chapter_heading":
            if alignment != "center":
                add_issue(issues, "error", "Headings", f"{location} chapter heading is not centered", f'Detected heading text: "{text}". Alignment is {alignment}.', location, "Center the chapter heading.")
            if not is_bold_paragraph(paragraph):
                add_issue(issues, "error", "Headings", f"{location} chapter heading is not bold", f'Detected heading text: "{text}".', location, "Make the chapter heading bold.")
            if 16.0 not in sizes:
                add_issue(issues, "error", "Headings", f"{location} chapter heading is not 16 pt", f"Detected font sizes: {sorted(sizes) or 'not explicitly set'}.", location, "Set chapter heading size to 16 pt.")
        elif kind == "section_heading":
            if not is_bold_paragraph(paragraph):
                add_issue(issues, "error", "Headings", f"{location} section heading is not bold", f'Detected heading text: "{text}".', location, "Make the section heading bold.")
            if 14.0 not in sizes:
                add_issue(issues, "error", "Headings", f"{location} section heading is not 14 pt", f"Detected font sizes: {sorted(sizes) or 'not explicitly set'}.", location, "Set section heading size to 14 pt.")
        elif kind in {"abstract_heading", "front_heading"}:
            if alignment != "center":
                add_issue(issues, "warning", "Front matter", f"{location} front-matter heading is not centered", f'Detected heading text: "{text}". Alignment is {alignment}.', location, "Center the heading.")
        elif kind == "keywords":
            if 14.0 not in sizes:
                add_issue(issues, "warning", "Abstract and keywords", f"{location} keyword paragraph is not 14 pt", f"Detected font sizes: {sorted(sizes) or 'not explicitly set'}.", location, "Set the keyword paragraph to 14 pt.")
        else:
            if in_abstract or kind == "body":
                if contains_cjk(text) and fonts and not fonts.intersection(ALLOWED_CHINESE_FONTS):
                    add_issue(issues, "warning", "Body font", f"{location} Chinese text may use the wrong font", f"Detected fonts: {', '.join(sorted(fonts))}. The spec expects DFKai-SB / BiauKai for Chinese body text.", location, "Change Chinese body text to BiauKai / DFKai-SB.")
                if contains_ascii_letters_or_digits(text):
                    english_fonts = {run_font_name(run) for run in iter_runs_with_text(paragraph) if contains_ascii_letters_or_digits(run.text)}
                    english_fonts = {font for font in english_fonts if font}
                    if english_fonts and not english_fonts.issubset(ALLOWED_ENGLISH_FONTS):
                        add_issue(issues, "warning", "English and numbers", f"{location} English or numeric text may use the wrong font", f"Detected fonts: {', '.join(sorted(english_fonts))}. The spec expects Times New Roman.", location, "Change English and numeric text to Times New Roman.")
                if sizes and 12.0 not in sizes:
                    add_issue(issues, "warning", "Body size", f"{location} body text is not 12 pt", f"Detected font sizes: {sorted(sizes)}. The spec expects 12 pt body text.", location, "Set body text to 12 pt.")
                if line_spacing_value is None or abs(line_spacing_value - 1.5) > 0.05:
                    add_issue(issues, "warning", "Line spacing", f"{location} is not using 1.5 line spacing", f"Detected line spacing: {line_spacing_label} {line_spacing_value or ''}".strip(), location, "Set abstract and body paragraphs to 1.5 line spacing.")
    return paragraph_summaries


def summarize_document(document: Document, docx_path: Path, issues: list[Issue]) -> dict:
    page_number = page_number_info(docx_path)
    watermark = has_watermark(docx_path)
    protection = document_protection(docx_path)
    if not page_number["present"]:
        add_issue(issues, "warning", "Page numbers", "No page-number field was detected", "The spec requires centered Roman page numbers for front matter and Arabic page numbers for the main body. No PAGE field was found in footer XML.", suggestion="Insert Word page-number fields in the footer and split front matter from the main body.")
    if not watermark:
        add_issue(issues, "warning", "Watermark", "No watermark object was detected", "The spec requires a watermark from the acknowledgements page onward. This tool can only detect whether watermark-like objects exist in the DOCX package.", suggestion="Check section headers in Word and confirm the watermark starts from the acknowledgements page.")
    if not protection["enabled"]:
        add_issue(issues, "info", "Protection", "No Word document protection was detected", "The library spec mentions document security, but that requirement often applies to the final upload file or PDF. No enabled documentProtection node was found in this DOCX.", suggestion="If the school expects PDF security settings, confirm them after exporting the final file.")
    return {"paragraph_count": len(document.paragraphs), "section_count": len(document.sections), "page_number": page_number, "watermark": watermark, "protection": protection}


def analyze_docx(docx_path: str | Path) -> dict:
    path = Path(docx_path)
    document = Document(path)
    issues: list[Issue] = []
    section_results = analyze_sections(document, issues)
    paragraph_results = check_paragraphs(document, issues)
    coverage = summarize_document(document, path, issues)
    severity_rank = {"error": 0, "warning": 1, "info": 2}
    issues_sorted = sorted(issues, key=lambda item: (severity_rank.get(item.severity, 9), item.category, item.location or ""))
    return {
        "file_name": path.name,
        "summary": {
            "errors": sum(1 for item in issues_sorted if item.severity == "error"),
            "warnings": sum(1 for item in issues_sorted if item.severity == "warning"),
            "infos": sum(1 for item in issues_sorted if item.severity == "info"),
            "checked_items": ["A4 page size", "Page margins", "Chapter heading format", "Section heading format", "Abstract and keywords", "Body font and size", "1.5 line spacing", "Page-number fields", "Watermark existence", "Word protection existence"],
            "limitations": ["DOCX analysis cannot fully reconstruct final pagination, so odd-page chapter starts cannot be guaranteed.", "Watermark checks only verify that watermark-like objects exist in the DOCX package.", "If fonts are inherited only from styles or themes, the tool may need to make conservative guesses.", "Paper weight, duplex printing, and final PDF security still need manual review."],
        },
        "coverage": coverage,
        "section_results": section_results,
        "paragraph_results": paragraph_results,
        "issues": [asdict(item) for item in issues_sorted],
    }
