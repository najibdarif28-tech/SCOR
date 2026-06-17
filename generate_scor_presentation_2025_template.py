import argparse
import os
import tempfile
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt


def get_layout(prs: Presentation, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise ValueError(f"Layout '{name}' not found in template.")


def has_layout(prs: Presentation, name: str) -> bool:
    return any(layout.name == name for layout in prs.slide_layouts)


def normalize_template_for_python_pptx(template_path: Path) -> tuple[Path, bool]:
    """
    python-pptx cannot open .potx directly (template content-type).
    Convert to a temporary .pptx package when needed.
    Returns: (loadable_path, should_cleanup)
    """
    if template_path.suffix.lower() != ".potx":
        return template_path, False

    fd, temp_path = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    with zipfile.ZipFile(template_path, "r") as zin, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                    b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                )
            zout.writestr(item, data)

    return Path(temp_path), True


def set_title(slide, text: str):
    if slide.shapes.title is not None:
        slide.shapes.title.text = text


def set_placeholder_text(slide, idx: int, text: str):
    if idx in slide.placeholders:
        slide.placeholders[idx].text = text


def add_cover(prs: Presentation):
    slide = prs.slides.add_slide(get_layout(prs, "Cover 1"))
    set_title(slide, "SCOR Managing Agency")
    set_placeholder_text(
        slide,
        11,
        "Contractual position, use case and recent client topics\n"
        "Prepared using Moody's Scenario Generator template style",
    )
    set_placeholder_text(slide, 99, "Client: SCOR Managing Agency Limited")
    set_placeholder_text(slide, 100, "Date: June 2026")
    set_placeholder_text(slide, 101, "Prepared by: Moody's Analytics")


def add_exec_summary(prs: Presentation):
    slide = prs.slides.add_slide(get_layout(prs, "Executive Summary/Key Takeaways 1"))
    set_title(slide, "Executive summary")
    body = (
        "• SCOR renews annually under the long-running Moody's agreement framework.\n"
        "• 2025 renewal fee is GBP 53,950 (+3.0% vs 2024), below referenced standard uplift range.\n"
        "• Amendment formalizes legal name/address move to SCOR Managing Agency Limited, 8 Bishopsgate.\n"
        "• Client requested one-year renewal over multi-year commitment.\n"
        "• Recent discussions center on CSV output delivery and SG corporate-bond parameter settings."
    )
    set_placeholder_text(slide, 13, body)
    set_placeholder_text(slide, 15, "Sources: 2023-2025 renewals, 2025 amendment, client correspondence")


def add_contract_table(prs: Presentation):
    layout_name = "Title and Table" if has_layout(prs, "Title and Table") else "1 Column"
    slide = prs.slides.add_slide(get_layout(prs, layout_name))
    set_title(slide, "Contractual position summary")

    if layout_name == "Title and Table":
        ph = slide.placeholders[11]
        table = ph.insert_table(rows=6, cols=5).table
    else:
        container = slide.placeholders[2]
        table = slide.shapes.add_table(
            rows=6,
            cols=5,
            left=container.left,
            top=container.top,
            width=container.width,
            height=container.height,
        ).table

    headers = ["Document", "Date", "Agreement Ref", "Key point", "Fee impact"]
    for c, h in enumerate(headers):
        table.cell(0, c).text = h

    rows = [
        ["Base Order Form", "23 Dec 2016", "00073841.0", "Framework baseline for current service", "Legacy baseline"],
        ["Renewal Notification", "23 Dec 2023", "00073841.7", "Annual renewal continuity", "GBP 50,364"],
        ["Renewal Notification", "23 Dec 2024", "00073841.8", "Annual renewal continuity", "GBP 52,379"],
        ["Renewal Notification", "23 Dec 2025", "00073841.9", "Annual renewal continuity", "GBP 53,950"],
        ["Amendment 1", "23 Dec 2025", "00073841.9", "Name/address + sanctions clause update", "No explicit change"],
    ]
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            table.cell(r, c).text = value


def add_fee_trend(prs: Presentation):
    slide = prs.slides.add_slide(get_layout(prs, "1 Column"))
    set_title(slide, "Commercial trend: annual renewal fees")
    set_placeholder_text(
        slide,
        17,
        "Fee progression reflects controlled uplift with 2025 increase at 3.0% year-over-year.",
    )
    set_placeholder_text(slide, 18, "Market context from sales notes: typical range ~6-10%.")

    chart_data = CategoryChartData()
    chart_data.categories = ["2023", "2024", "2025"]
    chart_data.add_series("Annual Fee (GBP)", (50364, 52379, 53950))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        slide.placeholders[2].left,
        slide.placeholders[2].top,
        slide.placeholders[2].width,
        slide.placeholders[2].height,
        chart_data,
    ).chart
    chart.has_legend = False
    chart.series[0].format.line.color.rgb = RGBColor(0x00, 0x35, 0x7A)


def add_use_case(prs: Presentation):
    slide = prs.slides.add_slide(get_layout(prs, "2 Column, Equal - Icons"))
    set_title(slide, "Use case and operating model")
    set_placeholder_text(
        slide,
        21,
        "Current usage\n"
        "• Scenario Generator used for Solvency II capital modelling\n"
        "• Team currently runs SG internally for calibration/output production",
    )
    set_placeholder_text(
        slide,
        22,
        "Requested evolution\n"
        "• Quarterly calibration output in CSV format\n"
        "• Clarification needed on NumberOfBonds and Coupon parameters",
    )

    # Icon accents in the two content columns
    left_box = slide.placeholders[27]
    right_box = slide.placeholders[28]
    for box, label in [(left_box, "1"), (right_box, "2")]:
        icon = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, box.left, box.top, box.width, box.height)
        icon.fill.solid()
        icon.fill.fore_color.rgb = RGBColor(0x00, 0x35, 0x7A)
        icon.line.color.rgb = RGBColor(0x00, 0x35, 0x7A)
        p = icon.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.CENTER


def add_recent_topics(prs: Presentation):
    layout_name = "Title and Table" if has_layout(prs, "Title and Table") else "1 Column"
    slide = prs.slides.add_slide(get_layout(prs, layout_name))
    set_title(slide, "Recent topics with SCOR")

    if layout_name == "Title and Table":
        ph = slide.placeholders[11]
        table = ph.insert_table(rows=5, cols=4).table
    else:
        container = slide.placeholders[2]
        table = slide.shapes.add_table(
            rows=5,
            cols=4,
            left=container.left,
            top=container.top,
            width=container.width,
            height=container.height,
        ).table

    headers = ["Topic", "Client ask", "Why it matters", "Recommended follow-up"]
    for c, h in enumerate(headers):
        table.cell(0, c).text = h

    rows = [
        [
            "Renewal term",
            "Single-year renewal instead of multi-year",
            "Preference for flexibility in commitment horizon",
            "Position multi-year as optional value lever, not prerequisite",
        ],
        [
            "Output delivery",
            "Quarterly calibration output in CSV",
            "Could reduce operational burden for SCOR team",
            "Assess delivery model and commercial treatment",
        ],
        [
            "Corporate bond parameters",
            "Clarify NumberOfBonds and Coupon settings",
            "Directly impacts modeled returns",
            "Provide documented parameter guidance",
        ],
        [
            "Dummy values",
            "Implication of NumberOfBonds=1 and Coupon=0",
            "Risk of unrealistic outputs/model distortion",
            "Run controlled test and share quantified impact",
        ],
    ]
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            table.cell(r, c).text = value


def add_next_steps(prs: Presentation):
    slide = prs.slides.add_slide(get_layout(prs, "Executive Summary/Key Takeaways 2"))
    set_title(slide, "Recommended next steps")
    body = (
        "1) Confirm one-year renewal final path and commercial narrative.\n"
        "2) Validate contract records reflect legal name and licensed address updates.\n"
        "3) Scope feasibility/options for quarterly CSV calibration output service.\n"
        "4) Provide technical response on NumberOfBonds/Coupon parameter boundaries.\n"
        "5) Align Sales, Product Specialist and Contracts follow-up owners."
    )
    set_placeholder_text(slide, 13, body)


def add_back_cover(prs: Presentation):
    slide = prs.slides.add_slide(get_layout(prs, "Back Cover 1"))
    set_title(slide, "Thank you")
    set_placeholder_text(slide, 15, "Questions and discussion")


def build(template_path: Path, output_path: Path):
    normalized_template, cleanup = normalize_template_for_python_pptx(template_path)
    prs = Presentation(str(normalized_template))

    try:
        add_cover(prs)
        add_exec_summary(prs)
        add_contract_table(prs)
        add_fee_trend(prs)
        add_use_case(prs)
        add_recent_topics(prs)
        add_next_steps(prs)
        add_back_cover(prs)
        prs.save(str(output_path))
    finally:
        if cleanup and normalized_template.exists():
            normalized_template.unlink()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--template",
        default="2025_Scenario_Generator_FRR (1).pptx",
        help="Path to Moody's template .pptx",
    )
    parser.add_argument(
        "--output",
        default="SCOR_Contractual_Summary_Moodys_2025_Template.pptx",
        help="Output presentation file",
    )
    args = parser.parse_args()
    build(Path(args.template), Path(args.output))
