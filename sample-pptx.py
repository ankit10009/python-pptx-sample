"""
PPTX Overflow Detection & Auto-Adjustment Feedback Loop
========================================================
Practical example for FastAPI project status report generation.

Handles:
  - Text overflow detection (by character estimation + line counting)
  - Auto font-size shrink loop
  - Auto box height expansion fallback
  - Table row height overflow detection
  - Logging of every adjustment made (useful for debugging in PyCharm)

Install:
    pip install python-pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import logging
import io

# ── Logging setup (visible in PyCharm Run/Debug console) ─────────────────────
logging.basicConfig(
    level=logging.DEBUG,
    format="%(levelname)s | %(funcName)s | %(message)s"
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# CORE: Overflow Detection
# ─────────────────────────────────────────────────────────────────────────────

def estimate_text_height(text: str, font_size_pt: float, box_width_inches: float) -> float:
    """
    Estimates the rendered height (in inches) of text inside a box.

    python-pptx does NOT render text — it has no layout engine.
    So we approximate using character width heuristics.

    Args:
        text:             The full string to measure
        font_size_pt:     Font size in points
        box_width_inches: Width of the text box in inches

    Returns:
        Estimated height in inches needed to display the text
    """
    # Average character width: ~0.6x of font height (rough typographic rule)
    # 1 point = 1/72 inch
    font_height_inches = font_size_pt / 72.0
    avg_char_width_inches = font_height_inches * 0.55

    chars_per_line = max(1, int(box_width_inches / avg_char_width_inches))

    # Count actual lines (respect manual newlines too)
    total_lines = 0
    for paragraph in text.split("\n"):
        if len(paragraph) == 0:
            total_lines += 1  # blank line still takes space
        else:
            # How many wrapped lines does this paragraph produce?
            total_lines += max(1, -(-len(paragraph) // chars_per_line))  # ceiling division

    line_height_inches = font_height_inches * 1.35  # 1.35 = line spacing multiplier
    estimated_height = total_lines * line_height_inches

    log.debug(
        f"text='{text[:30]}...' | font={font_size_pt}pt | "
        f"box_w={box_width_inches:.2f}\" | chars/line={chars_per_line} | "
        f"lines={total_lines} | est_h={estimated_height:.3f}\""
    )
    return estimated_height


def will_overflow(text: str, font_size_pt: float,
                  box_width_inches: float, box_height_inches: float) -> bool:
    """Returns True if text will overflow the given box at given font size."""
    needed = estimate_text_height(text, font_size_pt, box_width_inches)
    overflows = needed > box_height_inches
    if overflows:
        log.warning(
            f"OVERFLOW DETECTED: needs {needed:.3f}\" but box is only {box_height_inches:.3f}\""
        )
    return overflows


# ─────────────────────────────────────────────────────────────────────────────
# CORE: Auto-Adjustment Strategies
# ─────────────────────────────────────────────────────────────────────────────

def shrink_font_to_fit(
    text: str,
    box_width_inches: float,
    box_height_inches: float,
    starting_font_pt: float = 14.0,
    min_font_pt: float = 7.0
) -> float:
    """
    Strategy 1 — Font Shrink Loop.

    Decreases font size by 0.5pt steps until text fits or min font is reached.

    Returns:
        The adjusted font size in points
    """
    font_pt = starting_font_pt

    while font_pt >= min_font_pt:
        if not will_overflow(text, font_pt, box_width_inches, box_height_inches):
            log.info(f"Font shrink resolved at {font_pt}pt (started at {starting_font_pt}pt)")
            return font_pt
        font_pt -= 0.5

    log.warning(
        f"Font shrink hit minimum ({min_font_pt}pt) — text may still overflow. "
        f"Consider expanding the box."
    )
    return min_font_pt


def expand_box_to_fit(
    text: str,
    font_size_pt: float,
    box_width_inches: float,
    current_box_height_inches: float,
    max_box_height_inches: float = 5.0
) -> float:
    """
    Strategy 2 — Box Height Expansion Fallback.

    Used when font shrink alone isn't enough (e.g. very long DB descriptions).
    Expands box height to fit text, capped at max_box_height_inches.

    Returns:
        The new box height in inches
    """
    needed = estimate_text_height(text, font_size_pt, box_width_inches)
    padding = 0.15  # small breathing room
    new_height = min(needed + padding, max_box_height_inches)

    if new_height > current_box_height_inches:
        log.info(
            f"Box expanded: {current_box_height_inches:.2f}\" → {new_height:.2f}\" "
            f"(max allowed: {max_box_height_inches:.2f}\")"
        )
    return new_height


# ─────────────────────────────────────────────────────────────────────────────
# SMART TEXT BOX — wraps all adjustment logic
# ─────────────────────────────────────────────────────────────────────────────

def add_smart_textbox(
    slide,
    text: str,
    x: float, y: float,          # position in Inches (raw float, not Emu)
    w: float, h: float,          # size in Inches (raw float)
    font_size: float = 14.0,
    bold: bool = False,
    color: RGBColor = RGBColor(0x33, 0x33, 0x33),
    bg_color: RGBColor = None,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    min_font: float = 7.0,
    allow_expand: bool = True,
    max_expand_h: float = 5.0,
    label: str = ""              # just for logging context
):
    """
    Adds a text box with automatic overflow handling.

    Feedback loop:
        1. Check if text fits at given font_size
        2. If not → shrink font (Strategy 1)
        3. If still overflows at min_font → expand box height (Strategy 2)
        4. Place the adjusted text box on the slide

    Args:
        slide:          python-pptx slide object
        text:           Text content (from your Oracle DB result)
        x, y:           Top-left position in inches
        w, h:           Width and height in inches
        font_size:      Desired font size in points
        bold:           Bold text
        color:          RGBColor for text
        bg_color:       Optional background fill color (RGBColor)
        align:          Text alignment
        min_font:       Smallest font size to shrink to
        allow_expand:   If True, expand box height when shrink isn't enough
        max_expand_h:   Max height box can grow to in inches
        label:          Human-readable name for logging (e.g. "Risk Description")
    """
    log.debug(f"--- Processing textbox: '{label}' ---")

    adjusted_font = font_size
    adjusted_h = h

    # ── Step 1: Try font shrink ───────────────────────────────────────────────
    if will_overflow(text, font_size, w, h):
        adjusted_font = shrink_font_to_fit(text, w, h, font_size, min_font)

    # ── Step 2: If still overflows at min font, expand box ───────────────────
    if allow_expand and will_overflow(text, adjusted_font, w, h):
        adjusted_h = expand_box_to_fit(text, adjusted_font, w, h, max_expand_h)

    # ── Step 3: Optional colored background shape ─────────────────────────────
    if bg_color:
        bg = slide.shapes.add_shape(
            1,  # MSO rectangle
            Inches(x), Inches(y), Inches(w), Inches(adjusted_h)
        )
        bg.fill.fore_color.rgb = bg_color
        bg.line.fill.background()  # no border

    # ── Step 4: Add the actual text box ──────────────────────────────────────
    txBox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(adjusted_h)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(adjusted_font)
    run.font.bold = bold
    run.font.color.rgb = color

    if adjusted_font != font_size or adjusted_h != h:
        log.info(
            f"'{label}' adjusted: font {font_size}pt→{adjusted_font}pt, "
            f"height {h:.2f}\"→{adjusted_h:.2f}\""
        )

    return txBox


# ─────────────────────────────────────────────────────────────────────────────
# SMART TABLE — detects cell overflow and shrinks font per-cell
# ─────────────────────────────────────────────────────────────────────────────

def add_smart_table(
    slide,
    headers: list,
    rows: list,           # list of dicts from Oracle query result
    col_keys: list,       # which dict keys map to which column
    x: float, y: float,
    w: float,
    row_height: float = 0.4,
    header_font: float = 11.0,
    cell_font: float = 10.0,
    min_font: float = 7.0,
    header_bg: RGBColor = RGBColor(0x00, 0x70, 0xC0),
    alt_row_bg: RGBColor = RGBColor(0xF2, 0xF7, 0xFF),
):
    """
    Adds a table with per-cell overflow detection.
    Column widths are distributed proportionally based on number of columns.

    Args:
        rows:     List of dicts — direct output of your Oracle cursor.fetchall()
        col_keys: e.g. ["risk_id", "description", "owner", "status"]
    """
    num_cols = len(headers)
    num_rows = len(rows) + 1  # +1 for header

    col_width = w / num_cols

    table_shape = slide.shapes.add_table(
        num_rows, num_cols,
        Inches(x), Inches(y),
        Inches(w), Inches(row_height * num_rows)
    )
    table = table_shape.table

    # ── Header row ────────────────────────────────────────────────────────────
    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = ""  # clear default

        para = cell.text_frame.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()

        # Check overflow in header cell
        if will_overflow(header_text, header_font, col_width, row_height):
            adjusted = shrink_font_to_fit(header_text, col_width, row_height,
                                          header_font, min_font)
            log.info(f"Header '{header_text}' font shrunk to {adjusted}pt")
        else:
            adjusted = header_font

        run.text = header_text
        run.font.size = Pt(adjusted)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        cell.fill.fore_color.rgb = header_bg

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row_idx, record in enumerate(rows):
        actual_row = row_idx + 1  # offset for header
        use_alt = (row_idx % 2 == 1)

        for col_idx, key in enumerate(col_keys):
            cell_text = str(record.get(key, ""))
            cell = table.cell(actual_row, col_idx)
            cell.text = ""

            para = cell.text_frame.paragraphs[0]
            run = para.add_run()

            # Per-cell overflow check
            if will_overflow(cell_text, cell_font, col_width, row_height):
                adjusted = shrink_font_to_fit(cell_text, col_width, row_height,
                                              cell_font, min_font)
                log.info(
                    f"Row {row_idx}, col '{key}': "
                    f"'{cell_text[:20]}...' shrunk to {adjusted}pt"
                )
            else:
                adjusted = cell_font

            run.text = cell_text
            run.font.size = Pt(adjusted)
            run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

            if use_alt:
                cell.fill.fore_color.rgb = alt_row_bg

    return table_shape


# ─────────────────────────────────────────────────────────────────────────────
# DEMO: Simulate Oracle DB output + generate full PSR slide
# ─────────────────────────────────────────────────────────────────────────────

def simulate_oracle_data() -> dict:
    """
    Mimics what cx_Oracle or oracledb cursor.fetchall() would return.
    Replace this with your actual FastAPI DB call.
    """
    return {
        # Short values — fit easily
        "project_name": "Project Phoenix – Core Banking Migration",
        "report_date": "15-May-2026",
        "status": "ON TRACK",

        # KPI boxes — some long, some short
        "kpis": [
            {"label": "Budget Used", "value": "$2.4M / $3.1M"},
            {"label": "Sprint Velocity", "value": "78 pts"},
            # Intentionally long to trigger overflow
            {"label": "Defects Open / Resolved / Deferred", "value": "14 / 203 / 7"},
            {"label": "Team Utilisation", "value": "91%"},
        ],

        # Risks table — description column will overflow
        "risks": [
            {
                "risk_id": "R-001",
                "description": "Third-party API vendor has not confirmed SLA for UAT environment availability",
                "owner": "Ankit S.",
                "status": "OPEN"
            },
            {
                "risk_id": "R-002",
                "description": "DB migration script performance untested on full prod dataset",
                "owner": "Priya M.",
                "status": "MITIGATING"
            },
            {
                "risk_id": "R-003",
                "description": "Security review pending sign-off",
                "owner": "Raj K.",
                "status": "CLOSED"
            },
        ]
    }


def generate_psr_pptx(db_data: dict) -> bytes:
    """
    Main generation function. Call this from your FastAPI endpoint.
    Returns raw bytes of the .pptx file.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank slide
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(0xF8, 0xF9, 0xFA)

    # ── HEADER BAND ───────────────────────────────────────────────────────────
    header_bg = slide.shapes.add_shape(
        1, Inches(0), Inches(0), Inches(13.33), Inches(0.9)
    )
    header_bg.fill.fore_color.rgb = RGBColor(0x00, 0x38, 0x6B)
    header_bg.line.fill.background()

    add_smart_textbox(
        slide,
        text=db_data["project_name"],
        x=0.3, y=0.1, w=9.0, h=0.6,
        font_size=22, bold=True,
        color=RGBColor(0xFF, 0xFF, 0xFF),
        label="Project Title"
    )

    add_smart_textbox(
        slide,
        text=f"Status Report  |  {db_data['report_date']}",
        x=9.5, y=0.1, w=3.5, h=0.6,
        font_size=12,
        color=RGBColor(0xAD, 0xD8, 0xE6),
        align=PP_ALIGN.RIGHT,
        label="Report Date"
    )

    # ── STATUS BADGE ─────────────────────────────────────────────────────────
    status_colors = {
        "ON TRACK": RGBColor(0x00, 0x96, 0x64),
        "AT RISK":  RGBColor(0xFF, 0x8C, 0x00),
        "OFF TRACK": RGBColor(0xCC, 0x00, 0x00),
    }
    badge_color = status_colors.get(db_data["status"], RGBColor(0x80, 0x80, 0x80))

    add_smart_textbox(
        slide,
        text=f"● {db_data['status']}",
        x=0.3, y=1.0, w=2.5, h=0.45,
        font_size=13, bold=True,
        color=badge_color,
        label="Status Badge"
    )

    # ── KPI BOXES (top area, 4 across) ───────────────────────────────────────
    kpi_x_start = 0.3
    kpi_y = 1.55
    kpi_w = 2.9
    kpi_h = 1.1
    kpi_gap = 0.25

    kpi_colors = [
        RGBColor(0x00, 0x70, 0xC0),
        RGBColor(0x00, 0x70, 0x64),
        RGBColor(0x7B, 0x2D, 0x8B),
        RGBColor(0xC0, 0x50, 0x00),
    ]

    for i, kpi in enumerate(db_data["kpis"]):
        kx = kpi_x_start + i * (kpi_w + kpi_gap)

        # Label (small, white, top of box)
        add_smart_textbox(
            slide,
            text=kpi["label"],
            x=kx, y=kpi_y,
            w=kpi_w, h=0.45,
            font_size=9,
            bold=False,
            color=RGBColor(0xFF, 0xFF, 0xFF),
            bg_color=kpi_colors[i],
            min_font=6.5,
            label=f"KPI label {i}"
        )

        # Value (big, bold, bottom of box)
        add_smart_textbox(
            slide,
            text=kpi["value"],
            x=kx, y=kpi_y + 0.45,
            w=kpi_w, h=0.65,
            font_size=18,
            bold=True,
            color=RGBColor(0xFF, 0xFF, 0xFF),
            bg_color=kpi_colors[i],
            min_font=8.0,
            label=f"KPI value {i}"
        )

    # ── SECTION LABEL ────────────────────────────────────────────────────────
    add_smart_textbox(
        slide,
        text="RISKS & ISSUES",
        x=0.3, y=3.0, w=4.0, h=0.35,
        font_size=10, bold=True,
        color=RGBColor(0x00, 0x38, 0x6B),
        label="Section Label"
    )

    # ── RISKS TABLE ──────────────────────────────────────────────────────────
    add_smart_table(
        slide,
        headers=["Risk ID", "Description", "Owner", "Status"],
        rows=db_data["risks"],
        col_keys=["risk_id", "description", "owner", "status"],
        x=0.3, y=3.4,
        w=12.7,
        row_height=0.45,
        header_font=11.0,
        cell_font=10.0,
    )

    # ── Save to bytes ─────────────────────────────────────────────────────────
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    log.info("PPTX generation complete.")
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# FastAPI endpoint (wire this into your existing app)
# ─────────────────────────────────────────────────────────────────────────────
"""
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
import io

app = FastAPI()

@app.get("/report/project-status")
def project_status_report():
    db_data = simulate_oracle_data()   # replace with your Oracle fetch
    pptx_bytes = generate_psr_pptx(db_data)
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=project_status_report.pptx"}
    )
"""

# ─────────────────────────────────────────────────────────────────────────────
# Run standalone (test in PyCharm without starting FastAPI server)
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    data = simulate_oracle_data()
    pptx_bytes = generate_psr_pptx(data)

    output_path = "project_status_report.pptx"
    with open(output_path, "wb") as f:
        f.write(pptx_bytes)

    print(f"\n✅ Saved: {output_path}")
    print("Open in PowerPoint or LibreOffice to verify layout.")
