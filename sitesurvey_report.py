#!/usr/bin/env python3
"""
SiteSurvey PDF Report Generator
Ingests an AppSheet CSV export + a photo folder and produces a styled PDF report.

Usage:
    python sitesurvey_report.py --csv surveys.csv --photos ./photos --output report.pdf
    python sitesurvey_report.py --csv surveys.csv --photos ./photos --title "Q2 Site Surveys" --id-col "Site Name"

Requirements:
    pip install reportlab pillow
"""

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path

# ── Dependency check ─────────────────────────────────────────────────────────

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        BaseDocTemplate, Frame, HRFlowable, Image,
        NextPageTemplate, PageBreak, PageTemplate,
        Paragraph, Spacer, Table, TableStyle, KeepTogether,
    )
except ImportError:
    print("ERROR: reportlab not installed.\nRun: pip install reportlab")
    sys.exit(1)

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("WARNING: pillow not installed — photos may not scale correctly.\nRun: pip install pillow")

# ── Layout constants ─────────────────────────────────────────────────────────

PAGE_W, PAGE_H = A4
MARGIN     = 2.0 * cm
CONTENT_W  = PAGE_W - 2 * MARGIN
HEADER_H   = 1.2 * cm
FOOTER_H   = 0.8 * cm

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

# ── Brand colours ─────────────────────────────────────────────────────────────

DARK   = colors.HexColor("#1e293b")
BLUE   = colors.HexColor("#2563eb")
LBLUE  = colors.HexColor("#dbeafe")
HINT   = colors.HexColor("#93c5fd")
GRAY   = colors.HexColor("#6b7280")
LGRAY  = colors.HexColor("#f8fafc")
GRID_C = colors.HexColor("#e2e8f0")
WHITE  = colors.white
BLACK  = colors.black


# ── Paragraph styles ─────────────────────────────────────────────────────────

def _s(name, **kw):
    return ParagraphStyle(name, **kw)

ST = {
    "title":     _s("title",    fontName="Helvetica-Bold",    fontSize=26, textColor=WHITE,  alignment=TA_CENTER),
    "subtitle":  _s("subtitle", fontName="Helvetica",         fontSize=12, textColor=HINT,   alignment=TA_CENTER),
    "stat_n":    _s("stat_n",   fontName="Helvetica-Bold",    fontSize=34, textColor=WHITE,  alignment=TA_CENTER),
    "stat_l":    _s("stat_l",   fontName="Helvetica",         fontSize=10, textColor=HINT,   alignment=TA_CENTER),
    "sec_hdr":   _s("sec_hdr",  fontName="Helvetica-Bold",    fontSize=13, textColor=WHITE),
    "rec_title": _s("rec_title",fontName="Helvetica-Bold",    fontSize=11, textColor=WHITE),
    "rec_num":   _s("rec_num",  fontName="Helvetica",         fontSize=9,  textColor=HINT,   alignment=TA_RIGHT),
    "lbl":       _s("lbl",      fontName="Helvetica-Bold",    fontSize=8,  textColor=GRAY,   spaceAfter=1),
    "val":       _s("val",      fontName="Helvetica",         fontSize=9,  textColor=BLACK,  spaceAfter=3),
    "caption":   _s("caption",  fontName="Helvetica-Oblique", fontSize=8,  textColor=GRAY,   alignment=TA_CENTER, spaceBefore=2),
    "small":     _s("small",    fontName="Helvetica",         fontSize=8,  textColor=GRAY),
    "ph_label":  _s("ph_label", fontName="Helvetica-Bold",    fontSize=10, textColor=BLUE,   spaceBefore=10, spaceAfter=4),
    "tbl_hdr":   _s("tbl_hdr",  fontName="Helvetica-Bold",    fontSize=8,  textColor=WHITE),
    "tbl_cell":  _s("tbl_cell", fontName="Helvetica",         fontSize=8,  textColor=BLACK),
}


# ── Document template ────────────────────────────────────────────────────────

def make_doc(out_path, report_title):
    def on_cover(canvas, _doc):
        canvas.saveState()
        canvas.setFillColor(BLUE);  canvas.rect(0, PAGE_H * 0.30, PAGE_W, PAGE_H * 0.70, fill=1, stroke=0)
        canvas.setFillColor(DARK);  canvas.rect(0, 0,             PAGE_W, PAGE_H * 0.30, fill=1, stroke=0)
        canvas.restoreState()

    def on_page(canvas, doc):
        canvas.saveState()
        # Top header bar
        canvas.setFillColor(DARK)
        canvas.rect(0, PAGE_H - HEADER_H, PAGE_W, HEADER_H, fill=1, stroke=0)
        canvas.setFillColor(WHITE);  canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(MARGIN, PAGE_H - 0.82 * cm, report_title)
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 0.82 * cm,
                               f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        # Bottom footer bar
        canvas.setFillColor(LGRAY)
        canvas.rect(0, 0, PAGE_W, FOOTER_H, fill=1, stroke=0)
        canvas.setFillColor(GRAY);   canvas.setFont("Helvetica", 8)
        canvas.drawCentredString(PAGE_W / 2, 0.25 * cm, f"Page {doc.page}")
        canvas.restoreState()

    body_y = MARGIN + FOOTER_H
    body_h = PAGE_H - 2 * MARGIN - HEADER_H - FOOTER_H

    doc = BaseDocTemplate(
        out_path, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN + HEADER_H, bottomMargin=MARGIN + FOOTER_H,
    )
    doc.addPageTemplates([
        PageTemplate(id="Cover",
                     frames=[Frame(0, 0, PAGE_W, PAGE_H, id="cf")],
                     onPage=on_cover),
        PageTemplate(id="Content",
                     frames=[Frame(MARGIN, body_y, CONTENT_W, body_h, id="body")],
                     onPage=on_page),
    ])
    return doc


# ── CSV helpers ──────────────────────────────────────────────────────────────

def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def detect_photo_cols(rows, photo_dir):
    """Return column names whose values appear to be image file references."""
    if not rows:
        return set()
    known_files = set()
    if photo_dir and Path(photo_dir).exists():
        known_files = {
            f.name.lower()
            for f in Path(photo_dir).rglob("*")
            if f.suffix.lower() in IMAGE_EXTS
        }
    found = set()
    for col in rows[0]:
        for row in rows[:30]:
            v = (row.get(col) or "").strip()
            if not v:
                continue
            if Path(v).suffix.lower() in IMAGE_EXTS or Path(v).name.lower() in known_files:
                found.add(col)
                break
    return found


def find_photo(value, photo_dir):
    """Locate a photo by filename within photo_dir, case-insensitive."""
    if not value or not photo_dir:
        return None
    base = Path(value).name
    # Exact name match first
    for p in Path(photo_dir).rglob(base):
        if p.is_file():
            return str(p)
    # Case-insensitive fallback
    for p in Path(photo_dir).rglob("*"):
        if p.name.lower() == base.lower() and p.is_file():
            return str(p)
    return None


# ── Image helper ─────────────────────────────────────────────────────────────

def fit_image(path, max_w, max_h):
    if PIL_AVAILABLE:
        with PILImage.open(path) as img:
            w, h = img.size
    else:
        w, h = 800, 600  # safe fallback if pillow not installed
    scale = min(max_w / w, max_h / h, 1.0)
    return Image(path, width=w * scale, height=h * scale)


# ── Reusable section bar ─────────────────────────────────────────────────────

def section_bar(text):
    t = Table([[Paragraph(text, ST["sec_hdr"])]], colWidths=[CONTENT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), DARK),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


# ── Cover page ───────────────────────────────────────────────────────────────

def build_cover(title, subtitle, n_records):
    stat_table = Table(
        [
            [Paragraph(str(n_records), ST["stat_n"]),
             Paragraph(datetime.now().strftime("%B %d, %Y"), ST["stat_n"])],
            [Paragraph("Records", ST["stat_l"]),
             Paragraph("Generated", ST["stat_l"])],
        ],
        colWidths=[PAGE_W * 0.38, PAGE_W * 0.38],
    )
    stat_table.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    return [
        Spacer(1, PAGE_H * 0.13),
        Paragraph("SiteSurvey Field Report", ST["subtitle"]),
        Spacer(1, 0.5 * cm),
        Paragraph(title, ST["title"]),
        Spacer(1, 0.4 * cm),
        Paragraph(subtitle, ST["subtitle"]),
        Spacer(1, PAGE_H * 0.10),
        stat_table,
        NextPageTemplate("Content"),
        PageBreak(),
    ]


# ── Summary table ────────────────────────────────────────────────────────────

def build_summary_table(rows, photo_cols, max_cols=8):
    els = [section_bar("Survey Summary — All Records"), Spacer(1, 0.3 * cm)]
    if not rows:
        return els + [Paragraph("No data.", ST["val"])]

    cols = [c for c in rows[0] if c not in photo_cols][:max_cols]
    col_w = CONTENT_W / len(cols)

    data = [[Paragraph(c, ST["tbl_hdr"]) for c in cols]]
    for row in rows:
        data.append([
            Paragraph(str(row.get(c, ""))[:60], ST["tbl_cell"])
            for c in cols
        ])

    t = Table(data, colWidths=[col_w] * len(cols), repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  BLUE),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LGRAY]),
        ("GRID",          (0, 0), (-1, -1), 0.4, GRID_C),
        ("ALIGN",         (0, 0), (-1, -1), "LEFT"),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
    ]))
    return els + [t, Spacer(1, 0.5 * cm)]


# ── Detail record ────────────────────────────────────────────────────────────

def build_record(row, idx, photo_cols, photo_dir, id_col):
    record_id = row.get(id_col, f"Record {idx + 1}") if id_col else f"Record {idx + 1}"
    els = []

    # ── Header bar ──
    hdr = Table(
        [[Paragraph(f"#{idx + 1}  {record_id}", ST["rec_title"]),
          Paragraph(f"Row {idx + 1}", ST["rec_num"])]],
        colWidths=[CONTENT_W - 3 * cm, 3 * cm],
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), DARK),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    # ── Fields in 2-column layout ──
    fields = [
        (k, str(v))
        for k, v in row.items()
        if k not in photo_cols and str(v or "").strip()
    ]
    half = CONTENT_W / 2 - 0.2 * cm
    field_rows = []
    for i in range(0, len(fields), 2):
        lk, lv = fields[i]
        rk, rv = fields[i + 1] if i + 1 < len(fields) else ("", "")
        field_rows.append([
            [Paragraph(lk, ST["lbl"]), Paragraph(lv[:200], ST["val"])],
            [Paragraph(rk, ST["lbl"]), Paragraph(rv[:200], ST["val"])],
        ])

    if field_rows:
        ft = Table(field_rows, colWidths=[half, half])
        ft.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("BACKGROUND",    (0, 0), (-1, -1), LGRAY),
            ("LINEBELOW",     (0, 0), (-1, -2), 0.4, GRID_C),
        ]))
        els.append(KeepTogether([hdr, ft]))
    else:
        els.append(hdr)

    # ── Photos ──
    cell_w = CONTENT_W / 2 - 0.4 * cm
    photo_cells = []

    for col in photo_cols:
        raw = (row.get(col) or "").strip()
        if not raw:
            continue
        # Support comma/newline-separated lists of filenames in one cell
        for pval in [v.strip() for v in raw.replace("\n", ",").split(",") if v.strip()]:
            fpath = find_photo(pval, photo_dir)
            if fpath:
                try:
                    img = fit_image(fpath, cell_w, 6 * cm)
                    cap = Paragraph(f"{col}: {Path(pval).name}", ST["caption"])
                    inner = Table([[img], [cap]], colWidths=[cell_w])
                    inner.setStyle(TableStyle([
                        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
                        ("TOPPADDING",    (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]))
                    photo_cells.append(inner)
                except Exception as exc:
                    photo_cells.append(
                        Paragraph(f"[Photo load error: {Path(pval).name} — {exc}]", ST["small"])
                    )
            else:
                photo_cells.append(
                    Paragraph(f"[Photo not found: {pval}]", ST["small"])
                )

    if photo_cells:
        els.append(Paragraph("Photos", ST["ph_label"]))
        grid_rows = []
        for i in range(0, len(photo_cells), 2):
            left  = photo_cells[i]
            right = photo_cells[i + 1] if i + 1 < len(photo_cells) else Spacer(1, 1)
            grid_rows.append([left, right])

        grid = Table(grid_rows, colWidths=[CONTENT_W / 2, CONTENT_W / 2])
        grid.setStyle(TableStyle([
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("GRID",          (0, 0), (-1, -1), 0.4, GRID_C),
            ("BACKGROUND",    (0, 0), (-1, -1), WHITE),
        ]))
        els.append(grid)

    els.append(HRFlowable(width="100%", thickness=1, color=LBLUE,
                          spaceAfter=10, spaceBefore=6))
    return els


# ── Entry point ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate a PDF report from an AppSheet CSV export and photo folder."
    )
    parser.add_argument("--csv",          required=True,        help="Path to AppSheet CSV export file")
    parser.add_argument("--photos",       default=None,         help="Folder containing photos (optional)")
    parser.add_argument("--output",       default="report.pdf", help="Output PDF path (default: report.pdf)")
    parser.add_argument("--title",        default="Site Survey Report", help="Report title")
    parser.add_argument("--subtitle",     default="",           help="Report subtitle (auto-generated if blank)")
    parser.add_argument("--id-col",       default=None,         help="Column to use as record label (auto-detected if blank)")
    parser.add_argument("--max-records",  type=int, default=None, help="Limit number of records (useful for testing)")
    parser.add_argument("--summary-cols", type=int, default=8,  help="Max columns shown in summary table (default: 8)")
    args = parser.parse_args()

    # ── Validate ──
    if not Path(args.csv).exists():
        print(f"ERROR: CSV file not found: {args.csv}")
        sys.exit(1)
    if args.photos and not Path(args.photos).exists():
        print(f"WARNING: Photos folder not found: {args.photos} — photos will be skipped")

    # ── Load data ──
    print(f"Loading: {args.csv}")
    rows = load_csv(args.csv)
    if args.max_records:
        rows = rows[:args.max_records]
    print(f"  {len(rows)} records loaded")

    photo_cols = detect_photo_cols(rows, args.photos)
    print(f"  Photo columns: {', '.join(photo_cols) if photo_cols else 'none detected'}")

    # Auto-detect a sensible record label column
    id_col = args.id_col
    if not id_col and rows:
        for candidate in ("ID", "id", "Row ID", "_RowNumber", "Survey ID",
                          "Name", "Site Name", "Site", "Customer"):
            if candidate in rows[0]:
                id_col = candidate
                print(f"  Using '{id_col}' as record label")
                break

    subtitle = args.subtitle or f"{len(rows)} records  •  {datetime.now().strftime('%Y-%m-%d')}"

    # ── Build PDF ──
    print(f"Building PDF: {args.output}")
    doc = make_doc(args.output, args.title)
    story = []
    story += build_cover(args.title, subtitle, len(rows))
    story += build_summary_table(rows, photo_cols, args.summary_cols)
    story.append(PageBreak())

    for i, row in enumerate(rows):
        print(f"  Record {i + 1}/{len(rows)}\r", end="", flush=True)
        story += build_record(row, i, photo_cols, args.photos, id_col)

    print()
    doc.build(story)
    size_kb = Path(args.output).stat().st_size // 1024
    print(f"Done → {args.output}  ({size_kb} KB)")


if __name__ == "__main__":
    main()
