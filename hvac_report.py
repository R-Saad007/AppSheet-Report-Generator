#!/usr/bin/env python3
"""
HVAC Preventive Maintenance — PDF Report Generator
Reads an AppSheet CSV export and produces one PDF per row, embedding any
photos referenced in the row.

By default each report is saved INSIDE the site's own folder (next to its
``images`` folder). Pass ``--output-dir`` to force a single output directory.

Usage:
    # Render every row, each PDF placed in its site folder.
    python hvac_report.py --csv "AppSheet Report Generation Data.csv"

    # Render a single row by 1-based index or Survey ID.
    python hvac_report.py --csv data.csv --row 1
    python hvac_report.py --csv data.csv --survey-id 9f8d15dc

    # Force all PDFs into a single folder.
    python hvac_report.py --csv data.csv --output-dir ./reports

Output files are always named:
    PM_Report_[SiteID]_[ReportDate].pdf       (date in ISO YYYY-MM-DD)

Install:
    pip install reportlab pillow
    pip install "qrcode[pil]"   # optional — for QR code on last page
"""

import argparse, csv, re, sys
from datetime import datetime, timedelta
from pathlib import Path
from io import BytesIO

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        BaseDocTemplate, CondPageBreak, Frame, HRFlowable, Image, KeepTogether,
        NextPageTemplate, PageBreak, PageTemplate,
        Paragraph, Spacer, Table, TableStyle,
    )
    from reportlab.platypus.flowables import Flowable
except ImportError:
    print("ERROR: pip install reportlab"); sys.exit(1)

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import qrcode as _qrc
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

# ── Photo helpers (constants used early in the module) ──────────────────────
PHOTO_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

# ── Page geometry ─────────────────────────────────────────────────────────────
PW, PH   = A4
ML = MR  = 1.5 * cm
MB       = 1.2 * cm
STRIPE   = 3
HDR_H    = 22
MT       = STRIPE + HDR_H + 8
CW       = PW - ML - MR

CODE_W   = 38
VAL_W    = 152
LBL_W    = CW - CODE_W - VAL_W

# ── Colour palette ────────────────────────────────────────────────────────────
H = colors.HexColor
COV_BG  = H("#1a2942")
ACCENT  = H("#3b82f6")
SEC_BG  = H("#1e3a5f")
SUB_BG  = H("#1d4ed8")
SUB_ACC = H("#60a5fa")
ROW_ALT = H("#eff6ff")
ROW_LN  = H("#e5e7eb")
CODEGH  = H("#9ca3af")
PASS_BG = H("#16a34a")
FAIL_BG = H("#dc2626")
NA_BG   = H("#6b7280")
YES_C   = H("#2563eb")
DARK    = H("#111827")
MID     = H("#374151")
LIGHT   = H("#6b7280")
HINT    = H("#93c5fd")
STATCARD= H("#243555")
TL_BG   = H("#dbeafe")
TL_TXT  = H("#1e40af")
HR_C    = H("#bfdbfe")
W       = colors.white

# ── Text styles ───────────────────────────────────────────────────────────────
def _s(n, **kw): return ParagraphStyle(n, **kw)
S = dict(
    c_tag  = _s("c_tag",  fontName="Helvetica",         fontSize=9,   textColor=HINT,  alignment=TA_CENTER, letterSpacing=2.5),
    c_ttl  = _s("c_ttl",  fontName="Helvetica-Bold",    fontSize=34,  textColor=W,     alignment=TA_CENTER, leading=40),
    c_site = _s("c_site", fontName="Helvetica-Bold",    fontSize=20,  textColor=W,     alignment=TA_CENTER, leading=24),
    c_cust = _s("c_cust", fontName="Helvetica",         fontSize=12,  textColor=HINT,  alignment=TA_CENTER, letterSpacing=2),
    c_dlbl = _s("c_dlbl", fontName="Helvetica",         fontSize=8,   textColor=HINT,  alignment=TA_CENTER, letterSpacing=2.5),
    c_date = _s("c_date", fontName="Helvetica-Bold",    fontSize=22,  textColor=W,     alignment=TA_CENTER, leading=26),
    c_stl  = _s("c_stl",  fontName="Helvetica",         fontSize=7,   textColor=HINT,  alignment=TA_CENTER, letterSpacing=1.8),
    c_stv  = _s("c_stv",  fontName="Helvetica-Bold",    fontSize=11,  textColor=W,     alignment=TA_CENTER, leading=14),
    c_foot = _s("c_foot", fontName="Helvetica",         fontSize=8,   textColor=HINT,  alignment=TA_CENTER),
    sec    = _s("sec",    fontName="Helvetica-Bold",    fontSize=14,  textColor=W,     leading=18),
    sec_s  = _s("sec_s",  fontName="Helvetica-Oblique", fontSize=9,   textColor=HINT,  leading=12, spaceBefore=6),
    sub    = _s("sub",    fontName="Helvetica-Bold",    fontSize=10.5,textColor=W,     leading=14, letterSpacing=0.5),
    code   = _s("code",   fontName="Helvetica",         fontSize=7.5, textColor=CODEGH),
    lbl    = _s("lbl",    fontName="Helvetica-Bold",    fontSize=9.5, textColor=DARK,  leading=13),
    val    = _s("val",    fontName="Helvetica",         fontSize=9.5, textColor=DARK,  alignment=TA_RIGHT, leading=13),
    val_b  = _s("val_b",  fontName="Helvetica-Bold",    fontSize=9.5, textColor=YES_C, alignment=TA_RIGHT, leading=13),
    val_d  = _s("val_d",  fontName="Helvetica",         fontSize=9.5, textColor=LIGHT, alignment=TA_RIGHT, leading=13),
    bdg    = _s("bdg",    fontName="Helvetica-Bold",    fontSize=8,   textColor=W,     alignment=TA_CENTER),
    tl_ev  = _s("tl_ev",  fontName="Helvetica-Bold",    fontSize=10,  textColor=DARK,  leading=13),
    tl_tm  = _s("tl_tm",  fontName="Helvetica",         fontSize=10,  textColor=DARK,  alignment=TA_RIGHT, leading=13),
    tl_tl  = _s("tl_tl",  fontName="Helvetica-Bold",    fontSize=10,  textColor=TL_TXT, leading=13),
    tl_tv  = _s("tl_tv",  fontName="Helvetica-Bold",    fontSize=10,  textColor=TL_TXT, alignment=TA_RIGHT, leading=13),
    cap    = _s("cap",    fontName="Helvetica",         fontSize=8.5, textColor=MID,   alignment=TA_CENTER, leading=11),
    notes  = _s("notes",  fontName="Helvetica",         fontSize=9.5, textColor=DARK,  alignment=TA_LEFT, leading=14),
    end_h  = _s("end_h",  fontName="Helvetica-Bold",    fontSize=14,  textColor=DARK,  alignment=TA_CENTER),
    end_b  = _s("end_b",  fontName="Helvetica",         fontSize=9,   textColor=MID,   alignment=TA_CENTER, leading=12),
    end_by = _s("end_by", fontName="Helvetica",         fontSize=10,  textColor=DARK,  alignment=TA_CENTER),
    ps_l   = _s("ps_l",   fontName="Helvetica-Bold",    fontSize=10,  textColor=W),
    ps_r   = _s("ps_r",   fontName="Helvetica",         fontSize=10,  textColor=W,     alignment=TA_RIGHT),
    qr_c   = _s("qr_c",   fontName="Helvetica-Oblique", fontSize=8,   textColor=LIGHT, alignment=TA_CENTER),
    unit_b = _s("unit_b", fontName="Helvetica-Bold",    fontSize=9,   textColor=W,     alignment=TA_CENTER),
)

# ── Custom flowables ──────────────────────────────────────────────────────────

class Badge(Flowable):
    """Rounded-rectangle badge."""
    def __init__(self, text, bg=PASS_BG, width=46, height=18):
        super().__init__()
        self.text = text; self.bg = bg
        self.bw = width; self.bh = height
        self.width = width; self.height = height

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(self.bg)
        c.roundRect(0, 0, self.bw, self.bh, radius=4, fill=1, stroke=0)
        c.setFillColor(W)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(self.bw / 2, 4, self.text)
        c.restoreState()


class TimelineDot(Flowable):
    def __init__(self, r=5):
        super().__init__()
        self.r = r; self.width = r * 2; self.height = r * 2

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(ACCENT)
        c.circle(self.r, self.r, self.r, fill=1, stroke=0)
        c.restoreState()


# ── Photo helpers ─────────────────────────────────────────────────────────────

def _column_segment_from_filename(name):
    """Extract the AppSheet column-name segment from a photo filename.

    AppSheet stores files as ``<id>.<column>.<timestamp>.<ext>`` and, when
    the column name is too long, truncates and appends ``_<hash>`` before
    the timestamp. Returns ``(column_segment, was_truncated)`` or
    ``(None, False)``.
    """
    stem = Path(name).stem
    # Strip trailing timestamp segment if it is purely numeric.
    if "." in stem:
        head, _, tail = stem.rpartition(".")
        if tail.isdigit():
            stem = head
    # Drop the leading id segment.
    if "." in stem:
        _, _, col_part = stem.partition(".")
    else:
        col_part = stem
    if not col_part:
        return None, False
    # AppSheet truncates and appends ``_<long-numeric-hash>`` for long
    # column names. Detect and trim that suffix.
    if "_" in col_part:
        head, _, tail = col_part.rpartition("_")
        if tail.isdigit() and len(tail) >= 10:
            return head, True
    return col_part, False


def find_photos_by_column(col, site_folder):
    """Return every file in ``site_folder`` whose name encodes ``col`` as
    its column segment — for any leading id, not just the row's Survey ID.

    Handles AppSheet's punctuation substitution (``? / \\`` → ``-``) and
    its truncation-with-hash suffix for long column names.
    """
    if not site_folder:
        return []
    sf = Path(site_folder)
    if not sf.is_dir():
        return []
    safe_col = col.translate(str.maketrans({"?": "-", "/": "-", "\\": "-"}))
    candidates = {col, safe_col}
    matches = []
    for p in sf.rglob("*"):
        if not p.is_file() or p.suffix.lower() not in PHOTO_EXTS:
            continue
        seg, truncated = _column_segment_from_filename(p.name)
        if not seg:
            continue
        if seg in candidates:
            matches.append(str(p))
            continue
        if truncated:
            # File was truncated; the segment is a prefix of the real column.
            if any(c.startswith(seg) for c in candidates):
                matches.append(str(p))
    matches.sort()
    return matches


def find_photo(value, photo_root, site_folder=None):
    """Resolve a CSV photo reference to an actual file on disk.

    Lookup order:
      1. Literal CSV path joined to ``photo_root``.
      2. Basename in ``site_folder`` (and its subfolders) if provided.
      3. Recursive basename search across ``photo_root``.
    """
    if not value:
        return None
    v = str(value).strip()
    if not v:
        return None

    # 1. Literal path (handles AppSheet's leading "/" and "//" separators).
    if photo_root:
        cleaned = v.lstrip("/").replace("//", "/")
        cand = Path(photo_root) / cleaned
        if cand.is_file():
            return str(cand)

    base = Path(v).name
    base_lower = base.lower()

    # 2. Search inside the site's folder first (faster + more reliable).
    if site_folder:
        sf = Path(site_folder)
        if sf.is_dir():
            for p in sf.rglob(base):
                if p.is_file():
                    return str(p)
            for p in sf.rglob("*"):
                if p.is_file() and p.name.lower() == base_lower:
                    return str(p)

    # 3. Fallback: recursive search of the photo root.
    if photo_root:
        root = Path(photo_root)
        for p in root.rglob(base):
            if p.is_file():
                return str(p)
        for p in root.rglob("*"):
            if p.is_file() and p.name.lower() == base_lower:
                return str(p)
    return None


def fit_image(path, max_w, max_h):
    if PIL_AVAILABLE:
        try:
            with PILImage.open(path) as img:
                ow, oh = img.size
        except Exception:
            ow, oh = 800, 600
    else:
        ow, oh = 800, 600
    scale = min(max_w / ow, max_h / oh, 1.0)
    return Image(path, width=ow * scale, height=oh * scale)


# ── Page templates ────────────────────────────────────────────────────────────

def make_doc(out_path, site_name):
    def on_cover(canvas, _doc):
        canvas.saveState()
        canvas.setFillColor(COV_BG)
        canvas.rect(0, 0, PW, PH, fill=1, stroke=0)
        canvas.setFillColor(ACCENT)
        canvas.rect(0, PH - 3, PW, 3, fill=1, stroke=0)
        canvas.restoreState()

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(ACCENT)
        canvas.rect(0, PH - STRIPE, PW, STRIPE, fill=1, stroke=0)
        canvas.setFillColor(W)
        canvas.rect(0, PH - STRIPE - HDR_H, PW, HDR_H, fill=1, stroke=0)
        canvas.setFillColor(MID)
        canvas.setFont("Helvetica", 8)
        canvas.drawString(ML, PH - STRIPE - HDR_H + 6, site_name)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawCentredString(PW / 2, PH - STRIPE - HDR_H + 6, "HVAC PREVENTIVE MAINTENANCE REPORT")
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(PW - MR, PH - STRIPE - HDR_H + 6, f"Page {doc.page}")
        canvas.setStrokeColor(ROW_ALT)
        canvas.setLineWidth(0.5)
        canvas.line(ML, PH - STRIPE - HDR_H, PW - MR, PH - STRIPE - HDR_H)
        canvas.restoreState()

    body_y = MB
    body_h = PH - MT - MB

    doc = BaseDocTemplate(
        out_path, pagesize=A4,
        leftMargin=ML, rightMargin=MR,
        topMargin=MT, bottomMargin=MB,
    )
    doc.addPageTemplates([
        PageTemplate(id="Cover",
                     frames=[Frame(0, 0, PW, PH, id="cf", leftPadding=0, rightPadding=0,
                                   topPadding=0, bottomPadding=0)],
                     onPage=on_cover),
        PageTemplate(id="Content",
                     frames=[Frame(ML, body_y, CW, body_h, id="body",
                                   leftPadding=0, rightPadding=0,
                                   topPadding=0, bottomPadding=0)],
                     onPage=on_page),
    ])
    return doc


# ── Building blocks ───────────────────────────────────────────────────────────

def section_header(title, subtitle=None, unit_label=None):
    """Dark navy section heading. ``subtitle`` renders below the title with
    breathing room; ``unit_label`` renders as a bright accent badge on the
    right-hand side (e.g. ``Unit 1 of 2``)."""
    if subtitle:
        left_content = [Paragraph(title, S["sec"]),
                        Paragraph(subtitle, S["sec_s"])]
        top_pad, bot_pad = 12, 14
    else:
        left_content = [Paragraph(title, S["sec"])]
        top_pad, bot_pad = 12, 12

    if unit_label:
        badge_w = 80
        left_w  = CW - badge_w
        # Inner cell wrapping the title (and optional subtitle) so we can
        # place a fixed-width badge alongside it without VALIGN issues.
        left_cell = Table([[left_content]], colWidths=[left_w])
        left_cell.setStyle(TableStyle([
            ("LEFTPADDING",   (0, 0), (-1, -1), 0),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
            ("TOPPADDING",    (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        badge_cell = Table([[Paragraph(unit_label, S["unit_b"])]], colWidths=[badge_w - 16])
        badge_cell.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), ACCENT),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ]))
        outer = Table([[left_cell, badge_cell]], colWidths=[left_w, badge_w - 16])
        outer.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), SEC_BG),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING",   (0, 0), (0, -1),  12),
            ("RIGHTPADDING",  (-1, 0), (-1, -1), 12),
            ("TOPPADDING",    (0, 0), (-1, -1), top_pad),
            ("BOTTOMPADDING", (0, 0), (-1, -1), bot_pad),
        ]))
        return [outer, Spacer(1, 8)]

    content = [[left_content if subtitle else left_content[0]]]
    t = Table(content, colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), SEC_BG),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("TOPPADDING",    (0, 0), (-1, -1), top_pad),
        ("BOTTOMPADDING", (0, 0), (-1, -1), bot_pad),
    ]))
    return [t, Spacer(1, 8)]


def subsection_header(title):
    """Blue subsection heading with a bright left accent bar."""
    accent = Table([[""]], colWidths=[5])
    accent.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), SUB_ACC),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    body = Table([[Paragraph(title, S["sub"])]], colWidths=[CW - 5])
    body.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), SUB_BG),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]))
    row = Table([[accent, body]], colWidths=[5, CW - 5])
    row.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return [row, Spacer(1, 6)]


def _normalize_value(raw):
    """Translate AppSheet booleans/keywords to display strings."""
    v = str(raw or "").strip()
    upper = v.upper()
    if upper == "TRUE":  return "Yes"
    if upper == "FALSE": return "No"
    if upper == "NA":    return "N/A"
    return v


def _value_cell(raw, suffix="", required=False):
    """Return a flowable for a value.

    When ``required`` is True (show-all mode), empty cells render as "N/A"
    instead of being suppressed.  ``suffix`` (e.g. ``"V"``, ``"°C"``) is
    appended to numeric/text values only — not to PASS/FAIL/Yes/No badges.
    """
    v = _normalize_value(raw)
    if v == "":
        if required:
            return Badge("N/A", bg=NA_BG, width=38)
        return None
    upper = v.upper()
    if upper == "PASS":
        return Badge("PASS", bg=PASS_BG)
    if upper == "FAIL":
        return Badge("FAIL", bg=FAIL_BG)
    if upper == "N/A":
        return Badge("N/A", bg=NA_BG, width=38)
    if v == "Yes":
        return Paragraph("Yes", S["val_b"])
    if v == "No":
        return Paragraph("No", S["val"])
    if suffix:
        v = f"{v} {suffix}"
    return Paragraph(v, S["val"])


def field_rows(fields, data, start_alt=False, required=True):
    """Build alternating field rows.

    When ``required`` is True (default) every field is always shown; empty
    values render as an N/A badge.  Set ``required=False`` to silently drop
    fields whose value is absent.

    Each field tuple is ``(code, label, csv_col)`` or
    ``(code, label, csv_col, unit_suffix)``.

    Codes shown to the reader are renumbered sequentially within the
    subsection after empties are dropped — so a section with codes
    ``1.1.1``..``1.1.10`` whose 1.1.7 is empty still renders as
    ``1.1.1``..``1.1.9`` without gaps in the visible numbering.
    """
    rendered = []
    for spec in fields:
        if len(spec) == 4:
            code, label, col, suffix = spec
        else:
            code, label, col = spec
            suffix = ""
        cell = _value_cell(data.get(col, ""), suffix=suffix, required=required)
        if cell is None:
            continue
        rendered.append([code, label, cell])

    if rendered:
        first_code = rendered[0][0]
        if first_code and "." in first_code:
            prefix = first_code.rsplit(".", 1)[0] + "."
            for n, item in enumerate(rendered, start=1):
                item[0] = f"{prefix}{n}"

    rows = []
    for i, (code, label, cell) in enumerate(rendered):
        bg = ROW_ALT if (i + start_alt) % 2 else W
        row = Table(
            [[Paragraph(code, S["code"]), Paragraph(label, S["lbl"]), cell]],
            colWidths=[CODE_W, LBL_W, VAL_W],
        )
        row.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), bg),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN",         (2, 0), (2, 0),   "RIGHT"),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING",   (0, 0), (0, 0),   6),
            ("LEFTPADDING",   (1, 0), (1, 0),   2),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("LINEBELOW",     (0, 0), (-1, -1), 0.3, ROW_LN),
        ]))
        rows.append(row)
    return rows


def field_subsection(title, fields, data, required=True):
    """Subsection header + rows; returns ``[]`` only when every field is
    suppressed (required=False and all values empty)."""
    rows = field_rows(fields, data, required=required)
    if not rows:
        return []
    header = subsection_header(title)
    first_block = KeepTogether(header + [rows[0]])
    return [Spacer(1, 10), first_block] + rows[1:] + [Spacer(1, 10)]


_PHOTO_COL_TITLE = {
    # Survey-level photos
    "Controller Picture":                                    "Controller",
    "Serial Number Picture":                                 "Serial Number",
    "Software Version Picture":                              "Software Version",
    "Firmware Version Picture":                              "Firmware Version",
    "Gateway Picture":                                       "Gateway",
    "ASLLC Picture":                                         "ASLLC",
    "Controller Serial (ASLLC) picture":                     "ASLLC Serial",
    "Gateway Serial picture":                                "Gateway Serial",
    "Gateway Model picture":                                 "Gateway Model",
    "Gateway Software Version picture":                      "Gateway SW Version",
    "Firmware Version ASLLC Picture":                        "ASLLC Firmware",
    "TCU-NG2 Tag Confirmation picture":                      "TCU-NG2 Tag",
    "TCU-NG2 Tag before":                                    "TCU-NG2 Tag (Before)",
    "TCU-NG2 Tag after":                                     "TCU-NG2 Tag (After)",
    "TCU800-MINI Tag Confirmation picture":                  "TCU800-MINI Tag",
    "TCU800-MINI Tag before":                                "TCU800-MINI Tag (Before)",
    "TCU800-MINI Tag after":                                 "TCU800-MINI Tag (After)",
    "ASLLC Controller Tag Confirmation picture":             "ASLLC Controller Tag",
    "ASLLC Controller Tag before":                           "ASLLC Controller Tag (Before)",
    "ASLLC Controller Tag after":                            "ASLLC Controller Tag (After)",
    "Gateway Tag Confirmation picture":                      "Gateway Tag",
    "Gateway Tag before":                                    "Gateway Tag (Before)",
    "Gateway Tag after":                                     "Gateway Tag (After)",
    "Added Tags for Controller (if applicable) picture":     "Controller Tag",
    "Added Tags for Controller before":                      "Controller Tag (Before)",
    "Added Tags for Controller after":                       "Controller Tag (After)",
    "Active Alarms Present picture":                         "Active Alarms",
    "Alarm History Checked & Cleared picture":               "Alarm History",
    "Controller System current Mode picture":                "System Mode",
    "Controller Operating Mode picture":                     "Operating Mode",
    "Input DC Voltage (Controller) picture":                 "DC Voltage (Controller)",
    "Input DC Voltage (Gateway) picture":                    "DC Voltage (Gateway)",
    "Temp Sensor Photo":                                     "Temp Sensor",
    "Humidity Sensor Photo":                                 "Humidity Sensor",
    "Indoor Temperature (on controller) picture":            "Indoor Temp",
    "Site Overview Photo":                                   "Site Overview",
    "Log Book Entry Photo":                                  "Log Book Entry",
    "Log Book Sign-Out Photo":                               "Log Book Sign-Out",
    # HVAC unit — before state
    "Make / Model Photo":                                    "Make / Model",
    "Serial Number Photo":                                   "Serial Number",
    "Dial Position (Should be Max) Photo":                   "Dial Position",
    "Relay & Connection Condition Photo":                    "Relay & Connection",
    "Relay Secured with Cable Tie Photo":                    "Cable Tie (Before)",
    "Condition of Supply and Return Grills Photo":           "Supply/Return Grills",
    "Supply Air Temp Sensor Condition Photo":                "Supply Air Temp Sensor",
    "HVAC Paint & Label Condition Photo":                    "Paint & Label",
    "Outdoor Temp Sensor Box (Inside & Out) Photo":          "Outdoor Temp Sensor",
    "Main Air Filter (Note Date & Type, MUST be MERV 8+) Photo": "Main Air Filter",
    "Pre-Filter Condition / Type Photo":                     "Pre-Filter",
    "Condenser Coil Condition Photo":                        "Condenser Coil",
    "Evaporator Coil Condition Photo":                       "Evaporator Coil",
    "System Wiring Free from Wear/Damage? Photo":            "System Wiring",
    "Breakers & Wiring Correctly Sized for Load? Photo":     "Breakers & Wiring",
    "DAMPER Actuator Motor in Correct Position & Secured? Photo": "Damper Actuator",
    "Bug Screen on Exhaust? Photo":                          "Bug Screen (Exhaust)",
    "Bug Screen on Side Panel? Photo":                       "Bug Screen (Side Panel)",
    "Bug Screen on Evaporator Drain Hole? Photo":            "Bug Screen (Evap Drain)",
    "Gasket Added on Pre-Filter? Photo":                     "Gasket on Pre-Filter",
    "Condenser Fan Condition Photo":                         "Condenser Fan",
    "Supply Fan(s) Condition Photo":                         "Supply Fan(s)",
    "Compressor Compartment Cleanliness Photo":              "Compressor Compartment",
    "Evaporator Compartment Cleanliness Photo":              "Evaporator Compartment",
    "Condenser Compartment Cleanliness Photo":               "Condenser Compartment",
    # HVAC unit — maintenance
    "Replaced Main Air Filter & Labeled (Date, Location, MERV)? Photo": "Main Filter Replaced",
    "Replaced Nylon Mesh with Aluminum Mesh (if applicable)? Photo":     "Mesh Replaced",
    "Cleaned Aluminum Mesh Pre-filter? Photo":               "Pre-Filter Cleaned",
    "Added Gasket on Pre-filter Slot? Photo":                "Gasket Added",
    "Added Bug Screen on Exhaust Grill? Photo":              "Bug Screen (Exhaust) Added",
    "Added Bug Screen on Evap Coil Drain? Photo":            "Bug Screen (Evap Drain) Added",
    "Added Bug Screen on Left Panel? Photo":                 "Bug Screen (Left Panel) Added",
    "Cleaned Evap Coil & Compartment? Photo":                "Evap Coil Cleaned",
    "Cleaned Compressor Compartment? Photo":                 "Compressor Compartment Cleaned",
    "Cleaned Condenser Coil (cleaning agent + water)? Photo":"Condenser Coil Cleaned",
    "Cleaned Drainage Pipe? Photo":                          "Drainage Pipe Cleaned",
    "Lubricated Supply Fan? Photo":                          "Supply Fan Lubricated",
    "Lubricated Condenser Fan? Photo":                       "Condenser Fan Lubricated",
    "Straightened Coil Fins? Photo":                         "Coil Fins Straightened",
    "Replaced Bent ABS Fins on Grills? Photo":               "ABS Fins Replaced",
    "Replaced Rusted Screws (steel) & Freed Screws (self-tapping)? Photo": "Screws Replaced",
    "Grinded & Spray-Painted Rusted Areas? Photo":           "Rust Treated",
    "Confirmed Damper Actuator Motor is Secure/Correct? Photo": "Damper Actuator Confirmed",
    "Performed Leak Test for R410A Refrigerant Throughout the HVAC? Photo": "Leak Test",
    "Dial Position (Should be Max) [Post-Maintenance] Photo":"Dial Position (Post-Maint.)",
    "Secure Connection Condition Photo":                     "Secure Connection",
    "Relay Secured with Cable Tie [Post-Maintenance] Photo": "Cable Tie (Post-Maint.)",
    "Added Tags for HVAC Unit (if applicable)? Photo":       "HVAC Unit Tag",
    # HVAC unit — testing
    "Supply Fan — Amps Photo":                               "Supply Fan Amps",
    "Supply Fan — Volts Photo":                              "Supply Fan Volts",
    "Compressor + Condenser Fan — Amps Photo":               "Comp+Cond Fan Amps",
    "Compressor + Condenser Fan — Volts Photo":              "Comp+Cond Fan Volts",
    "Heater — Amps Photo":                                   "Heater Amps",
    "Heater — Volts Photo":                                  "Heater Volts",
    "Discharge Pressure (Without Compressor) Photo":         "Discharge Pressure (Off)",
    "Discharge Pressure (With Compressor) Photo":            "Discharge Pressure (On)",
    "High Pressure Switch Operating Above 400 PSI Photo":    "High Pressure Switch",
    "Suction Pressure (Without Compressor) Photo":           "Suction Pressure (Off)",
    "Suction Pressure (With Compressor) Photo":              "Suction Pressure (On)",
    # Final inspection
    "Compressor terminal condition photo":                   "Compressor Terminal",
    "HVAC Breaker type in AC panel photo":                   "HVAC Breaker",
    "HVAC Manuals on Site? Photo":                           "HVAC Manuals",
    "Thermostat Set to Auto and Reading Accurately? Photo":  "Thermostat",
    "Grills Photo":                                          "Grills",
    "HVAC Controller Indicates No Outstanding Alarms? Photo":"No Outstanding Alarms",
    # Shared system testing
    "Lag HVAC Provides Cooling/Heating if Lead Fails/Demo Photo": "Failover Test",
    "DC Fan Failover Works? Photo":                          "DC Fan Failover",
    "Temp Controls Maintaining Heat >+11°C and Cooling <+27°C? Photo": "Temp Controls",
    "Primary Indoor Temp Sensor Clean & Working? Photo":     "Primary Temp Sensor",
    "Humidity Sensor Clean & Working? Photo":                "Humidity Sensor",
    "External Temp Sensor Clean of Debris & Working? Photo": "External Temp Sensor",
    "Air Supply Sensors for WPU1 & WPU2 Clean & Working? Photo": "Air Supply Sensors",
    "System Clock Has the Correct Time? Photo":              "System Clock",
    "Controller Correctly Logging Errors/History (up to 6 months / 500 entries)? Photo": "Error Logging",
    "Individual Units Fail to Report Alarms Correctly? (TCU800-MINI only) Photo": "Unit Alarm Reporting",
    "High/Low Temp Alarms Reporting Correctly? Photo":       "High/Low Temp Alarms",
    "Dirty Filter Condition Reporting Correctly? Photo":     "Dirty Filter Alarm",
    "Refrigerant Low Pressure Reporting Correctly? Photo":   "Refrigerant Low Pressure Alarm",
    # Setpoints
    "Lead Compressor Start (ASLLC) Photo":                   "Lead Comp. Start (ASLLC)",
    "Compressor Stop (ASLLC) Photo":                         "Comp. Stop (ASLLC)",
    "Lag Compressor Start (ASLLC) Photo":                    "Lag Comp. Start (ASLLC)",
    "Lag Compressor Stop (ASLLC) Photo":                     "Lag Comp. Stop (ASLLC)",
    "Lead Heater Start (ASLLC) Photo":                       "Lead Heater Start (ASLLC)",
    "Heaters Stop (ASLLC) Photo":                            "Heaters Stop (ASLLC)",
    "Lag Heater Start (ASLLC) Photo":                        "Lag Heater Start (ASLLC)",
    "Lag Heater Stop (ASLLC) Photo":                         "Lag Heater Stop (ASLLC)",
    "Supply Fans Start (ASLLC) Photo":                       "Supply Fans Start (ASLLC)",
    "Supply Fan Stop (ASLLC) Photo":                         "Supply Fan Stop (ASLLC)",
    "High Temp Alarm (ASLLC) Photo":                         "High Temp Alarm (ASLLC)",
    "Low Temp Alarm (ASLLC) Photo":                          "Low Temp Alarm (ASLLC)",
    "Lead Compressor Start (TCU-NG2) Photo":                 "Lead Comp. Start (TCU-NG2)",
    "Lead Compressors Stop (TCU-NG2) Photo":                 "Comp. Stop (TCU-NG2)",
    "Lag Compressor Start (TCU-NG2) Photo":                  "Lag Comp. Start (TCU-NG2)",
    "Lag Compressor Stop (TCU-NG2) Photo":                   "Lag Comp. Stop (TCU-NG2)",
    "Lead Heater Start (TCU-NG2) Photo":                     "Lead Heater Start (TCU-NG2)",
    "Heaters Stop (TCU-NG2) Photo":                          "Heaters Stop (TCU-NG2)",
    "Lag Heater Start (TCU-NG2) Photo":                      "Lag Heater Start (TCU-NG2)",
    "Lag Heater Stop (TCU-NG2) Photo":                       "Lag Heater Stop (TCU-NG2)",
    "Supply Fans Start (TCU-NG2) Photo":                     "Supply Fans Start (TCU-NG2)",
    "Supply Fan Stop (TCU-NG2) Photo":                       "Supply Fan Stop (TCU-NG2)",
    "High Temp Alarm (TCU-NG2) Photo":                       "High Temp Alarm (TCU-NG2)",
    "Low Temp Alarm (TCU-NG2) Photo":                        "Low Temp Alarm (TCU-NG2)",
    "Lead Compressor Start Photo":                           "Lead Comp. Start",
    "Compressors Stop Photo":                                "Comp. Stop",
    "Lag Compressor Start Photo":                            "Lag Comp. Start",
    "Lag Compressor Stop Photo":                             "Lag Comp. Stop",
    "Lead Heater Start Photo":                               "Lead Heater Start",
    "Heaters Stop Photo":                                    "Heaters Stop",
    "Lag Heater Start Photo":                                "Lag Heater Start",
    "Lag Heater Stop Photo":                                 "Lag Heater Stop",
    "Supply Fans Start Photo":                               "Supply Fans Start",
    "Supply Fan Stop Photo":                                 "Supply Fan Stop",
    "High Temp Alarm Photo":                                 "High Temp Alarm",
    "Low Temp Alarm Photo":                                  "Low Temp Alarm",
    # Failover test photos
    "Failover Test Photo 1":                                 "Failover Test 1",
    "Failover Test Photo 2":                                 "Failover Test 2",
    "Failover Test Photo 3":                                 "Failover Test 3",
}


def _photo_title(col):
    """Return a short human-readable title for a photo column."""
    if col in _PHOTO_COL_TITLE:
        return _PHOTO_COL_TITLE[col]
    # Strip trailing "Photo" / "photo" / "picture" suffixes and numbered
    # suffixes like "Site Overview Photo 2".
    title = re.sub(r'\s+(photo|picture)\s*\d*$', '', col, flags=re.IGNORECASE).strip()
    title = re.sub(r'\s*\?\s*$', '', title).strip()
    return title or col


def _photo_grid_rows(photo_cols, data, photo_root, site_folder=None,
                     max_h_cm=6.5, used=None):
    """Return a list of single-row Table flowables (one row = up to 2 photos).

    Splitting the grid into one flowable per row lets ReportLab page-break
    cleanly between rows and lets us wrap the header + first row in
    ``KeepTogether`` to prevent orphan headers.

    ``photo_cols`` is a list of CSV column names; for each, the column's
    value is used to resolve the file. Falls back to a filename-pattern
    search inside ``site_folder`` when the CSV cell is empty so a photo
    sitting on disk still appears in the report. Every successfully-rendered
    file path is recorded in ``used`` (a mutable set) so the catch-all
    "Additional Photos" section can skip duplicates.
    """
    cells = []
    cell_w = CW / 2 - 6
    for col in photo_cols:
        # Resolve the file via the CSV-referenced path. Each row (survey
        # or HVAC unit) carries its own values, so per-unit photos come
        # from the unit's row, not from a cross-folder pattern search.
        seen_paths = []
        val = (data.get(col, "") or "").strip()
        if val:
            fpath = find_photo(val, photo_root, site_folder)
            if fpath:
                seen_paths.append(fpath)

        for fpath in seen_paths:
            try:
                img = fit_image(fpath, cell_w, max_h_cm * cm)
                cap = Paragraph(_photo_title(col), S["cap"])
                inner = Table([[img], [cap]], colWidths=[cell_w])
                inner.setStyle(TableStyle([
                    ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
                    ("TOPPADDING",    (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]))
                cells.append(inner)
                if used is not None:
                    try:
                        used.add(str(Path(fpath).resolve()))
                    except Exception:
                        used.add(fpath)
            except Exception:
                continue

    if not cells:
        return []

    rows = []
    for i in range(0, len(cells), 2):
        left  = cells[i]
        right = cells[i + 1] if i + 1 < len(cells) else Spacer(1, 1)
        row = Table([[left, right]], colWidths=[CW / 2, CW / 2])
        row.setStyle(TableStyle([
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        rows.append(row)
    return rows


def photo_subsection(title, photo_cols, data, photo_root,
                     site_folder=None, max_h_cm=6.5, used=None):
    rows = _photo_grid_rows(photo_cols, data, photo_root,
                            site_folder=site_folder, max_h_cm=max_h_cm,
                            used=used)
    if not rows:
        return []
    header = subsection_header(title)
    # Bind header + first row together so the header never strands on the
    # previous page when photos flow to the next.
    first_block = KeepTogether(header + [rows[0]])
    return [Spacer(1, 10), first_block] + rows[1:] + [Spacer(1, 10)]


def sp(h=6):
    return Spacer(1, h)


# ── Date / value formatting ───────────────────────────────────────────────────

def iso_date(s):
    s = (s or "").strip()
    for fmt in ("%m/%d/%Y", "%-m/%-d/%Y", "%Y-%m-%d", "%m-%d-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def _parse_time_of_day(s):
    """Parse 'HH:MM:SS' or 'HH:MM' as a datetime; ``None`` if unparseable."""
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def duration_str(start, end):
    """Return ``"X hours and Y minutes"`` between two HH:MM[:SS] strings.

    If the end is earlier than the start (technician worked past midnight),
    the duration wraps to the next day. Returns ``""`` if either input is
    unparseable.
    """
    s = _parse_time_of_day(start)
    e = _parse_time_of_day(end)
    if not s or not e:
        return ""
    delta = e - s
    if delta.total_seconds() < 0:
        delta += timedelta(days=1)
    minutes = int(delta.total_seconds() // 60)
    h, m = divmod(minutes, 60)
    return f"{h} hours and {m} minutes"


# ── Cover page ────────────────────────────────────────────────────────────────

def _stat_card(label, value, width):
    content = [
        [Paragraph(label, S["c_stl"])],
        [Paragraph(str(value or "—"), S["c_stv"])],
    ]
    t = Table(content, colWidths=[width])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), STATCARD),
        ("TOPPADDING",    (0, 0), (0, 0),   10),
        ("BOTTOMPADDING", (0, 0), (0, 0),   4),
        ("TOPPADDING",    (0, 1), (0, 1),   2),
        ("BOTTOMPADDING", (0, 1), (0, 1),   12),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("LINEABOVE",     (0, 0), (-1, 0),  2, ACCENT),
    ]))
    return t


def build_cover(data):
    site    = data.get("Site Name", "") or "—"
    cust    = (data.get("Customer", "") or "").upper()
    date    = iso_date(data.get("Report Date", "")) or datetime.now().strftime("%Y-%m-%d")
    tech    = (data.get("Technician", "") or "—").strip()
    units   = data.get("Number of HVAC Units", "") or "—"
    srv_id  = data.get("Survey ID", "") or "—"
    location = data.get("Location Code", "") or "—"

    tech_disp = tech.split("@", 1)[0] if "@" in tech else tech

    gen_at = datetime.now().strftime("Generated %b %d, %Y · %I:%M %p").replace("AM", "a.m.").replace("PM", "p.m.")

    card_w = (PW - 2 * ML) / 4 - 6
    cards  = [[
        _stat_card("TECHNICIAN",    tech_disp, card_w),
        _stat_card("HVAC UNITS",    units,     card_w),
        _stat_card("LOCATION CODE", location,  card_w),
        _stat_card("SURVEY ID",     srv_id,    card_w),
    ]]
    cards_tbl = Table(cards, colWidths=[card_w + 6] * 4)
    cards_tbl.setStyle(TableStyle([
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]))

    els = [
        sp(PH * 0.135),
        Paragraph("HVAC PREVENTIVE MAINTENANCE", S["c_tag"]),
        HRFlowable(width="55%", thickness=1.5, color=ACCENT, spaceAfter=14, spaceBefore=10),
        Paragraph("FIELD SURVEY REPORT", S["c_ttl"]),
        sp(34),
        Paragraph(site, S["c_site"]),
    ]
    if cust:
        els += [sp(26), Paragraph(cust, S["c_cust"])]
    els += [
        sp(PH * 0.16),
        Paragraph("REPORT DATE", S["c_dlbl"]),
        sp(8),
        Paragraph(date, S["c_date"]),
        sp(PH * 0.21),
        cards_tbl,
        sp(34),
        Paragraph("PLC Group · HVAC Preventive Maintenance", S["c_foot"]),
        sp(3),
        Paragraph(gen_at, S["c_foot"]),
        NextPageTemplate("Content"),
        PageBreak(),
    ]
    return els


# ── Activity Timeline ─────────────────────────────────────────────────────────

def build_timeline(data):
    arrival = (data.get("Time of Arrival", "") or "").strip()
    finish  = (data.get("Finish Time", "") or "").strip()
    total   = duration_str(arrival, finish)
    if not total:
        # Fallback to AppSheet's pre-computed value when we can't parse.
        total = (data.get("Total Survey Time", "") or "").strip()

    events = [(label, t) for label, t in
              (("Time of Arrival", arrival), ("Finish Time", finish))
              if t]

    if not events and not total:
        return []

    els = section_header("ACTIVITY TIMELINE")

    if events:
        tbl_data = [
            [TimelineDot(),
             Paragraph(label, S["tl_ev"]),
             Paragraph(t, S["tl_tm"])]
            for label, t in events
        ]
        tbl = Table(tbl_data, colWidths=[18, LBL_W + CODE_W - 18, VAL_W])
        tbl.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 9),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
            ("LEFTPADDING",   (0, 0), (0, 0),   6),
            ("LEFTPADDING",   (1, 0), (1, 0),   6),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("LINEBELOW",     (0, 0), (-1, -1), 0.3, ROW_LN),
        ]))
        els.append(tbl)

    if total:
        total_row = Table(
            [[Paragraph("Total time taken", S["tl_tl"]),
              Paragraph(total, S["tl_tv"])]],
            colWidths=[LBL_W + CODE_W, VAL_W],
        )
        total_row.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), TL_BG),
            ("TOPPADDING",    (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("LEFTPADDING",   (0, 0), (-1, -1), 12),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ]))
        els += [sp(4), total_row]

    els += [sp(14)]
    return els


# ── Survey Overview ───────────────────────────────────────────────────────────

def build_survey_overview(data, photo_root, site_folder, used=None):
    els = section_header("SURVEY OVERVIEW")

    overview_fields = [
        ("1.1.1",  "Customer",             "Customer"),
        ("1.1.2",  "Site Name",            "Site Name"),
        ("1.1.3",  "Site Coordinates",     "Site Coordinates"),
        ("1.1.4",  "Location Code",        "Location Code"),
        ("1.1.5",  "Report Date",          "Report Date"),
        ("1.1.6",  "Time of Arrival",      "Time of Arrival"),
        ("1.1.7",  "Technician",           "Technician"),
        ("1.1.8",  "Contact Person",       "Contact Person"),
        ("1.1.9",  "Start Location (GPS)", "Start Location"),
        ("1.1.10", "Travel Distance (km)", "Travel Distance"),
        ("1.1.11", "Travel Time",          "Travel Time"),
        ("1.1.12", "Submitted At",         "Submitted At"),
        ("1.1.13", "Number of HVAC Units", "Number of HVAC Units"),
    ]
    els += field_rows(overview_fields, data)

    site_photo_cols = ["Site Overview Photo"] + [f"Site Overview Photo {i}" for i in range(2, 7)]
    els += photo_subsection("Site Pictures", site_photo_cols,
                            data, photo_root, site_folder=site_folder, used=used)

    els += photo_subsection("Log Book Entry",
                            ["Log Book Entry Photo"],
                            data, photo_root, site_folder=site_folder, used=used)
    return els


# ── System Identification ─────────────────────────────────────────────────────

def build_system_id(data, photo_root, site_folder, used=None):
    ctype = (data.get("Controller Type", "") or "").strip()
    els = section_header("SYSTEM IDENTIFICATION",
                         subtitle=f"Controller: {ctype}" if ctype else None)

    # Merge firmware version: TCU-NG2 stores it in "Firmware Version";
    # ASLLC_TCU800MINI stores it in "Firmware Version ASLLC". Show whichever
    # is populated, preferring the non-ASLLC column when both have values.
    data = dict(data)
    if not (data.get("Firmware Version") or "").strip():
        data["Firmware Version"] = data.get("Firmware Version ASLLC", "")
    if not (data.get("Firmware Version Picture") or "").strip():
        data["Firmware Version Picture"] = data.get("Firmware Version ASLLC Picture", "")

    els += field_subsection("Controller Information", [
        ("2.1.1", "Controller Type",     "Controller Type"),
        ("2.1.2", "Serial Number",       "Serial Number"),
        ("2.1.3", "Software Version",    "Software Version"),
        ("2.1.4", "Firmware Version",    "Firmware Version"),
        ("2.1.5", "Number of HVAC Units","Number of HVAC Units"),
    ], data)
    els += photo_subsection("Controller Photos", [
        "Controller Picture",
        "Serial Number Picture",
        "Software Version Picture",
        "Firmware Version Picture",
    ], data, photo_root, site_folder=site_folder, used=used)

    # ASLLC subsystem block (only present for ASLLC_TCU800MINI sites).
    els += field_subsection("ASLLC Subsystem", [
        ("2.2.1", "Controller Serial (ASLLC)",     "Controller Serial (ASLLC)"),
        ("2.2.2", "Gateway Serial",                "Gateway Serial"),
        ("2.2.3", "Gateway Model",                 "Gateway Model"),
        ("2.2.4", "Gateway Software Version",      "Gateway Software Version"),
        ("2.2.5", "Firmware Version (ASLLC)",      "Firmware Version ASLLC"),
        ("2.2.6", "Load 13B87 Firmware on ASLLC",  "Load 13B87 Firmware on ASLLC (Airsys Smart Key)"),
    ], data)
    els += photo_subsection("ASLLC / Gateway Photos", [
        "Gateway Picture",
        "ASLLC Picture",
        "Controller Serial (ASLLC) picture",
        "Gateway Serial picture",
        "Gateway Model picture",
        "Gateway Software Version picture",
        "Firmware Version ASLLC Picture",
    ], data, photo_root, site_folder=site_folder, used=used)

    # Equipment-tag groups — each rendered only if it has any value.
    els += field_subsection("Equipment Tag — TCU-NG2", [
        ("2.3.1", "TCU-NG2 Tag already Present", "TCU-NG2 Tag already Present"),
    ], data)
    els += photo_subsection("TCU-NG2 Tag Photos", [
        "TCU-NG2 Tag Confirmation picture",
        "TCU-NG2 Tag before",
        "TCU-NG2 Tag after",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Equipment Tag — TCU800-MINI", [
        ("2.4.1", "TCU800-MINI Tag Present", "TCU800-MINI Tag Present"),
    ], data)
    els += photo_subsection("TCU800-MINI Tag Photos", [
        "TCU800-MINI Tag Confirmation picture",
        "TCU800-MINI Tag before",
        "TCU800-MINI Tag after",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Equipment Tag — ASLLC Controller", [
        ("2.5.1", "ASLLC Controller Tag Present", "ASLLC Controller Tag Present"),
    ], data)
    els += photo_subsection("ASLLC Controller Tag Photos", [
        "ASLLC Controller Tag Confirmation picture",
        "ASLLC Controller Tag before",
        "ASLLC Controller Tag after",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Tagging — Gateway", [
        ("2.6.1", "Added Tags for Gateway (if applicable)", "Added Tags for Gateway (if applicable)"),
    ], data)
    els += photo_subsection("Gateway Tag Photos", [
        "Gateway Tag Confirmation picture",
        "Gateway Tag before",
        "Gateway Tag after",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Tagging — Controller", [
        ("2.7.1", "Added Tags for Controller (if applicable)", "Added Tags for Controller (if applicable)"),
    ], data)
    els += photo_subsection("Controller Tag Photos", [
        "Added Tags for Controller (if applicable) picture",
        "Added Tags for Controller before",
        "Added Tags for Controller after",
    ], data, photo_root, site_folder=site_folder, used=used)
    return els


# ── Controller & General Site Checks ─────────────────────────────────────────

def build_controller_checks(data, photo_root, site_folder, used=None):
    els = section_header("CONTROLLER & GENERAL SITE CHECKS")

    els += field_subsection("Alarm and System Status", [
        ("3.1.1", "Active Alarms Present",             "Active Alarms Present"),
        ("3.1.2", "Active Alarms Details",             "Active Alarms Details"),
        ("3.1.3", "Alarm History Checked & Cleared",   "Alarm History Checked & Cleared"),
        ("3.1.4", "Alarm History Details",             "Alarm History Details"),
        ("3.1.5", "Input DC Voltage (Controller)",     "Input DC Voltage (Controller)", "V"),
        ("3.1.6", "Input DC Voltage (Gateway)",        "Input DC Voltage (Gateway)",    "V"),
    ], data)
    els += photo_subsection("Alarm Photos", [
        "Active Alarms Present picture",
        "Alarm History Checked & Cleared picture",
        "Input DC Voltage (Controller) picture",
        "Input DC Voltage (Gateway) picture",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Operating Mode", [
        ("3.2.1", "Controller System Mode",            "Controller System current Mode"),
        ("3.2.2", "System Mode Details (if not Auto)", "System Mode Details (if not Auto)"),
        ("3.2.3", "Controller Operating Mode",         "Controller Operating Mode"),
        ("3.2.4", "Fault Details (if Fault)",          "Fault Details (if Fault)"),
    ], data)
    els += photo_subsection("Operating Mode Photos", [
        "Controller System current Mode picture",
        "Controller Operating Mode picture",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Sensor — Indoor Humidity Sensor Condition", [
        ("3.3.1", "Indoor Air Temp Sensor Condition",   "Indoor Air Temp Sensor Condition"),
        ("3.3.2", "Indoor Humidity Sensor Condition",   "Indoor Humidity Sensor Condition"),
        ("3.3.3", "Indoor Temperature (on controller)", "Indoor Temperature (on controller)", "°C"),
    ], data)
    els += photo_subsection("Sensor Photos", [
        "Temp Sensor Photo",
        "Humidity Sensor Photo",
        "Indoor Temperature (on controller) picture",
    ], data, photo_root, site_folder=site_folder, used=used)

    return els


# ── Setpoints / Thresholds on Arrival ────────────────────────────────────────

ASLLC_SETPOINTS = [
    ("4.1.1",  "Lead Compressor Start (ASLLC)",  "Lead Compressor Start (ASLLC)",  "Lead Compressor Start (ASLLC) Photo"),
    ("4.1.2",  "Compressor Stop (ASLLC)",        "Compressor Stop (ASLLC)",        "Compressor Stop (ASLLC) Photo"),
    ("4.1.3",  "Lag Compressor Start (ASLLC)",   "Lag Compressor Start (ASLLC)",   "Lag Compressor Start (ASLLC) Photo"),
    ("4.1.4",  "Lag Compressor Stop (ASLLC)",    "Lag Compressor Stop (ASLLC)",    "Lag Compressor Stop (ASLLC) Photo"),
    ("4.1.5",  "Lead Heater Start (ASLLC)",      "Lead Heater Start (ASLLC)",      "Lead Heater Start (ASLLC) Photo"),
    ("4.1.6",  "Heaters Stop (ASLLC)",           "Heaters Stop (ASLLC)",           "Heaters Stop (ASLLC) Photo"),
    ("4.1.7",  "Lag Heater Start (ASLLC)",       "Lag Heater Start (ASLLC)",       "Lag Heater Start (ASLLC) Photo"),
    ("4.1.8",  "Lag Heater Stop (ASLLC)",        "Lag Heater Stop (ASLLC)",        "Lag Heater Stop (ASLLC) Photo"),
    ("4.1.9",  "Supply Fans Start (ASLLC)",      "Supply Fans Start (ASLLC)",      "Supply Fans Start (ASLLC) Photo"),
    ("4.1.10", "Supply Fan Stop (ASLLC)",        "Supply Fan Stop (ASLLC)",        "Supply Fan Stop (ASLLC) Photo"),
    ("4.1.11", "High Temp Alarm (ASLLC)",        "High Temp Alarm (ASLLC)",        "High Temp Alarm (ASLLC) Photo"),
    ("4.1.12", "Low Temp Alarm (ASLLC)",         "Low Temp Alarm (ASLLC)",         "Low Temp Alarm (ASLLC) Photo"),
]

NG2_SETPOINTS = [
    ("4.2.1",  "Lead Compressor Start (TCU-NG2)",   "Lead Compressor Start (TCU-NG2)",   "Lead Compressor Start (TCU-NG2) Photo"),
    ("4.2.2",  "Lead Compressors Stop (TCU-NG2)",   "Lead Compressors Stop (TCU-NG2)",   "Lead Compressors Stop (TCU-NG2) Photo"),
    ("4.2.3",  "Lag Compressor Start (TCU-NG2)",    "Lag Compressor Start (TCU-NG2)",    "Lag Compressor Start (TCU-NG2) Photo"),
    ("4.2.4",  "Lag Compressor Stop (TCU-NG2)",     "Lag Compressor Stop (TCU-NG2)",     "Lag Compressor Stop (TCU-NG2) Photo"),
    ("4.2.5",  "Lead Heater Start (TCU-NG2)",       "Lead Heater Start (TCU-NG2)",       "Lead Heater Start (TCU-NG2) Photo"),
    ("4.2.6",  "Heaters Stop (TCU-NG2)",            "Heaters Stop (TCU-NG2)",            "Heaters Stop (TCU-NG2) Photo"),
    ("4.2.7",  "Lag Heater Start (TCU-NG2)",        "Lag Heater Start (TCU-NG2)",        "Lag Heater Start (TCU-NG2) Photo"),
    ("4.2.8",  "Lag Heater Stop (TCU-NG2)",         "Lag Heater Stop (TCU-NG2)",         "Lag Heater Stop (TCU-NG2) Photo"),
    ("4.2.9",  "Supply Fans Start (TCU-NG2)",       "Supply Fans Start (TCU-NG2)",       "Supply Fans Start (TCU-NG2) Photo"),
    ("4.2.10", "Supply Fan Stop (TCU-NG2)",         "Supply Fan Stop (TCU-NG2)",         "Supply Fan Stop (TCU-NG2) Photo"),
    ("4.2.11", "High Temp Alarm (TCU-NG2)",         "High Temp Alarm (TCU-NG2)",         "High Temp Alarm (TCU-NG2) Photo"),
    ("4.2.12", "Low Temp Alarm (TCU-NG2)",          "Low Temp Alarm (TCU-NG2)",          "Low Temp Alarm (TCU-NG2) Photo"),
]

LEAVING_SETPOINTS = [
    ("6.1.1",  "Lead Compressor Start",  "Lead Compressor Start",  "Lead Compressor Start Photo"),
    ("6.1.2",  "Compressors Stop",       "Compressors Stop",       "Compressors Stop Photo"),
    ("6.1.3",  "Lag Compressor Start",   "Lag Compressor Start",   "Lag Compressor Start Photo"),
    ("6.1.4",  "Lag Compressor Stop",    "Lag Compressor Stop",    "Lag Compressor Stop Photo"),
    ("6.1.5",  "Lead Heater Start",      "Lead Heater Start",      "Lead Heater Start Photo"),
    ("6.1.6",  "Heaters Stop",           "Heaters Stop",           "Heaters Stop Photo"),
    ("6.1.7",  "Lag Heater Start",       "Lag Heater Start",       "Lag Heater Start Photo"),
    ("6.1.8",  "Lag Heater Stop",        "Lag Heater Stop",        "Lag Heater Stop Photo"),
    ("6.1.9",  "Supply Fans Start",      "Supply Fans Start",      "Supply Fans Start Photo"),
    ("6.1.10", "Supply Fan Stop",        "Supply Fan Stop",        "Supply Fan Stop Photo"),
    ("6.1.11", "High Temp Alarm",        "High Temp Alarm",        "High Temp Alarm Photo"),
    ("6.1.12", "Low Temp Alarm",         "Low Temp Alarm",         "Low Temp Alarm Photo"),
]


def build_setpoints(title, spec, data, photo_root, site_folder, used=None,
                    suffix="°C", required=True):
    """Render a setpoints subsection (values + accompanying photo grid).

    ``spec`` is a list of ``(code, label, value_col, photo_col)``.
    ``suffix`` is appended to numeric values (default ``"°C"``).
    ``required`` controls whether empty fields show as N/A (default True).
    Returns ``[]`` if neither values nor photos are populated.
    """
    rendered = []
    photo_cols = []
    for code, label, vcol, pcol in spec:
        cell = _value_cell(data.get(vcol, ""), suffix=suffix, required=required)
        if cell is not None:
            rendered.append([code, label, cell])
        if (data.get(pcol, "") or "").strip():
            photo_cols.append(pcol)

    if not rendered and not photo_cols:
        return []

    if rendered:
        first_code = rendered[0][0]
        if first_code and "." in first_code:
            prefix = first_code.rsplit(".", 1)[0] + "."
            for n, item in enumerate(rendered, start=1):
                item[0] = f"{prefix}{n}"

    header = subsection_header(title)
    rows = []
    for i, (code, label, cell) in enumerate(rendered):
        bg = ROW_ALT if i % 2 else W
        row = Table(
            [[Paragraph(code, S["code"]), Paragraph(label, S["lbl"]), cell]],
            colWidths=[CODE_W, LBL_W, VAL_W],
        )
        row.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), bg),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN",         (2, 0), (2, 0),   "RIGHT"),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING",   (0, 0), (0, 0),   6),
            ("LEFTPADDING",   (1, 0), (1, 0),   2),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
            ("LINEBELOW",     (0, 0), (-1, -1), 0.3, ROW_LN),
        ]))
        rows.append(row)

    # Keep header with first data row to avoid orphan headers at page bottom.
    if rows:
        first_block = KeepTogether(header + [rows[0]])
        els = [Spacer(1, 10), first_block] + rows[1:]
    else:
        els = [Spacer(1, 10)] + header
    if photo_cols:
        photo_rows = _photo_grid_rows(photo_cols, data, photo_root,
                                      site_folder=site_folder, max_h_cm=5,
                                      used=used)
        if photo_rows:
            els += [sp(8)] + photo_rows
    els += [sp(10)]
    return els


def build_arrival_setpoints(data, photo_root, site_folder, used=None):
    asllc = build_setpoints("4.1  Setpoints / Thresholds — ASLLC (on Arrival)",
                            ASLLC_SETPOINTS, data, photo_root, site_folder, used=used)
    ng2   = build_setpoints("4.2  Setpoints / Thresholds — TCU-NG2 (on Arrival)",
                            NG2_SETPOINTS, data, photo_root, site_folder, used=used)
    if not asllc and not ng2:
        return []
    return section_header("SETPOINTS / THRESHOLDS ON ARRIVAL") + asllc + ng2


# ── Shared System Testing & Failover ─────────────────────────────────────────

def build_shared_testing(data, photo_root, site_folder, used=None):
    units = (data.get("Number of HVAC Units", "") or "").strip()
    els = section_header("SHARED SYSTEM TESTING & FAILOVER",
                         subtitle=f"HVAC Units on Site: {units}" if units else None)

    els += field_subsection("Operational & Failover Testing", [
        ("5.1.1", "Lag HVAC Provides Cooling/Heating if Lead Fails/Demo",
                  "Lag HVAC Provides Cooling/Heating if Lead Fails/Demo"),
        ("5.1.2", "DC Fan Failover Works?",                                   "DC Fan Failover Works?"),
        ("5.1.3", "Temp Controls Maintaining Heat > +11°C and Cooling < +27°C?",
                  "Temp Controls Maintaining Heat >+11°C and Cooling <+27°C?"),
    ], data)
    els += photo_subsection("Failover Test Photos", [
        "Lag HVAC Provides Cooling/Heating if Lead Fails/Demo Photo",
        "DC Fan Failover Works? Photo",
        "Temp Controls Maintaining Heat >+11°C and Cooling <+27°C? Photo",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Sensor Verifications", [
        ("5.2.1", "Primary Indoor Temp Sensor Clean & Working?",        "Primary Indoor Temp Sensor Clean & Working?"),
        ("5.2.2", "Humidity Sensor Clean & Working?",                   "Humidity Sensor Clean & Working?"),
        ("5.2.3", "External Temp Sensor Clean of Debris & Working?",    "External Temp Sensor Clean of Debris & Working?"),
        ("5.2.4", "Air Supply Sensors for WPU1 & WPU2 Clean & Working?","Air Supply Sensors for WPU1 & WPU2 Clean & Working?"),
    ], data)
    els += photo_subsection("Sensor Verification Photos", [
        "Primary Indoor Temp Sensor Clean & Working? Photo",
        "Humidity Sensor Clean & Working? Photo",
        "External Temp Sensor Clean of Debris & Working? Photo",
        "Air Supply Sensors for WPU1 & WPU2 Clean & Working? Photo",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("Controller Health & Alarm Simulation Verification", [
        ("5.3.1", "System Clock Has the Correct Time?",                          "System Clock Has the Correct Time?"),
        ("5.3.2", "Controller Correctly Logging Errors/History (up to 6 months / 500 entries)?",
                  "Controller Correctly Logging Errors/History (up to 6 months / 500 entries)?"),
        ("5.3.3", "Individual Units Fail to Report Alarms Correctly? (TCU800-MINI only)",
                  "Individual Units Fail to Report Alarms Correctly? (TCU800-MINI only)"),
        ("5.3.4", "High/Low Temp Alarms Reporting Correctly?",                   "High/Low Temp Alarms Reporting Correctly?"),
        ("5.3.5", "Dirty Filter Condition Reporting Correctly?",                 "Dirty Filter Condition Reporting Correctly?"),
        ("5.3.6", "Refrigerant Low Pressure Reporting Correctly?",               "Refrigerant Low Pressure Reporting Correctly?"),
    ], data)
    els += photo_subsection("Controller Health Photos", [
        "System Clock Has the Correct Time? Photo",
        "Controller Correctly Logging Errors/History (up to 6 months / 500 entries)? Photo",
        "Individual Units Fail to Report Alarms Correctly? (TCU800-MINI only) Photo",
        "High/Low Temp Alarms Reporting Correctly? Photo",
        "Dirty Filter Condition Reporting Correctly? Photo",
        "Refrigerant Low Pressure Reporting Correctly? Photo",
    ], data, photo_root, site_folder=site_folder, used=used)

    return els


# ── Final Inspection & WNOC Clearance ────────────────────────────────────────

def build_final_inspection(data, photo_root, site_folder, used=None):
    els = section_header("FINAL INSPECTION & WNOC CLEARANCE")

    els += build_setpoints("6.1  Setpoints on Leaving",
                           LEAVING_SETPOINTS, data, photo_root, site_folder,
                           used=used)

    notes = (data.get("Additional Setpoint Notes", "") or "").strip()
    if notes:
        els += [Spacer(1, 6)] + subsection_header("Additional Setpoint Notes")
        els += [Paragraph(notes, S["notes"]), sp(10)]

    els += field_subsection("6.2  Hardware & Site Compliance", [
        ("6.2.1", "Compressor Terminal Condition",  "Compressor terminal condition"),
        ("6.2.2", "HVAC Breaker Type in AC Panel",  "HVAC Breaker type in AC panel"),
        ("6.2.3", "HVAC Manuals on Site?",          "HVAC Manuals on Site?"),
        ("6.2.4", "TSSA Sticker Valid?",            "TSSA Sticker Valid?"),
        ("6.2.5", "TSSA Sticker Expiry Date",       "TSSA Sticker Expiry Date"),
    ], data)
    els += photo_subsection("Hardware & Compliance Photos", [
        "Compressor terminal condition photo",
        "HVAC Breaker type in AC panel photo",
        "HVAC Manuals on Site? Photo",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("6.3  Final Physical Checks", [
        ("6.3.1", "Thermostat Set to Auto and Reading Accurately?",
                  "Thermostat Set to Auto and Reading Accurately?"),
        ("6.3.2", "All HVAC Vent Grills and Covers in Place?",
                  "All HVAC Vent Grills and Covers in Place?"),
        ("6.3.3", "Grills Notes",                                  "Grills"),
    ], data)
    els += photo_subsection("Final Physical Photos", [
        "Thermostat Set to Auto and Reading Accurately? Photo",
        "Grills Photo",
    ], data, photo_root, site_folder=site_folder, used=used)

    els += field_subsection("6.4  Final Controller Check & WNOC HVAC Status", [
        ("6.4.1", "HVAC Controller Indicates No Outstanding Alarms?",   "HVAC Controller Indicates No Outstanding Alarms?"),
        ("6.4.2", "Confirmed with WNOC that System is Free of Alarms?", "Confirmed with WNOC that System is Free of Alarms?"),
        ("6.4.3", "Finish Time",                                        "Finish Time"),
    ], data)
    els += photo_subsection("Final Controller Check Photos", [
        "HVAC Controller Indicates No Outstanding Alarms? Photo",
    ], data, photo_root, site_folder=site_folder, used=used)

    final_notes = (data.get("Final Notes", "") or "").strip()
    if final_notes:
        els += [Spacer(1, 6)] + subsection_header("Final Notes")
        els += [Paragraph(final_notes, S["notes"]), sp(10)]

    els += photo_subsection("Log Book Sign-Out", [
        "Log Book Sign-Out Photo",
    ], data, photo_root, site_folder=site_folder, used=used)

    return els


# ── Per-HVAC-unit data ──────────────────────────────────────────────────────

# Column layout of the AppSheet HVAC_Units export (header-less). Each row is
# one HVAC unit linked to a Survey by ``Survey ID`` (column 0). Order and
# names mirror the SiteSurvey_Field_Matrix codes 5.x → 7.x.
HVAC_UNIT_COLUMNS = [
    "Survey ID",                                                                  # 0
    "Make / Model",                                                                # 1   5.1.1
    "Make / Model Photo",                                                          # 2
    "Serial Number",                                                               # 3   5.1.2
    "Serial Number Photo",                                                         # 4
    "Running Status",                                                              # 5   5.1.3
    "Dial Position (Should be Max)",                                               # 6   5.2.1
    "Dial Position (Should be Max) Photo",                                         # 7
    "Relay & Connection Condition",                                                # 8   5.2.2
    "Relay & Connection Condition Photo",                                          # 9
    "Relay Secured with Cable Tie",                                                # 10  5.2.3
    "Relay Secured with Cable Tie Photo",                                          # 11
    "Condition of Supply and Return Grills",                                       # 12  5.3.1
    "Condition of Supply and Return Grills Photo",                                 # 13
    "Supply Air Temp Sensor Condition",                                            # 14  5.3.2
    "Supply Air Temp Sensor Condition Photo",                                      # 15
    "HVAC Paint & Label Condition",                                                # 16  5.4.1
    "HVAC Paint & Label Condition Photo",                                          # 17
    "Outdoor Temp Sensor Box (Inside & Out)",                                      # 18  5.4.2
    "Outdoor Temp Sensor Box (Inside & Out) Photo",                                # 19
    "Main Air Filter (Note Date & Type, MUST be MERV 8+)",                         # 20  5.5.1
    "Main Air Filter (Note Date & Type, MUST be MERV 8+) Photo",                   # 21
    "Pre-Filter Condition / Type",                                                 # 22  5.5.2
    "Pre-Filter Condition / Type Photo",                                           # 23
    "Condenser Coil Condition",                                                    # 24  5.5.3
    "Condenser Coil Condition Photo",                                              # 25
    "Evaporator Coil Condition",                                                   # 26  5.5.4
    "Evaporator Coil Condition Photo",                                             # 27
    "System Wiring Free from Wear/Damage?",                                        # 28  5.6.1
    "System Wiring Free from Wear/Damage? Photo",                                  # 29
    "Breakers & Wiring Correctly Sized for Load?",                                 # 30  5.6.2
    "Breakers & Wiring Correctly Sized for Load? Photo",                           # 31
    "DAMPER Actuator Motor in Correct Position & Secured?",                        # 32  5.6.3
    "DAMPER Actuator Motor in Correct Position & Secured? Photo",                  # 33
    "Bug Screen on Exhaust?",                                                      # 34  5.7.1
    "Bug Screen on Exhaust? Photo",                                                # 35
    "Bug Screen on Side Panel?",                                                   # 36  5.7.2
    "Bug Screen on Side Panel? Photo",                                             # 37
    "Bug Screen on Evaporator Drain Hole?",                                        # 38  5.7.3
    "Bug Screen on Evaporator Drain Hole? Photo",                                  # 39
    "Gasket Added on Pre-Filter?",                                                 # 40  5.7.4
    "Gasket Added on Pre-Filter? Photo",                                           # 41
    "Condenser Fan Condition",                                                     # 42  5.8.1
    "Condenser Fan Condition Photo",                                               # 43
    "Supply Fan(s) Condition",                                                     # 44  5.8.2
    "Supply Fan(s) Condition Photo",                                               # 45
    "Compressor Compartment Cleanliness",                                          # 46  5.8.3
    "Compressor Compartment Cleanliness Photo",                                    # 47
    "Evaporator Compartment Cleanliness",                                          # 48  5.8.4
    "Evaporator Compartment Cleanliness Photo",                                    # 49
    "Condenser Compartment Cleanliness",                                           # 50  5.8.5
    "Condenser Compartment Cleanliness Photo",                                     # 51
    "Replaced Main Air Filter & Labeled (Date, Location, MERV)?",                  # 52  6.1.1
    "Replaced Main Air Filter & Labeled (Date, Location, MERV)? Photo",            # 53
    "Replaced Nylon Mesh with Aluminum Mesh (if applicable)?",                     # 54  6.1.2
    "Replaced Nylon Mesh with Aluminum Mesh (if applicable)? Photo",               # 55
    "Cleaned Aluminum Mesh Pre-filter?",                                           # 56  6.1.3
    "Cleaned Aluminum Mesh Pre-filter? Photo",                                     # 57
    "Added Gasket on Pre-filter Slot?",                                            # 58  6.1.4
    "Added Gasket on Pre-filter Slot? Photo",                                      # 59
    "Added Bug Screen on Exhaust Grill?",                                          # 60  6.2.1
    "Added Bug Screen on Exhaust Grill? Photo",                                    # 61
    "Added Bug Screen on Evap Coil Drain?",                                        # 62  6.2.2
    "Added Bug Screen on Evap Coil Drain? Photo",                                  # 63
    "Added Bug Screen on Left Panel?",                                             # 64  6.2.3
    "Added Bug Screen on Left Panel? Photo",                                       # 65
    "Cleaned Evap Coil & Compartment?",                                            # 66  6.3.1
    "Cleaned Evap Coil & Compartment? Photo",                                      # 67
    "Cleaned Compressor Compartment?",                                             # 68  6.3.2
    "Cleaned Compressor Compartment? Photo",                                       # 69
    "Cleaned Condenser Coil (cleaning agent + water)?",                            # 70  6.3.3
    "Cleaned Condenser Coil (cleaning agent + water)? Photo",                      # 71
    "Cleaned Drainage Pipe?",                                                      # 72  6.3.4
    "Cleaned Drainage Pipe? Photo",                                                # 73
    "Lubricated Supply Fan?",                                                      # 74  6.3.5
    "Lubricated Supply Fan? Photo",                                                # 75
    "Lubricated Condenser Fan?",                                                   # 76  6.3.6
    "Lubricated Condenser Fan? Photo",                                             # 77
    "Straightened Coil Fins?",                                                     # 78  6.4.1
    "Straightened Coil Fins? Photo",                                               # 79
    "Replaced Bent ABS Fins on Grills?",                                           # 80  6.4.2
    "Replaced Bent ABS Fins on Grills? Photo",                                     # 81
    "Replaced Rusted Screws (steel) & Freed Screws (self-tapping)?",               # 82  6.4.3
    "Replaced Rusted Screws (steel) & Freed Screws (self-tapping)? Photo",         # 83
    "Grinded & Spray-Painted Rusted Areas?",                                       # 84  6.4.4
    "Grinded & Spray-Painted Rusted Areas? Photo",                                 # 85
    "Confirmed Damper Actuator Motor is Secure/Correct?",                          # 86  6.4.5
    "Confirmed Damper Actuator Motor is Secure/Correct? Photo",                    # 87
    "Performed Leak Test for R410A Refrigerant Throughout the HVAC?",              # 88  6.4.6
    "Performed Leak Test for R410A Refrigerant Throughout the HVAC? Photo",        # 89
    "Dial Position (Should be Max) [Post-Maintenance]",                            # 90  6.5.1
    "Dial Position (Should be Max) [Post-Maintenance] Photo",                      # 91
    "Secure Connection Condition",                                                 # 92  6.5.2
    "Secure Connection Condition Photo",                                           # 93
    "Relay Secured with Cable Tie [Post-Maintenance]",                             # 94  6.5.3
    "Relay Secured with Cable Tie [Post-Maintenance] Photo",                       # 95
    "Added Tags for HVAC Unit (if applicable)?",                                   # 96  6.6.1
    "Added Tags for HVAC Unit (if applicable)? Photo",                             # 97
    "Supply Fan — Amps",                                                            # 98  7.1.1
    "Supply Fan — Amps Photo",                                                      # 99
    "Supply Fan — Volts",                                                           # 100 7.1.2
    "Supply Fan — Volts Photo",                                                     # 101
    "Compressor + Condenser Fan — Amps",                                            # 102 7.1.3
    "Compressor + Condenser Fan — Amps Photo",                                      # 103
    "Compressor + Condenser Fan — Volts",                                           # 104 7.1.4
    "Compressor + Condenser Fan — Volts Photo",                                     # 105
    "Heater — Amps",                                                                # 106 7.1.5
    "Heater — Amps Photo",                                                          # 107
    "Heater — Volts",                                                               # 108 7.1.6
    "Heater — Volts Photo",                                                         # 109
    "Discharge Pressure (Without Compressor)",                                     # 110 7.2.1
    "Discharge Pressure (Without Compressor) Photo",                               # 111
    "Discharge Pressure (With Compressor)",                                        # 112 7.2.2
    "Discharge Pressure (With Compressor) Photo",                                  # 113
    "High Pressure Switch Operating Above 400 PSI",                                # 114 7.2.3
    "High Pressure Switch Operating Above 400 PSI Photo",                          # 115
    "Suction Pressure (Without Compressor)",                                       # 116 7.2.4
    "Suction Pressure (Without Compressor) Photo",                                 # 117
    "Suction Pressure (With Compressor)",                                          # 118 7.2.5
]


# Subsection layout for each per-unit section. Tuples in the field list are
# either ``(code, label, csv_col)`` or ``(code, label, csv_col, suffix)``;
# the photo list is just CSV column names. Subsection titles use the matrix
# numbering as a parent reference (5.1, 5.2, …).

HVAC_UNIT_BEFORE_LAYOUT = [
    ("5.1  Unit Identification", [
        ("5.1.1", "Make / Model",    "Make / Model"),
        ("5.1.2", "Serial Number",   "Serial Number"),
        ("5.1.3", "Running Status",  "Running Status"),
        ("5.1.4", "Operating Mode",  "Operating Mode"),
    ], ["Make / Model Photo", "Serial Number Photo"]),
    ("5.2  Emergency Thermostats", [
        ("5.2.1", "Dial Position (Should be Max)",   "Dial Position (Should be Max)"),
        ("5.2.2", "Relay & Connection Condition",    "Relay & Connection Condition"),
        ("5.2.3", "Relay Secured with Cable Tie",    "Relay Secured with Cable Tie"),
    ], [
        "Dial Position (Should be Max) Photo",
        "Relay & Connection Condition Photo",
        "Relay Secured with Cable Tie Photo",
    ]),
    ("5.3  Indoor Components", [
        ("5.3.1", "Condition of Supply and Return Grills", "Condition of Supply and Return Grills"),
        ("5.3.2", "Supply Air Temp Sensor Condition",      "Supply Air Temp Sensor Condition"),
    ], [
        "Condition of Supply and Return Grills Photo",
        "Supply Air Temp Sensor Condition Photo",
    ]),
    ("5.4  Outdoor Components", [
        ("5.4.1", "HVAC Paint & Label Condition",          "HVAC Paint & Label Condition"),
        ("5.4.2", "Outdoor Temp Sensor Box (Inside & Out)","Outdoor Temp Sensor Box (Inside & Out)"),
    ], [
        "HVAC Paint & Label Condition Photo",
        "Outdoor Temp Sensor Box (Inside & Out) Photo",
    ]),
    ("5.5  Filters & Coils", [
        ("5.5.1", "Main Air Filter (Note Date & Type, MUST be MERV 8+)",
                  "Main Air Filter (Note Date & Type, MUST be MERV 8+)"),
        ("5.5.2", "Pre-Filter Condition / Type",           "Pre-Filter Condition / Type"),
        ("5.5.3", "Condenser Coil Condition",              "Condenser Coil Condition"),
        ("5.5.4", "Evaporator Coil Condition",             "Evaporator Coil Condition"),
    ], [
        "Main Air Filter (Note Date & Type, MUST be MERV 8+) Photo",
        "Pre-Filter Condition / Type Photo",
        "Condenser Coil Condition Photo",
        "Evaporator Coil Condition Photo",
    ]),
    ("5.6  Hardware & Structural Checks", [
        ("5.6.1", "System Wiring Free from Wear/Damage?",  "System Wiring Free from Wear/Damage?"),
        ("5.6.2", "Breakers & Wiring Correctly Sized for Load?",
                  "Breakers & Wiring Correctly Sized for Load?"),
        ("5.6.3", "DAMPER Actuator Motor in Correct Position & Secured?",
                  "DAMPER Actuator Motor in Correct Position & Secured?"),
    ], [
        "System Wiring Free from Wear/Damage? Photo",
        "Breakers & Wiring Correctly Sized for Load? Photo",
        "DAMPER Actuator Motor in Correct Position & Secured? Photo",
    ]),
    ("5.7  Presence of Protective Elements", [
        ("5.7.1", "Bug Screen on Exhaust?",                "Bug Screen on Exhaust?"),
        ("5.7.2", "Bug Screen on Side Panel?",             "Bug Screen on Side Panel?"),
        ("5.7.3", "Bug Screen on Evaporator Drain Hole?",  "Bug Screen on Evaporator Drain Hole?"),
        ("5.7.4", "Gasket Added on Pre-Filter?",           "Gasket Added on Pre-Filter?"),
    ], [
        "Bug Screen on Exhaust? Photo",
        "Bug Screen on Side Panel? Photo",
        "Bug Screen on Evaporator Drain Hole? Photo",
        "Gasket Added on Pre-Filter? Photo",
    ]),
    ("5.8  Fan Noise & Cleanliness", [
        ("5.8.1", "Condenser Fan Condition",               "Condenser Fan Condition"),
        ("5.8.2", "Supply Fan(s) Condition",               "Supply Fan(s) Condition"),
        ("5.8.3", "Compressor Compartment Cleanliness",    "Compressor Compartment Cleanliness"),
        ("5.8.4", "Evaporator Compartment Cleanliness",    "Evaporator Compartment Cleanliness"),
        ("5.8.5", "Condenser Compartment Cleanliness",     "Condenser Compartment Cleanliness"),
    ], [
        "Condenser Fan Condition Photo",
        "Supply Fan(s) Condition Photo",
        "Compressor Compartment Cleanliness Photo",
        "Evaporator Compartment Cleanliness Photo",
        "Condenser Compartment Cleanliness Photo",
    ]),
]

HVAC_UNIT_MAINTENANCE_LAYOUT = [
    ("6.1  Filters", [
        ("6.1.1", "Replaced Main Air Filter & Labeled (Date, Location, MERV)?",
                  "Replaced Main Air Filter & Labeled (Date, Location, MERV)?"),
        ("6.1.2", "Replaced Nylon Mesh with Aluminum Mesh (if applicable)?",
                  "Replaced Nylon Mesh with Aluminum Mesh (if applicable)?"),
        ("6.1.3", "Cleaned Aluminum Mesh Pre-filter?",  "Cleaned Aluminum Mesh Pre-filter?"),
        ("6.1.4", "Added Gasket on Pre-filter Slot?",   "Added Gasket on Pre-filter Slot?"),
    ], [
        "Replaced Main Air Filter & Labeled (Date, Location, MERV)? Photo",
        "Replaced Nylon Mesh with Aluminum Mesh (if applicable)? Photo",
        "Cleaned Aluminum Mesh Pre-filter? Photo",
        "Added Gasket on Pre-filter Slot? Photo",
    ]),
    ("6.2  Hardware Additions (Bug Screens)", [
        ("6.2.1", "Added Bug Screen on Exhaust Grill?",   "Added Bug Screen on Exhaust Grill?"),
        ("6.2.2", "Added Bug Screen on Evap Coil Drain?", "Added Bug Screen on Evap Coil Drain?"),
        ("6.2.3", "Added Bug Screen on Left Panel?",      "Added Bug Screen on Left Panel?"),
    ], [
        "Added Bug Screen on Exhaust Grill? Photo",
        "Added Bug Screen on Evap Coil Drain? Photo",
        "Added Bug Screen on Left Panel? Photo",
    ]),
    ("6.3  Cleaning & Lubrication", [
        ("6.3.1", "Cleaned Evap Coil & Compartment?",     "Cleaned Evap Coil & Compartment?"),
        ("6.3.2", "Cleaned Compressor Compartment?",      "Cleaned Compressor Compartment?"),
        ("6.3.3", "Cleaned Condenser Coil (cleaning agent + water)?",
                  "Cleaned Condenser Coil (cleaning agent + water)?"),
        ("6.3.4", "Cleaned Drainage Pipe?",               "Cleaned Drainage Pipe?"),
        ("6.3.5", "Lubricated Supply Fan?",               "Lubricated Supply Fan?"),
        ("6.3.6", "Lubricated Condenser Fan?",            "Lubricated Condenser Fan?"),
    ], [
        "Cleaned Evap Coil & Compartment? Photo",
        "Cleaned Compressor Compartment? Photo",
        "Cleaned Condenser Coil (cleaning agent + water)? Photo",
        "Cleaned Drainage Pipe? Photo",
        "Lubricated Supply Fan? Photo",
        "Lubricated Condenser Fan? Photo",
    ]),
    ("6.4  Repairs & Corrective Actions", [
        ("6.4.1", "Straightened Coil Fins?",              "Straightened Coil Fins?"),
        ("6.4.2", "Replaced Bent ABS Fins on Grills?",    "Replaced Bent ABS Fins on Grills?"),
        ("6.4.3", "Replaced Rusted Screws (steel) & Freed Screws (self-tapping)?",
                  "Replaced Rusted Screws (steel) & Freed Screws (self-tapping)?"),
        ("6.4.4", "Grinded & Spray-Painted Rusted Areas?",
                  "Grinded & Spray-Painted Rusted Areas?"),
        ("6.4.5", "Confirmed Damper Actuator Motor is Secure/Correct?",
                  "Confirmed Damper Actuator Motor is Secure/Correct?"),
        ("6.4.6", "Performed Leak Test for R410A Refrigerant Throughout the HVAC?",
                  "Performed Leak Test for R410A Refrigerant Throughout the HVAC?"),
    ], [
        "Straightened Coil Fins? Photo",
        "Replaced Bent ABS Fins on Grills? Photo",
        "Replaced Rusted Screws (steel) & Freed Screws (self-tapping)? Photo",
        "Grinded & Spray-Painted Rusted Areas? Photo",
        "Confirmed Damper Actuator Motor is Secure/Correct? Photo",
        "Performed Leak Test for R410A Refrigerant Throughout the HVAC? Photo",
    ]),
    ("6.5  Emergency Thermostat (Post-Maintenance)", [
        ("6.5.1", "Dial Position (Should be Max)",       "Dial Position (Should be Max) [Post-Maintenance]"),
        ("6.5.2", "Secure Connection Condition",         "Secure Connection Condition"),
        ("6.5.3", "Relay Secured with Cable Tie",        "Relay Secured with Cable Tie [Post-Maintenance]"),
    ], [
        "Dial Position (Should be Max) [Post-Maintenance] Photo",
        "Secure Connection Condition Photo",
        "Relay Secured with Cable Tie [Post-Maintenance] Photo",
    ]),
    ("6.6  Tagging", [
        ("6.6.1", "Added Tags for HVAC Unit (if applicable)?",
                  "Added Tags for HVAC Unit (if applicable)?"),
    ], [
        "Added Tags for HVAC Unit (if applicable)? Photo",
    ]),
]

HVAC_UNIT_TESTING_LAYOUT = [
    ("7.1  Electrical Loads", [
        ("7.1.1", "Supply Fan — Amps",                    "Supply Fan — Amps",  "A"),
        ("7.1.2", "Supply Fan — Volts",                   "Supply Fan — Volts", "V"),
        ("7.1.3", "Compressor + Condenser Fan — Amps",    "Compressor + Condenser Fan — Amps",  "A"),
        ("7.1.4", "Compressor + Condenser Fan — Volts",   "Compressor + Condenser Fan — Volts", "V"),
        ("7.1.5", "Heater — Amps",                        "Heater — Amps",  "A"),
        ("7.1.6", "Heater — Volts",                       "Heater — Volts", "V"),
    ], [
        "Supply Fan — Amps Photo",
        "Supply Fan — Volts Photo",
        "Compressor + Condenser Fan — Amps Photo",
        "Compressor + Condenser Fan — Volts Photo",
        "Heater — Amps Photo",
        "Heater — Volts Photo",
    ]),
    ("7.2  Refrigerant Pressures", [
        ("7.2.1", "Discharge Pressure (Without Compressor)", "Discharge Pressure (Without Compressor)", "PSI"),
        ("7.2.2", "Discharge Pressure (With Compressor)",    "Discharge Pressure (With Compressor)",    "PSI"),
        ("7.2.3", "High Pressure Switch Operating Above 400 PSI",
                  "High Pressure Switch Operating Above 400 PSI"),
        ("7.2.4", "Suction Pressure (Without Compressor)",   "Suction Pressure (Without Compressor)",   "PSI"),
        ("7.2.5", "Suction Pressure (With Compressor)",      "Suction Pressure (With Compressor)",      "PSI"),
    ], [
        "Discharge Pressure (Without Compressor) Photo",
        "Discharge Pressure (With Compressor) Photo",
        "High Pressure Switch Operating Above 400 PSI Photo",
        "Suction Pressure (Without Compressor) Photo",
        "Suction Pressure (With Compressor) Photo",
    ]),
]


def _field_and_photo_subsection(title, field_specs, photo_cols,
                                data, photo_root, site_folder, used=None,
                                required=True):
    """Subsection that renders the field rows and then a photo grid below.

    Either or both halves may be empty; if everything is empty, returns []
    so the subsection header doesn't strand alone.
    """
    rows = field_rows(field_specs, data, required=required)
    photo_rows = []
    if photo_cols:
        photo_rows = _photo_grid_rows(
            photo_cols, data, photo_root,
            site_folder=site_folder, max_h_cm=6, used=used,
        )
    if not rows and not photo_rows:
        return []
    header = subsection_header(title)
    if rows:
        first_block = KeepTogether(header + [rows[0]])
        out = [Spacer(1, 10), first_block] + rows[1:]
    else:
        first_block = KeepTogether(header + ([photo_rows[0]] if photo_rows else []))
        out = [Spacer(1, 10), first_block]
        photo_rows = photo_rows[1:]
    if photo_rows:
        out += [sp(6)] + photo_rows
    return out + [Spacer(1, 10)]


def build_hvac_unit_section(title, layout, unit_data, photo_root, site_folder,
                             unit_idx=None, total_units=None, used=None):
    """Render a per-unit section with field rows + photos per subsection.

    ``layout`` is a list of ``(subsection_title, field_specs, photo_cols)``.
    """
    inner = []
    for sub_title, field_specs, photo_cols in layout:
        inner += _field_and_photo_subsection(
            sub_title, field_specs, photo_cols,
            unit_data, photo_root, site_folder, used=used,
        )
    if not inner:
        return []
    unit_label = None
    if unit_idx and total_units:
        unit_label = f"Unit {unit_idx} of {total_units}"
    return section_header(title, unit_label=unit_label) + inner


def build_hvac_before_state(unit_data, photo_root, site_folder,
                             unit_idx=None, total_units=None, used=None):
    return build_hvac_unit_section(
        "HVAC UNIT — BEFORE STATE", HVAC_UNIT_BEFORE_LAYOUT,
        unit_data, photo_root, site_folder,
        unit_idx=unit_idx, total_units=total_units, used=used,
    )


def build_hvac_maintenance(unit_data, photo_root, site_folder,
                            unit_idx=None, total_units=None, used=None):
    return build_hvac_unit_section(
        "HVAC UNIT — MAINTENANCE EXECUTION", HVAC_UNIT_MAINTENANCE_LAYOUT,
        unit_data, photo_root, site_folder,
        unit_idx=unit_idx, total_units=total_units, used=used,
    )


def build_hvac_testing(unit_data, photo_root, site_folder,
                        unit_idx=None, total_units=None, used=None):
    return build_hvac_unit_section(
        "HVAC UNIT — OPERATIONAL TESTING", HVAC_UNIT_TESTING_LAYOUT,
        unit_data, photo_root, site_folder,
        unit_idx=unit_idx, total_units=total_units, used=used,
    )


# ── End of report ─────────────────────────────────────────────────────────────

def build_end_of_report(data, photo_count):
    tech  = (data.get("Technician", "") or "").strip()
    prep  = tech if tech else "—"

    ps = Table(
        [[Paragraph("PHOTO SUMMARY", S["ps_l"]),
          Paragraph(f"{photo_count} photos rendered", S["ps_r"])]],
        colWidths=[CW * 0.55, CW * 0.45],
    )
    ps.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), SUB_BG),
        ("LEFTPADDING",   (0, 0), (0, 0),   12),
        ("RIGHTPADDING",  (1, 0), (1, 0),   12),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LINEABOVE",     (0, 0), (-1, 0),  2, SUB_ACC),
    ]))

    hr = HRFlowable(width="100%", thickness=1, color=HR_C, spaceAfter=20, spaceBefore=8)

    els = [ps, hr,
           Paragraph("END OF REPORT", S["end_h"]), hr,
           Paragraph("This report was automatically generated by the SiteSurvey field operations platform.",
                     S["end_b"]),
           Paragraph("Contents reflect data entered by the field technician during the on-site inspection.",
                     S["end_b"]),
           sp(14),
           Paragraph(f"<b>Prepared by:</b> {prep}", S["end_by"]),
           sp(20)]

    if QR_AVAILABLE:
        qr_data = f"PLC HVAC PM | {data.get('Site Name','')} | {data.get('Survey ID','')}"
        qr = _qrc.make(qr_data)
        buf = BytesIO(); qr.save(buf, format="PNG"); buf.seek(0)
        qr_img = Image(buf, width=3 * cm, height=3 * cm)
        qr_tbl = Table([[qr_img]], colWidths=[CW])
        qr_tbl.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))
        els += [qr_tbl, Paragraph("Scan to view report summary", S["qr_c"])]

    return els


# ── Photo counter & site folder detection ───────────────────────────────────

PHOTO_PATH_HINT = ("/", "\\")


def _is_photo_value(v):
    if not v:
        return False
    if not any(c in v for c in PHOTO_PATH_HINT):
        return False
    return Path(v).suffix.lower() in PHOTO_EXTS


def count_photos(data, photo_root, site_folder=None):
    n = 0
    seen = set()
    for val in data.values():
        v = (val or "").strip()
        if not _is_photo_value(v) or v in seen:
            continue
        if find_photo(v, photo_root, site_folder):
            seen.add(v); n += 1
    return n


def detect_site_folder(data, photo_root):
    """Locate the per-site folder based on the row's photo paths.

    Returns the folder where the PDF should sit (the parent of an ``images``
    subfolder when present), or ``None`` if no photos resolve.
    """
    if not photo_root:
        return None
    root = Path(photo_root)
    for val in data.values():
        v = (val or "").strip()
        if not _is_photo_value(v):
            continue
        cleaned = v.lstrip("/").replace("//", "/")
        cand = root / cleaned
        if cand.is_file():
            site = cand.parent
            if site.name.lower() == "images":
                site = site.parent
            return site
        # Resolve via basename search
        base = Path(v).name
        for p in root.rglob(base):
            if p.is_file():
                site = p.parent
                if site.name.lower() == "images":
                    site = site.parent
                return site
    return None


# ── CSV loaders ──────────────────────────────────────────────────────────────

def load_rows(path):
    """Read the Surveys CSV, deduplicating any repeated header names."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        try:
            headers = next(reader)
        except StopIteration:
            print("ERROR: CSV is empty"); sys.exit(1)
        seen = {}
        clean = []
        for h in headers:
            h = h.strip()
            if h in seen:
                seen[h] += 1
                clean.append(f"{h} ({seen[h]})")
            else:
                seen[h] = 1
                clean.append(h)
        rows = []
        for raw in reader:
            if not any((c or "").strip() for c in raw):
                continue
            d = {h: (raw[i] if i < len(raw) else "").strip()
                 for i, h in enumerate(clean)}
            rows.append(d)
    if not rows:
        print("ERROR: CSV has no data rows"); sys.exit(1)
    return rows


def load_hvac_units(path):
    """Read the HVAC_Units CSV.

    Supports both the old header-less positional format and the new format
    that includes a header row.  When a header row is present, duplicate
    column names (e.g. the second ``Dial Position (Should be Max)`` which
    belongs to the post-maintenance thermostat block) are disambiguated by
    appending ``[Post-Maintenance]`` to their name so the report layout can
    reference them unambiguously.

    Returns a list of per-unit dicts linked to their Survey by ``Survey ID``.
    """
    if not path or not Path(path).exists():
        return []

    with open(path, newline="", encoding="utf-8-sig") as f:
        raw_reader = csv.reader(f)
        try:
            first_row = next(raw_reader)
        except StopIteration:
            return []
        remaining = list(raw_reader)

    # Detect whether the file has a header: the old format has the Survey ID
    # value (an 8-char hex string) in the first column of the first row.
    has_header = not (len(first_row[0].strip()) == 8 and
                      all(c in "0123456789abcdef" for c in first_row[0].strip().lower()))

    if has_header:
        raw_headers = first_row
        data_rows   = remaining
    else:
        # Fall back to the legacy positional mapping.
        raw_headers = list(HVAC_UNIT_COLUMNS)
        data_rows   = [first_row] + remaining

    # Disambiguate duplicate column names by tagging the second occurrence
    # of each duplicate with [Post-Maintenance], inserted before a trailing
    # " Photo" suffix so the name matches the layout and photo-title dict.
    def _postmaint_tag(name):
        if name.endswith(" Photo"):
            base = name[: -len(" Photo")]
            return f"{base} [Post-Maintenance] Photo"
        return f"{name} [Post-Maintenance]"

    seen = {}
    headers = []
    for h in raw_headers:
        h = h.strip()
        if h in seen:
            seen[h] += 1
            headers.append(_postmaint_tag(h))
        else:
            seen[h] = 1
            headers.append(h)

    rows = []
    for raw in data_rows:
        if not any((c or "").strip() for c in raw):
            continue
        d = {h: (raw[i] if i < len(raw) else "").strip()
             for i, h in enumerate(headers)}
        rows.append(d)
    return rows


def units_for_survey(units, survey_id):
    """Filter HVAC unit rows to those belonging to a given Survey ID,
    preserving CSV order so Unit 1 / Unit 2 / Unit 3 align with how the
    technician walked the site."""
    sid = (survey_id or "").strip()
    if not sid:
        return []
    return [u for u in units if (u.get("Survey ID", "") or "").strip() == sid]


# ── Per-row build ─────────────────────────────────────────────────────────────

def build_report(data, photo_root, site_folder, out_path, units=None):
    """Render one survey to PDF.

    ``units`` is the list of per-HVAC-unit rows (already filtered to this
    survey). When omitted, sections 5/6/7 are simply skipped.
    """
    # Normalize the displayed date so the Survey Overview shows ISO format.
    if data.get("Report Date"):
        data["Report Date"] = iso_date(data["Report Date"])

    site_name = data.get("Site Name", "") or "—"
    used = set()
    units = list(units or [])

    doc   = make_doc(out_path, site_name)
    story = []

    story += build_cover(data)

    timeline = build_timeline(data)
    if timeline:
        story += timeline

    story += build_survey_overview(data, photo_root, site_folder, used=used)
    story.append(CondPageBreak(1 * cm))

    story += build_system_id(data, photo_root, site_folder, used=used)
    story.append(CondPageBreak(1 * cm))

    story += build_controller_checks(data, photo_root, site_folder, used=used)
    story.append(CondPageBreak(1 * cm))

    arrival_sp = build_arrival_setpoints(data, photo_root, site_folder, used=used)
    if arrival_sp:
        story += arrival_sp
        story.append(CondPageBreak(1 * cm))

    # Sections 5, 6, 7 — repeat once per HVAC unit, in CSV order.
    total = len(units)
    for idx, unit in enumerate(units, start=1):
        before = build_hvac_before_state(
            unit, photo_root, site_folder,
            unit_idx=idx, total_units=total, used=used,
        )
        if before:
            story += before
            story.append(CondPageBreak(1 * cm))
        maintenance = build_hvac_maintenance(
            unit, photo_root, site_folder,
            unit_idx=idx, total_units=total, used=used,
        )
        if maintenance:
            story += maintenance
            story.append(CondPageBreak(1 * cm))
        testing = build_hvac_testing(
            unit, photo_root, site_folder,
            unit_idx=idx, total_units=total, used=used,
        )
        if testing:
            story += testing
            story.append(CondPageBreak(1 * cm))

    story += build_shared_testing(data, photo_root, site_folder, used=used)
    story.append(CondPageBreak(1 * cm))

    story += build_final_inspection(data, photo_root, site_folder, used=used)
    story.append(CondPageBreak(1 * cm))

    n_photos = len(used) or count_photos(data, photo_root, site_folder)
    story += build_end_of_report(data, n_photos)

    doc.build(story)
    return n_photos


# ── Main ──────────────────────────────────────────────────────────────────────

def output_filename(data):
    site_id = (data.get("Location Code", "") or data.get("Survey ID", "") or "report").strip()
    rdate   = iso_date(data.get("Report Date", "")) or datetime.now().strftime("%Y-%m-%d")
    safe = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in site_id)
    return f"PM_Report_{safe}_{rdate}.pdf"


def _default_units_path(surveys_path: Path) -> Path | None:
    """Try to locate the HVAC_Units CSV next to the surveys CSV.

    Checks sibling files derived from the surveys filename, then a set of
    known fixed names.  Returns ``None`` if nothing matches — sections
    5/6/7 just stay empty in that case.
    """
    parent = surveys_path.parent
    stem = surveys_path.stem
    candidates = []
    if " - Surveys" in stem:
        base = stem.replace(" - Surveys", "")
        candidates += [
            parent / f"{base} - HVAC_Units.csv",
            parent / f"{base} - HVAC Units.csv",
        ]
    candidates += [
        parent / "HVAC_Units.csv",
        parent / "Bell Site Survey AppSheet Headers - HVAC_Units.csv",
        parent / "AppSheet Report Generation Data - HVAC_Units.csv",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def main():
    ap = argparse.ArgumentParser(
        description="Generate HVAC PM PDF report(s) from AppSheet CSV exports.")
    ap.add_argument("--csv",       required=True,
                    help="Surveys CSV (one row per survey).")
    ap.add_argument("--units-csv", default=None,
                    help="HVAC_Units CSV (one row per HVAC unit). Auto-detected "
                         "as a sibling of --csv if omitted.")
    ap.add_argument("--photos",    default=None,
                    help="Photo root (default: CSV's parent dir). The images/ "
                         "folder inside each site folder is searched too.")
    ap.add_argument("--output-dir",default=None,
                    help="Force all PDFs into this directory. By default each "
                         "PDF is saved next to its site's images folder.")
    ap.add_argument("--row",       type=int,      help="1-based row index to render (default: all)")
    ap.add_argument("--survey-id", default=None,  help="Render only the row with this Survey ID")
    args = ap.parse_args()

    csv_path = Path(args.csv).resolve()
    if not csv_path.exists():
        print(f"ERROR: {csv_path} not found"); sys.exit(1)

    if args.units_csv:
        units_path = Path(args.units_csv).resolve()
        if not units_path.exists():
            print(f"ERROR: {units_path} not found"); sys.exit(1)
    else:
        units_path = _default_units_path(csv_path)

    photo_root = Path(args.photos).resolve() if args.photos else csv_path.parent
    forced_out = Path(args.output_dir).resolve() if args.output_dir else None
    if forced_out:
        forced_out.mkdir(parents=True, exist_ok=True)

    rows  = load_rows(csv_path)
    units = load_hvac_units(units_path) if units_path else []

    selected = rows
    if args.survey_id:
        selected = [r for r in rows if r.get("Survey ID", "").strip() == args.survey_id.strip()]
        if not selected:
            print(f"ERROR: no row with Survey ID '{args.survey_id}'"); sys.exit(1)
    elif args.row:
        if args.row < 1 or args.row > len(rows):
            print(f"ERROR: --row must be between 1 and {len(rows)}"); sys.exit(1)
        selected = [rows[args.row - 1]]

    print(f"Surveys CSV: {csv_path}")
    print(f"Units CSV:   {units_path or '(none — sections 5/6/7 will be skipped)'}")
    print(f"Photo root:  {photo_root}")
    if forced_out:
        print(f"Output dir:  {forced_out}  (forced)")
    else:
        print("Output dir:  per-site (next to each site's images folder)")
    print(f"Rendering:   {len(selected)} of {len(rows)} row(s) — {len(units)} HVAC unit row(s) total")
    print("-" * 64)

    for i, data in enumerate(selected, 1):
        site_folder = detect_site_folder(data, photo_root)
        out_dir = forced_out or site_folder or photo_root
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / output_filename(data)

        site = data.get("Site Name", "")
        loc  = data.get("Location Code", "")
        tech = data.get("Technician", "")
        survey_id = data.get("Survey ID", "")
        my_units = units_for_survey(units, survey_id)
        print(f"[{i}/{len(selected)}] {loc} - {site}")
        print(f"        technician: {tech}")
        print(f"        site dir:   {site_folder}")
        print(f"        units:      {len(my_units)} (declared: {data.get('Number of HVAC Units','?')})")
        print(f"        output:     {out_path}")
        try:
            n_photos = build_report(data, photo_root, site_folder, str(out_path),
                                     units=my_units)
        except Exception as e:
            print(f"        FAILED: {e}")
            continue
        size_kb = out_path.stat().st_size // 1024
        print(f"        done ({n_photos} photo(s), {size_kb} KB)")

    print("-" * 64)
    print("All done.")


if __name__ == "__main__":
    main()
