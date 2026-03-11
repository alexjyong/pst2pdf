"""PDF generation using reportlab, with attachment embedding via pypdf."""

from __future__ import annotations

import io
import logging
import os
from pathlib import Path

import html2text
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (
    HRFlowable,
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from models import Attachment, Email

logger = logging.getLogger(__name__)

_HEADER_BG   = colors.HexColor("#E8EEF4")
_LABEL_COLOR = colors.HexColor("#444444")
_ATT_BG      = colors.HexColor("#F5F5F5")
_PAGE_W, _PAGE_H = LETTER
_MARGIN = 0.75 * inch

_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif"}
_TEXT_EXTS  = {".txt", ".csv", ".log", ".md", ".rst", ".xml", ".json"}
_PDF_EXTS   = {".pdf"}


def _fmt_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{n} B"
        n /= 1024


def _styles():
    base = getSampleStyleSheet()
    label = ParagraphStyle(
        "HeaderLabel",
        parent=base["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=_LABEL_COLOR,
        leading=12,
    )
    value = ParagraphStyle(
        "HeaderValue",
        parent=base["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=12,
        wordWrap="CJK",
    )
    body = ParagraphStyle(
        "Body",
        parent=base["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=13,
        wordWrap="CJK",
    )
    att_header = ParagraphStyle(
        "AttHeader",
        parent=base["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=_LABEL_COLOR,
        leading=12,
    )
    att_notice = ParagraphStyle(
        "AttNotice",
        parent=base["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        textColor=colors.HexColor("#666666"),
        leading=12,
    )
    mono = ParagraphStyle(
        "Mono",
        parent=base["Normal"],
        fontName="Courier",
        fontSize=8,
        leading=11,
        wordWrap="CJK",
    )
    return label, value, body, att_header, att_notice, mono


def _header_table(email: Email, label_style, value_style):
    date_str = email.date.strftime("%Y-%m-%d %H:%M:%S %Z") if email.date else "Unknown"
    to_str   = ", ".join(email.to) if email.to else ""
    cc_str   = ", ".join(email.cc) if email.cc else ""
    att_names = ", ".join(a.name for a in email.attachments) if email.attachments else "None"

    rows = [("From:", email.from_), ("To:", to_str)]
    if cc_str:
        rows.append(("CC:", cc_str))
    rows += [
        ("Date:",       date_str),
        ("Subject:",    email.subject),
        ("Message-ID:", email.message_id or "(none)"),
        ("PST Path:",   email.folder_path),
        ("Attachments:", att_names),
    ]

    table_data = [
        [Paragraph(lbl, label_style), Paragraph(val, value_style)]
        for lbl, val in rows
    ]
    col_w = _PAGE_W - 2 * _MARGIN
    table = Table(table_data, colWidths=[1.1 * inch, col_w - 1.1 * inch])
    table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), _HEADER_BG),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID",          (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
    ]))
    return table


def _att_separator(att: Attachment, index: int, total: int, att_header_style, att_notice_style):
    """Return flowables for an attachment section header."""
    flowables = []
    flowables.append(Spacer(1, 0.15 * inch))
    flowables.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#AAAAAA"), dash=(2, 2)))
    flowables.append(Spacer(1, 0.05 * inch))

    size_str = _fmt_size(att.size) if att.size else "unknown size"
    label = f"Attachment {index + 1} of {total}: {att.name}  ({size_str})"
    flowables.append(Paragraph(label, att_header_style))
    flowables.append(Spacer(1, 0.05 * inch))
    return flowables


def _att_notice_flowables(message: str, att_notice_style):
    return [Paragraph(message, att_notice_style), Spacer(1, 0.1 * inch)]


def _get_body_text(email: Email) -> str:
    if email.body_text and email.body_text.strip():
        return email.body_text
    if email.body_html and email.body_html.strip():
        converter = html2text.HTML2Text()
        converter.ignore_links = False
        converter.body_width = 0
        return converter.handle(email.body_html)
    return "(No message body)"


def _make_bates_callback(bates_prefix: str, bates_start: int, pad_width: int, page_counter: list[int]):
    def draw_bates(canvas, doc):
        label = f"{bates_prefix}_{page_counter[0]:0{pad_width}d}"
        canvas.saveState()
        canvas.setFont("Helvetica-Bold", 8)
        label_w = canvas.stringWidth(label, "Helvetica-Bold", 8)
        x = _PAGE_W - _MARGIN - label_w - 4
        y = 0.35 * inch
        canvas.setFillColor(colors.white)
        canvas.rect(x - 3, y - 2, label_w + 10, 12, fill=1, stroke=0)
        canvas.setStrokeColor(colors.HexColor("#999999"))
        canvas.setLineWidth(0.5)
        canvas.rect(x - 3, y - 2, label_w + 10, 12, fill=0, stroke=1)
        canvas.setFillColor(colors.black)
        canvas.drawString(x, y, label)
        canvas.restoreState()
        page_counter[0] += 1
    return draw_bates


def _bates_stamp_page(pdf_page, label: str):
    """Overlay a Bates label onto a pypdf page object in-place."""
    pw = float(pdf_page.mediabox.width)
    ph = float(pdf_page.mediabox.height)
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(pw, ph))
    c.setFont("Helvetica-Bold", 8)
    label_w = c.stringWidth(label, "Helvetica-Bold", 8)
    x = pw - 54 - label_w
    y = 25
    c.setFillColorRGB(1, 1, 1)
    c.rect(x - 3, y - 2, label_w + 10, 12, fill=1, stroke=0)
    c.setStrokeColorRGB(0.6, 0.6, 0.6)
    c.setLineWidth(0.5)
    c.rect(x - 3, y - 2, label_w + 10, 12, fill=0, stroke=1)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(x, y, label)
    c.save()
    buf.seek(0)
    stamp_page = PdfReader(buf).pages[0]
    pdf_page.merge_page(stamp_page)


def _build_story(email: Email) -> tuple[list, list[Attachment]]:
    """Build the reportlab story and return (story, pdf_attachments_to_append)."""
    label_style, value_style, body_style, att_header_style, att_notice_style, mono_style = _styles()
    total_atts = len(email.attachments)

    story = []
    story.append(_header_table(email, label_style, value_style))
    story.append(Spacer(1, 0.15 * inch))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Spacer(1, 0.1 * inch))

    # Email body
    for para in _get_body_text(email).split("\n"):
        para = para.strip()
        if para:
            para = para.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(para, body_style))
        else:
            story.append(Spacer(1, 0.07 * inch))

    # Attachments
    pdf_atts: list[Attachment] = []
    avail_w = _PAGE_W - 2 * _MARGIN

    for i, att in enumerate(email.attachments):
        story.extend(_att_separator(att, i, total_atts, att_header_style, att_notice_style))
        ext = Path(att.name).suffix.lower()

        if att.is_embedded_msg:
            story.extend(_att_notice_flowables(
                f"Embedded message — open the source PST to view: {att.name}",
                att_notice_style,
            ))

        elif att.skipped:
            story.extend(_att_notice_flowables(
                f"Not embedded — file exceeds size limit ({_fmt_size(att.size)}). "
                f"Extract from PST separately: {att.name}",
                att_notice_style,
            ))

        elif not att.data:
            story.extend(_att_notice_flowables(
                f"Attachment data could not be read: {att.name}",
                att_notice_style,
            ))

        elif ext in _IMAGE_EXTS:
            try:
                img_buf = io.BytesIO(att.data)
                img_reader = ImageReader(img_buf)
                iw, ih = img_reader.getSize()
                # Scale to fit page width, cap height at 2/3 page
                scale = min(avail_w / iw, (_PAGE_H * 0.65) / ih, 1.0)
                story.append(Image(io.BytesIO(att.data), width=iw * scale, height=ih * scale))
            except Exception as exc:
                logger.warning("Could not embed image '%s': %s", att.name, exc)
                story.extend(_att_notice_flowables(f"Image could not be rendered: {exc}", att_notice_style))

        elif ext in _TEXT_EXTS:
            try:
                text = att.data.decode("utf-8", errors="replace")
                for line in text.splitlines():
                    line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    story.append(Paragraph(line or "&nbsp;", mono_style))
            except Exception as exc:
                logger.warning("Could not embed text '%s': %s", att.name, exc)
                story.extend(_att_notice_flowables(f"Text could not be rendered: {exc}", att_notice_style))

        elif ext in _PDF_EXTS:
            # PDF attachments are appended post-build via pypdf
            pdf_atts.append(att)
            story.extend(_att_notice_flowables(
                f"PDF attachment — pages follow below: {att.name}",
                att_notice_style,
            ))

        else:
            story.extend(_att_notice_flowables(
                f"Format not supported for embedding ({ext or 'no extension'}): {att.name}",
                att_notice_style,
            ))

    return story, pdf_atts


def _build_story_no_attachments(email: Email) -> list:
    """Build story with only email header + body — no attachment pages."""
    label_style, value_style, body_style, _, _, _ = _styles()
    story = []
    story.append(_header_table(email, label_style, value_style))
    story.append(Spacer(1, 0.15 * inch))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Spacer(1, 0.1 * inch))
    for para in _get_body_text(email).split("\n"):
        para = para.strip()
        if para:
            para = para.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(para, body_style))
        else:
            story.append(Spacer(1, 0.07 * inch))
    return story


def render_email_to_pdf(
    email: Email,
    output_path: str,
    bates_prefix: str = "",
    bates_start: int = 0,
    bates_pad_width: int = 6,
    embed_attachments: bool = True,
) -> int:
    """Render a single Email to a PDF file at output_path.

    Returns the next available Bates number.
    """
    story, pdf_atts = _build_story(email) if embed_attachments else (_build_story_no_attachments(email), [])
    bottom_margin = (1.0 if bates_prefix else 0.75) * inch

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=LETTER,
        leftMargin=_MARGIN,
        rightMargin=_MARGIN,
        topMargin=_MARGIN,
        bottomMargin=bottom_margin,
        title=email.subject or "(No Subject)",
        author=email.from_,
    )

    if bates_prefix:
        page_counter = [bates_start]
        cb = _make_bates_callback(bates_prefix, bates_start, bates_pad_width, page_counter)
        doc.build(story, onFirstPage=cb, onLaterPages=cb)
        next_bates = page_counter[0]
    else:
        doc.build(story)
        next_bates = bates_start

    # Merge PDF attachments using pypdf
    if pdf_atts:
        buf.seek(0)
        writer = PdfWriter()
        writer.append(buf)

        for att in pdf_atts:
            try:
                att_reader = PdfReader(io.BytesIO(att.data))
                for page in att_reader.pages:
                    if bates_prefix:
                        label = f"{bates_prefix}_{next_bates:0{bates_pad_width}d}"
                        _bates_stamp_page(page, label)
                        next_bates += 1
                    writer.add_page(page)
            except Exception as exc:
                logger.warning("Could not append PDF attachment '%s': %s", att.name, exc)

        with open(output_path, "wb") as f:
            writer.write(f)
    else:
        buf.seek(0)
        with open(output_path, "wb") as f:
            f.write(buf.read())

    return next_bates
