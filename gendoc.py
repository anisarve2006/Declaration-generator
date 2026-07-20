#!/usr/bin/env python3
"""
Declaration Form Generator
===========================
A single-file desktop app (Tkinter GUI) that generates the VIT
"Student Undertaking for Ethical Academic Practice" declaration form
as a Word (.docx) and/or PDF file, for any subject/vertical and any
assignment/experiment number.

Visual design (font, sizing, spacing, layout) matches the reference
form exactly:
    - Font: Quattrocento Sans, 10pt body/labels/titles (bold for values
      & titles), justified for the ethics-warning paragraph.
    - Page: A4, 1 inch margins all round.
    - Paragraph spacing: 8pt after, 1.15x line spacing.
    - Titles are left-aligned (not centered), bold, same 10pt size.
    - Two-column fields (Branch/Vertical, Semester/Division,
      Subject Name/Code) are tab-aligned into a second column, and
      Semester+Division / Subject Name+Code share one paragraph via a
      line break, exactly as in the reference form.
    - Declaration statements live in a single paragraph separated by
      line breaks, using check (checked) / box (unchecked) glyphs.
    - Student Name/Roll No and Signature/Date share one paragraph via
      a line break.

Requirements (install once):
    pip install python-docx reportlab

Run:
    python declaration_form_generator.py

Output location:
    Files are saved automatically into "forms/docs" (.docx) and
    "forms/pdfs" (.pdf), created next to this script if they don't
    already exist. File name pattern:
    <SubjectCode>-<AssignmentNo>-Declaration Form.docx / .pdf
    e.g.  forms/docs/PCCE10T-Experiment1-Declaration Form.docx

Optional: drop a "college_logo.png" file next to this script and it
will be embedded at the top of both the DOCX and PDF outputs. You can
also attach a signature image (PNG/JPG) from the app's "Signature
Image" field -- it gets embedded on the Signature line of both files.

Fonts: this script looks for the Quattrocento Sans .ttf family in a
"fonts" folder next to it (Regular / Bold / Italic / BoldItalic). If
present, the PDF embeds the real typeface so it matches the DOCX
pixel-for-pixel; if absent, the PDF quietly falls back to Helvetica
and the DOCX still requests "Quattrocento Sans" (Word substitutes a
similar font on machines that don't have it installed).
"""

import os
import sys
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --------------------------------------------------------------------------
# 0. OUTPUT FOLDERS  (created automatically next to this script)
# --------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOCX_DIR = os.path.join(SCRIPT_DIR, "forms", "docs")
PDF_DIR = os.path.join(SCRIPT_DIR, "forms", "pdfs")
FONT_DIR = os.path.join(SCRIPT_DIR, "fonts")
os.makedirs(DOCX_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

# --------------------------------------------------------------------------
# 0.5 DESIGN CONSTANTS  (lifted from the reference form)
# --------------------------------------------------------------------------
FONT_NAME = "Quattrocento Sans"
BODY_SIZE_PT = 10          # w:sz val="20" (half-points) -> 10pt
PARA_SPACE_AFTER_PT = 8    # w:spacing w:after="160" (twips) -> 8pt
LINE_SPACING = 1.1583      # w:line="278" auto -> 278/240
SECOND_COL_TAB_IN = 3.3    # tab stop position for the 2nd field column
CHECK_MARK = "\u2713"      # (checked)
BOX_EMPTY = "\u2610"       # (unchecked)

# --------------------------------------------------------------------------
# 1. VERTICAL / SUBJECT DATA
#    Edit this dictionary to add / change verticals, subjects or codes.
#    Format:  "Vertical Key": [("Subject Name", "Subject Code"), ...]
# --------------------------------------------------------------------------
VERTICALS = {
    "HSSM_VEC": [
        ("E-Waste and Environmental Management", "VEC02"),
    ],
    "MDC_MDM": [
        ("Intelligent Mobile Robotics", "MDMRB03"),
        ("Innovation Management and Scaling Startups", "MDMIE03"),
        ("Machine Learning Applications in Bioinfomatics", "MDMBI03"),
        ("Strategic Marketing and Business Planning", "MDMBI03"),
    ],
    "MDC_OEC": [
        ("Professional Competency Development", "OEC15"),
    ],
    "PC_PCC": [
        ("Artificial Intelligence", "PCCE10T"),
        ("Computer Networks", "PCCE11T"),
        ("Software Engineering", "PCCE12T"),
        ("Theory Of Computation", "PCCE09"),
    ],
    "PC_PEC": [
        ("Data Warehousing and Mining", "PECE02T"),
    ],
}

DECLARATION_OPTIONS = [
    "I have prepared the assignment / academic activity entirely by myself.",
    "I have completed the assignment / academic activity with academic "
    "guidance or discussion from seniors / friends / peers, while ensuring "
    "the work is my own.",
    "I have used AI-based tools only as an aid and have appropriately "
    "acknowledged or cited their use as per the ethical guidelines of the "
    "Institute.",
    "I have used AI-based tools and directly copy-pasted content for the "
    "assignment / academic activity.",
    "I have copied the assignment / academic activity from peers / friends.",
]

LOGO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "college_logo.png")

# --------------------------------------------------------------------------
# 1.5 SIGNATURE PROCESSOR (Fixed Ratio & Max Dimensions)
# --------------------------------------------------------------------------
def process_signature_image(img_path_or_bytes):
    """
    Auto-enhances contrast, smooths freehand strokes using sub-pixel anti-aliasing,
    trims whitespace/transparent padding, and scales signature proportionally
    within fixed maximum limits (max height: 0.32 in, max width: 0.8 in --
    matching the reference form's embedded signature size).
    Returns: (processed_bytes_io, target_width_inches, target_height_inches)
    """
    if not img_path_or_bytes:
        return None, 0, 0
    try:
        from PIL import Image, ImageEnhance, ImageFilter
        import io
        if isinstance(img_path_or_bytes, str):
            if not os.path.exists(img_path_or_bytes):
                return None, 0, 0
            img = Image.open(img_path_or_bytes).convert("RGBA")
        else:
            img_path_or_bytes.seek(0)
            img = Image.open(img_path_or_bytes).convert("RGBA")

        # 1. Upsample for sub-pixel stroke precision
        w_orig, h_orig = img.size
        if w_orig > 0 and h_orig > 0:
            scale = 3
            img_work = img.resize((w_orig * scale, h_orig * scale), Image.Resampling.LANCZOS)
        else:
            img_work = img

        # 2. Extract luminance & enhance contrast to remove background noise
        gray = img_work.convert("L")
        gray_enhanced = ImageEnhance.Contrast(gray).enhance(2.2)

        # 3. Create smooth opacity mask
        mask = Image.eval(gray_enhanced, lambda v: 255 - v if v < 225 else 0)
        mask_smooth = mask.filter(ImageFilter.GaussianBlur(radius=1.2))
        mask_final = Image.eval(mask_smooth, lambda v: min(255, int(v * 1.6)))

        # 4. Generate crisp dark ink layer
        ink_layer = Image.new("RGBA", img_work.size, (15, 23, 42, 255))
        ink_layer.putalpha(mask_final)

        # 5. Downscale with Lanczos anti-aliasing for silky smooth strokes
        smooth_img = ink_layer.resize((w_orig, h_orig), Image.Resampling.LANCZOS)

        # 6. Crop tightly to signature content
        bbox = smooth_img.getchannel('A').getbbox()
        if bbox:
            smooth_img = smooth_img.crop(bbox)

        out = io.BytesIO()
        smooth_img.save(out, format="PNG")
        out.seek(0)

        w, h = smooth_img.size
        aspect = w / float(h) if h > 0 else 2.4

        MAX_H = 0.32  # matches reference form's embedded signature height
        MAX_W = 0.80  # matches reference form's embedded signature width

        target_h = MAX_H
        target_w = target_h * aspect

        if target_w > MAX_W:
            target_w = MAX_W
            target_h = target_w / aspect

        return out, target_w, target_h
    except Exception:
        return None, 0.79, 0.32


# --------------------------------------------------------------------------
# 2. DOCX GENERATION
# --------------------------------------------------------------------------
def generate_docx(data, out_path):
    try:
        from docx import Document
        from docx.shared import Pt, Inches, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
        from docx.oxml.ns import qn
    except ImportError:
        raise RuntimeError(
            "The 'python-docx' package is not installed.\n\n"
            "Open a terminal and run:\n"
            "    pip install python-docx\n\n"
            "(If 'pip' isn't recognized, try 'python -m pip install python-docx' "
            "or 'pip3 install python-docx'.)"
        )

    doc = Document()

    # --- Page setup: A4, 1 inch margins (matches reference sectPr) -------
    section = doc.sections[0]
    section.page_width = Cm(21.0)   # A4 width  (11909 twips)
    section.page_height = Cm(29.7)  # A4 height (16834 twips)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    # --- Base style: Quattrocento Sans 10pt, 8pt-after / 1.15 line -------
    normal = doc.styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = Pt(BODY_SIZE_PT)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.append(rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rFonts.set(qn(attr), FONT_NAME)
    normal.paragraph_format.space_after = Pt(PARA_SPACE_AFTER_PT)
    normal.paragraph_format.line_spacing = LINE_SPACING

    def set_font(run, bold=False):
        run.font.name = FONT_NAME
        run.font.size = Pt(BODY_SIZE_PT)
        run.bold = bold
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = rPr.makeelement(qn('w:rFonts'), {})
            rPr.append(rFonts)
        for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            rFonts.set(qn(attr), FONT_NAME)

    def add_para(justify=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(PARA_SPACE_AFTER_PT)
        p.paragraph_format.line_spacing = LINE_SPACING
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return p

    def add_run(p, text, bold=False):
        r = p.add_run(text)
        set_font(r, bold=bold)
        return r

    def add_tab(p):
        r = p.add_run("\t")
        set_font(r)
        p.paragraph_format.tab_stops.add_tab_stop(
            Inches(SECOND_COL_TAB_IN), WD_TAB_ALIGNMENT.LEFT)

    def add_break(p):
        r = p.add_run()
        set_font(r)
        r.add_break()

    # Optional logo
    if os.path.exists(LOGO_FILE):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(LOGO_FILE, width=Inches(1.2))

    # --- Title block: left-aligned, bold, same 10pt body size -----------
    p = add_para()
    add_run(p, "Student Undertaking for Ethical Academic Practice", bold=True)

    p = add_para()
    add_run(p, "Vidyalankar Institute of Technology, Mumbai", bold=True)

    add_para()  # blank spacer line, as in the reference form

    # --- Assignment No -----------------------------------------------
    p = add_para()
    add_run(p, "Assignment No: ")
    add_run(p, data["assignment_no"], bold=True)

    # --- Branch / Vertical --------------------------------------------
    p = add_para()
    add_run(p, "Branch: ")
    add_run(p, data["branch"], bold=True)
    add_tab(p)
    add_run(p, "Vertical: ")
    add_run(p, data["vertical"], bold=True)

    # --- Semester / Division  +  Subject Name / Code (one paragraph) --
    p = add_para()
    add_run(p, "Semester: ")
    add_run(p, data["semester"], bold=True)
    add_tab(p)
    add_run(p, "Division: ")
    add_run(p, data["division"], bold=True)
    add_break(p)
    add_run(p, "Subject Name: ")
    add_run(p, data["subject_name"], bold=True)
    add_tab(p)
    add_run(p, "Subject Code: ")
    add_run(p, data["subject_code"], bold=True)

    # --- Declaration intro line -----------------------------------------
    p = add_para()
    add_run(
        p,
        "I hereby declare that for the assignment / academic activity "
        "submitted by me, I confirm the following statement(s) by ticking "
        f"({BOX_EMPTY}) the appropriate box(es):"
    )

    # --- Declaration options: one paragraph, line breaks between items --
    p = add_para()
    add_break(p)
    for idx, text in enumerate(DECLARATION_OPTIONS, start=1):
        mark = CHECK_MARK if idx in data["selected_options"] else BOX_EMPTY
        add_run(p, f"{mark} ({idx}) {text}")
        if idx != len(DECLARATION_OPTIONS):
            add_break(p)

    # --- Ethics warning: justified ---------------------------------------
    p = add_para(justify=True)
    add_run(
        p,
        "I understand that selecting options (4) or (5) indicates unethical "
        "academic practice and may attract academic penalties, including "
        "rejection of the assignment or disciplinary action, as per the "
        "rules of Vidyalankar Institute of Technology, Mumbai. I submit "
        "this declaration truthfully and accept full responsibility for "
        "the same."
    )

    # --- Student Name / Roll No  +  Signature / Date (one paragraph) ----
    p = add_para()
    add_run(p, "Student Name: ")
    add_run(p, data["student_name"], bold=True)
    add_tab(p)
    add_run(p, "Roll No.: ")
    add_run(p, data["roll_no"], bold=True)
    add_break(p)
    add_run(p, "Signature: ")
    signature_path = data.get("signature_path")
    if signature_path:
        processed_sig, tw, th = process_signature_image(signature_path)
        if processed_sig:
            sig_run = p.add_run()
            set_font(sig_run)
            sig_run.add_picture(processed_sig, width=Inches(tw), height=Inches(th))
        else:
            add_run(p, "\u2003" * 6)
    else:
        add_run(p, "\u2003" * 6)
    add_tab(p)
    add_run(p, "Date: ")
    add_run(p, data["date"], bold=True)

    doc.save(out_path)


# --------------------------------------------------------------------------
# 3. PDF GENERATION  (reportlab -- no MS Word/LibreOffice required)
# --------------------------------------------------------------------------
def _register_pdf_fonts():
    """Register the real Quattrocento Sans family if the .ttf files are
    present in FONT_DIR; otherwise fall back to Helvetica silently."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.fonts import addMapping

    files = {
        "QuattrocentoSans": "QuattrocentoSans-Regular.ttf",
        "QuattrocentoSans-Bold": "QuattrocentoSans-Bold.ttf",
        "QuattrocentoSans-Italic": "QuattrocentoSans-Italic.ttf",
        "QuattrocentoSans-BoldItalic": "QuattrocentoSans-BoldItalic.ttf",
    }
    all_present = all(
        os.path.exists(os.path.join(FONT_DIR, fname)) for fname in files.values()
    )
    if not all_present:
        return "Helvetica", "Helvetica-Bold"

    for name, fname in files.items():
        pdfmetrics.registerFont(TTFont(name, os.path.join(FONT_DIR, fname)))
    addMapping("QuattrocentoSans", 0, 0, "QuattrocentoSans")
    addMapping("QuattrocentoSans", 1, 0, "QuattrocentoSans-Bold")
    addMapping("QuattrocentoSans", 0, 1, "QuattrocentoSans-Italic")
    addMapping("QuattrocentoSans", 1, 1, "QuattrocentoSans-BoldItalic")
    return "QuattrocentoSans", "QuattrocentoSans-Bold"


def generate_pdf(data, out_path):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import inch
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle
        from reportlab.lib.utils import ImageReader
    except ImportError:
        raise RuntimeError(
            "The 'reportlab' package is not installed.\n\n"
            "Open a terminal and run:\n"
            "    pip install reportlab\n\n"
            "(If 'pip' isn't recognized, try 'python -m pip install reportlab' "
            "or 'pip3 install reportlab'.)"
        )

    base_font, bold_font = _register_pdf_fonts()

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleL", parent=styles["Normal"], alignment=TA_LEFT,
        fontName=bold_font, fontSize=BODY_SIZE_PT,
        leading=BODY_SIZE_PT * LINE_SPACING, spaceAfter=PARA_SPACE_AFTER_PT,
    )
    normal = ParagraphStyle(
        "NormalL", parent=styles["Normal"], fontName=base_font,
        fontSize=BODY_SIZE_PT, leading=BODY_SIZE_PT * LINE_SPACING,
        spaceAfter=PARA_SPACE_AFTER_PT,
    )
    justify = ParagraphStyle("JustifyL", parent=normal, alignment=TA_JUSTIFY)

    # Tab-like second column via a preserved-space run of non-breaking
    # spaces sized to roughly reach SECOND_COL_TAB_IN from the margin.
    def col2(label_run_len_chars):
        pad = max(4, 46 - label_run_len_chars)
        return "&nbsp;" * pad

    story = []

    if os.path.exists(LOGO_FILE):
        img = RLImage(LOGO_FILE, width=1.2 * inch, height=1.2 * inch)
        img.hAlign = "CENTER"
        story.append(img)
        story.append(Spacer(1, 6))

    story.append(Paragraph("Student Undertaking for Ethical Academic Practice", title_style))
    story.append(Paragraph("Vidyalankar Institute of Technology, Mumbai", title_style))
    story.append(Spacer(1, PARA_SPACE_AFTER_PT))

    story.append(Paragraph(f"Assignment No: <b>{data['assignment_no']}</b>", normal))

    branch_txt = f"Branch: <b>{data['branch']}</b>"
    story.append(Paragraph(
        branch_txt + col2(len(f"Branch: {data['branch']}"))
        + f"Vertical: <b>{data['vertical']}</b>", normal))

    sem_txt = f"Semester: <b>{data['semester']}</b>"
    story.append(Paragraph(
        sem_txt + col2(len(f"Semester: {data['semester']}"))
        + f"Division: <b>{data['division']}</b>"
        + f"<br/>Subject Name: <b>{data['subject_name']}</b>"
        + col2(len(f"Subject Name: {data['subject_name']}"))
        + f"Subject Code: <b>{data['subject_code']}</b>", normal))

    story.append(Paragraph(
        "I hereby declare that for the assignment / academic activity "
        "submitted by me, I confirm the following statement(s) by ticking "
        "(&#9744;) the appropriate box(es):", normal))

    option_lines = []
    for idx, text in enumerate(DECLARATION_OPTIONS, start=1):
        mark = "&#10003;" if idx in data["selected_options"] else "&#9744;"
        option_lines.append(f"{mark} ({idx}) {text}")
    story.append(Paragraph("<br/>".join(option_lines), normal))

    story.append(Paragraph(
        "I understand that selecting options (4) or (5) indicates unethical "
        "academic practice and may attract academic penalties, including "
        "rejection of the assignment or disciplinary action, as per the "
        "rules of Vidyalankar Institute of Technology, Mumbai. I submit "
        "this declaration truthfully and accept full responsibility for "
        "the same.", justify))

    story.append(Paragraph(
        f"Student Name: <b>{data['student_name']}</b>"
        + col2(len(f"Student Name: {data['student_name']}"))
        + f"Roll No.: <b>{data['roll_no']}</b>", normal))

    signature_path = data.get("signature_path")
    sig_label = "Signature:"
    date_label = f"Date: <b>{data['date']}</b>"
    if signature_path:
        processed_sig, tw, th = process_signature_image(signature_path)
        if processed_sig:
            img_reader = ImageReader(processed_sig)
            sig_img = RLImage(img_reader, width=tw * inch, height=th * inch)
            sig_row = Table(
                [[Paragraph(sig_label, normal), sig_img, Paragraph(date_label, normal)]],
                colWidths=[0.75 * inch, tw * inch + 0.15 * inch, 2.5 * inch],
            )
            sig_row.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))
            story.append(sig_row)
        else:
            story.append(Paragraph(f"{sig_label}{col2(9)}{date_label}", normal))
    else:
        story.append(Paragraph(f"{sig_label}{col2(9)}{date_label}", normal))

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        topMargin=1 * inch, bottomMargin=1 * inch,
        leftMargin=1 * inch, rightMargin=1 * inch,
    )
    doc.build(story)


# --------------------------------------------------------------------------
# 4. FILE NAME HELPER
# --------------------------------------------------------------------------
def build_filename(subject_code, assignment_no, ext):
    safe_assignment = assignment_no.replace(" ", "").replace("/", "-")
    return f"{subject_code}-{safe_assignment}-Declaration Form.{ext}"


# --------------------------------------------------------------------------
# 5. TKINTER GUI
# --------------------------------------------------------------------------
class DeclarationFormApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Declaration Form Generator")
        self.geometry("600x750")
        self.minsize(520, 500)
        self.resizable(True, True)

        # Main scrollable canvas container
        main_canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        container = ttk.Frame(main_canvas)

        container.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        window_id = main_canvas.create_window((0, 0), window=container, anchor="nw")

        def _on_canvas_configure(event):
            main_canvas.itemconfig(window_id, width=event.width)

        main_canvas.bind("<Configure>", _on_canvas_configure)
        main_canvas.configure(yscrollcommand=scrollbar.set)

        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        scrollbar.pack(side="right", fill="y")
        main_canvas.pack(side="left", fill="both", expand=True)

        pad = {"padx": 14, "pady": 4}
        row = 0

        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Declaration Form Generator", font=("Segoe UI", 14, "bold")) \
            .grid(row=row, column=0, columnspan=2, pady=(12, 2)); row += 1

        ttk.Label(
            container,
            text="DOCX \u2192 forms/docs   \u2022   PDF \u2192 forms/pdfs",
            foreground="#555555",
        ).grid(row=row, column=0, columnspan=2, pady=(0, 8)); row += 1

        # Helper list of all (subject_name, subject_code, vertical)
        self.all_subjects = []
        for vert, subs in VERTICALS.items():
            for name, code in subs:
                self.all_subjects.append((name, code, vert))

        # Subject dropdown
        ttk.Label(container, text="Subject:").grid(row=row, column=0, sticky="w", **pad)
        self.subject_var = tk.StringVar()
        self.subject_combo = ttk.Combobox(
            container, textvariable=self.subject_var,
            values=[s[0] for s in self.all_subjects], state="readonly")
        self.subject_combo.grid(row=row, column=1, sticky="ew", **pad)
        self.subject_combo.bind("<<ComboboxSelected>>", self.on_subject_change)
        row += 1

        # Subject code (auto, read-only)
        ttk.Label(container, text="Subject Code:").grid(row=row, column=0, sticky="w", **pad)
        self.code_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.code_var, state="readonly") \
            .grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # Vertical (auto, read-only)
        ttk.Label(container, text="Vertical:").grid(row=row, column=0, sticky="w", **pad)
        self.vertical_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.vertical_var, state="readonly") \
            .grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # Assignment / Experiment No
        ttk.Label(container, text="Assignment / Experiment No:").grid(row=row, column=0, sticky="w", **pad)
        self.assignment_var = tk.StringVar(value="Experiment 1")
        ttk.Entry(container, textvariable=self.assignment_var) \
            .grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # Simple text fields
        fields = [
            ("Student Name", "student_name", "Anirudh Ghanshyam Sarve"),
            ("Roll No.", "roll_no", "24102A0062"),
            ("Branch", "branch", "CMPN"),
            ("Semester", "semester", "V"),
            ("Division", "division", "A"),
        ]
        self.field_vars = {}
        for label, key, default in fields:
            ttk.Label(container, text=label + ":").grid(row=row, column=0, sticky="w", **pad)
            var = tk.StringVar(value=default)
            ttk.Entry(container, textvariable=var).grid(row=row, column=1, sticky="ew", **pad)
            self.field_vars[key] = var
            row += 1

        # Date
        ttk.Label(container, text="Date:").grid(row=row, column=0, sticky="w", **pad)
        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%d-%m-%Y"))
        ttk.Entry(container, textvariable=self.date_var).grid(row=row, column=1, sticky="ew", **pad)
        row += 1

        # Signature image (optional)
        ttk.Label(container, text="Signature Image:").grid(row=row, column=0, sticky="w", **pad)
        sig_frame = ttk.Frame(container)
        sig_frame.grid(row=row, column=1, sticky="ew", **pad)
        sig_frame.columnconfigure(0, weight=1)
        self.signature_path_var = tk.StringVar(value="")
        ttk.Entry(sig_frame, textvariable=self.signature_path_var, state="readonly") \
            .grid(row=0, column=0, sticky="ew")
        ttk.Button(sig_frame, text="Browse...", command=self.browse_signature) \
            .grid(row=0, column=1, padx=(4, 0))
        ttk.Button(sig_frame, text="Clear", command=lambda: self.signature_path_var.set("")) \
            .grid(row=0, column=2, padx=(4, 0))
        row += 1

        # Declaration checkboxes
        ttk.Label(container, text="Tick applicable declaration statement(s):") \
            .grid(row=row, column=0, columnspan=2, sticky="w", padx=14, pady=(10, 2)); row += 1

        self.option_vars = []
        for idx, text in enumerate(DECLARATION_OPTIONS, start=1):
            var = tk.BooleanVar(value=(idx == 1))
            wrapped = self._wrap_text(f"({idx}) {text}", width=68)
            cb = ttk.Checkbutton(container, text=wrapped, variable=var)
            cb.grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=2)
            self.option_vars.append(var)
            row += 1

        # Buttons
        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=16)
        ttk.Button(btn_frame, text="Generate DOCX", command=lambda: self.generate(["docx"])) \
            .grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="Generate PDF", command=lambda: self.generate(["pdf"])) \
            .grid(row=0, column=1, padx=6)
        ttk.Button(btn_frame, text="Generate Both", command=lambda: self.generate(["docx", "pdf"])) \
            .grid(row=0, column=2, padx=6)

        self.status_var = tk.StringVar(value="")
        ttk.Label(container, textvariable=self.status_var, foreground="green") \
            .grid(row=row + 1, column=0, columnspan=2, pady=(0, 14))

        # init subject list
        if self.all_subjects:
            self.subject_combo.current(0)
            self.on_subject_change()

    # -- helpers ---------------------------------------------------------
    @staticmethod
    def _wrap_text(text, width=68):
        import textwrap
        return "\n".join(textwrap.wrap(text, width=width))

    # -- callbacks -----------------------------------------------------
    def browse_signature(self):
        path = filedialog.askopenfilename(
            title="Choose signature image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")],
        )
        if path:
            self.signature_path_var.set(path)

    def on_subject_change(self, event=None):
        subject = self.subject_var.get()
        for name, code, vertical in self.all_subjects:
            if name == subject:
                self.code_var.set(code)
                self.vertical_var.set(vertical)
                break

    def collect_data(self):
        selected = [i + 1 for i, v in enumerate(self.option_vars) if v.get()]
        if not selected:
            messagebox.showwarning("Missing selection", "Please tick at least one declaration statement.")
            return None
        if not self.assignment_var.get().strip():
            messagebox.showwarning("Missing field", "Please enter the Assignment / Experiment No.")
            return None
        data = {
            "vertical": self.vertical_var.get(),
            "subject_name": self.subject_var.get(),
            "subject_code": self.code_var.get(),
            "assignment_no": self.assignment_var.get().strip(),
            "date": self.date_var.get().strip(),
            "selected_options": selected,
            "signature_path": self.signature_path_var.get().strip() or None,
        }
        for key, var in self.field_vars.items():
            data[key] = var.get().strip()
        return data

    def generate(self, formats):
        data = self.collect_data()
        if data is None:
            return

        generated = []
        try:
            if "docx" in formats:
                fname = build_filename(data["subject_code"], data["assignment_no"], "docx")
                path = os.path.join(DOCX_DIR, fname)
                generate_docx(data, path)
                generated.append(path)
            if "pdf" in formats:
                fname = build_filename(data["subject_code"], data["assignment_no"], "pdf")
                path = os.path.join(PDF_DIR, fname)
                generate_pdf(data, path)
                generated.append(path)
        except Exception as exc:
            messagebox.showerror("Error generating file", str(exc))
            return

        self.status_var.set("Saved: " + ", ".join(os.path.basename(p) for p in generated))
        messagebox.showinfo("Done", "File(s) generated:\n" + "\n".join(generated))


# --------------------------------------------------------------------------
# 6. CLI FALLBACK (in case Tkinter / a display is unavailable, e.g. servers)
# --------------------------------------------------------------------------
def run_cli():
    print("Declaration Form Generator (CLI mode)\n")
    all_subjects = []
    for vert, subs in VERTICALS.items():
        for name, code in subs:
            all_subjects.append((name, code, vert))

    for i, (name, code, vert) in enumerate(all_subjects, 1):
        print(f"  {i}. {name} ({code}) [{vert}]")
    s_idx = int(input("Choose subject number: ")) - 1
    subject_name, subject_code, vertical = all_subjects[s_idx]

    data = {
        "vertical": vertical,
        "subject_name": subject_name,
        "subject_code": subject_code,
        "assignment_no": input("Assignment / Experiment No: ") or "Experiment 1",
        "student_name": input("Student Name [Anirudh Ghanshyam Sarve]: ") or "Anirudh Ghanshyam Sarve",
        "roll_no": input("Roll No. [24102A0062]: ") or "24102A0062",
        "branch": input("Branch [CMPN]: ") or "CMPN",
        "semester": input("Semester [V]: ") or "V",
        "division": input("Division [A]: ") or "A",
        "date": input(f"Date [{datetime.date.today().strftime('%d-%m-%Y')}]: ")
        or datetime.date.today().strftime("%d-%m-%Y"),
        "signature_path": input("Path to signature image (leave blank to skip): ").strip() or None,
    }
    print("Tick applicable statement(s), comma separated (e.g. 1,3):")
    for idx, text in enumerate(DECLARATION_OPTIONS, start=1):
        print(f"  ({idx}) {text}")
    raw = input("> ")
    data["selected_options"] = [int(x) for x in raw.split(",") if x.strip().isdigit()] or [1]

    fmt = input("Generate [docx/pdf/both]: ").strip().lower() or "both"
    formats = ["docx", "pdf"] if fmt == "both" else [fmt]

    for f in formats:
        fname = build_filename(data["subject_code"], data["assignment_no"], f)
        if f == "docx":
            path = os.path.join(DOCX_DIR, fname)
            generate_docx(data, path)
        else:
            path = os.path.join(PDF_DIR, fname)
            generate_pdf(data, path)
        print("Saved:", path)


if __name__ == "__main__":
    if "--cli" in sys.argv:
        run_cli()
    else:
        try:
            app = DeclarationFormApp()
            app.mainloop()
        except tk.TclError:
            print("No display available for the GUI -- falling back to CLI mode.\n")
            run_cli()