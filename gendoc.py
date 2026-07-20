#!/usr/bin/env python3
"""
Declaration Form Generator
===========================
A single-file desktop app (Tkinter GUI) that generates the VIT
"Student Undertaking for Ethical Academic Practice" declaration form
as a Word (.docx) and/or PDF file, for any subject/vertical and any
assignment/experiment number.

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
Image" field — it gets embedded on the Signature line of both files.
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
os.makedirs(DOCX_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

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
# 2. DOCX GENERATION
# --------------------------------------------------------------------------
def generate_docx(data, out_path):
    try:
        from docx import Document
        from docx.shared import Pt, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        raise RuntimeError(
            "The 'python-docx' package is not installed.\n\n"
            "Open a terminal and run:\n"
            "    pip install python-docx\n\n"
            "(If 'pip' isn't recognized, try 'python -m pip install python-docx' "
            "or 'pip3 install python-docx'.)"
        )

    doc = Document()

    # Optional logo
    if os.path.exists(LOGO_FILE):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(LOGO_FILE, width=Inches(1.2))

    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(14)

    def add_line(pairs):
        """pairs: list of (label, value, bold_value) rendered on one line."""
        p = doc.add_paragraph()
        for i, (label, value) in enumerate(pairs):
            if i > 0:
                p.add_run("\u2003\u2003\u2003\u2003")
            p.add_run(label)
            r = p.add_run(str(value))
            r.bold = True
        return p

    add_title("Student Undertaking for Ethical Academic Practice")
    add_title("Vidyalankar Institute of Technology, Mumbai")
    doc.add_paragraph("")

    add_line([("Assignment No: ", data["assignment_no"])])
    add_line([("Branch: ", data["branch"]), ("Vertical: ", data["vertical"])])

    p = doc.add_paragraph()
    p.add_run("Semester: ")
    p.add_run(data["semester"]).bold = True
    p.add_run("\u2003\u2003\u2003\u2003Division: ")
    p.add_run(data["division"]).bold = True
    p.add_run("\nSubject Name: ")
    p.add_run(data["subject_name"]).bold = True
    p.add_run("\u2003\u2003\u2003\u2003Subject Code: ")
    p.add_run(data["subject_code"]).bold = True

    doc.add_paragraph(
        "I hereby declare that for the assignment / academic activity "
        "submitted by me, I confirm the following statement(s) by ticking "
        "(\u2611) the appropriate box(es):"
    )

    for idx, text in enumerate(DECLARATION_OPTIONS, start=1):
        mark = "\u2611" if idx in data["selected_options"] else "\u2610"
        doc.add_paragraph(f"{mark} ({idx}) {text}")

    warn = doc.add_paragraph(
        "I understand that selecting options (4) or (5) indicates unethical "
        "academic practice and may attract academic penalties, including "
        "rejection of the assignment or disciplinary action, as per the "
        "rules of Vidyalankar Institute of Technology, Mumbai. I submit "
        "this declaration truthfully and accept full responsibility for "
        "the same."
    )
    warn.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    p = doc.add_paragraph()
    p.add_run("Student Name: ")
    p.add_run(data["student_name"]).bold = True
    p.add_run("\u2003\u2003Roll No.: ")
    p.add_run(data["roll_no"]).bold = True

    sig_p = doc.add_paragraph()
    sig_p.add_run("Signature: ")
    signature_path = data.get("signature_path")
    if signature_path and os.path.exists(signature_path):
        sig_p.add_run().add_picture(signature_path, height=Inches(0.45))
    else:
        sig_p.add_run("\u2003" * 10)  # blank space for a physical signature
    sig_p.add_run("\u2003\u2003Date: ")
    sig_p.add_run(data["date"]).bold = True

    doc.save(out_path)


# --------------------------------------------------------------------------
# 3. PDF GENERATION  (reportlab -- no MS Word/LibreOffice required)
# --------------------------------------------------------------------------
def generate_pdf(data, out_path):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import inch
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
    except ImportError:
        raise RuntimeError(
            "The 'reportlab' package is not installed.\n\n"
            "Open a terminal and run:\n"
            "    pip install reportlab\n\n"
            "(If 'pip' isn't recognized, try 'python -m pip install reportlab' "
            "or 'pip3 install reportlab'.)"
        )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleC", parent=styles["Normal"], alignment=TA_CENTER,
        fontName="Helvetica-Bold", fontSize=14, spaceAfter=4,
    )
    normal = ParagraphStyle("NormalL", parent=styles["Normal"], fontSize=11, spaceAfter=8, leading=15)
    justify = ParagraphStyle("JustifyL", parent=normal, alignment=TA_JUSTIFY)

    story = []

    if os.path.exists(LOGO_FILE):
        img = Image(LOGO_FILE, width=1.2 * inch, height=1.2 * inch)
        img.hAlign = "CENTER"
        story.append(img)
        story.append(Spacer(1, 6))

    story.append(Paragraph("Student Undertaking for Ethical Academic Practice", title_style))
    story.append(Paragraph("Vidyalankar Institute of Technology, Mumbai", title_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph(f"Assignment No: <b>{data['assignment_no']}</b>", normal))
    story.append(Paragraph(
        f"Branch: <b>{data['branch']}</b>"
        f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        f"Vertical: <b>{data['vertical']}</b>", normal))
    story.append(Paragraph(
        f"Semester: <b>{data['semester']}</b>"
        f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        f"Division: <b>{data['division']}</b>", normal))
    story.append(Paragraph(
        f"Subject Name: <b>{data['subject_name']}</b>"
        f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        f"Subject Code: <b>{data['subject_code']}</b>", normal))

    story.append(Paragraph(
        "I hereby declare that for the assignment / academic activity "
        "submitted by me, I confirm the following statement(s) by ticking "
        "(&#9745;) the appropriate box(es):", normal))

    for idx, text in enumerate(DECLARATION_OPTIONS, start=1):
        mark = "&#9745;" if idx in data["selected_options"] else "&#9744;"
        story.append(Paragraph(f"{mark} ({idx}) {text}", normal))

    story.append(Paragraph(
        "I understand that selecting options (4) or (5) indicates unethical "
        "academic practice and may attract academic penalties, including "
        "rejection of the assignment or disciplinary action, as per the "
        "rules of Vidyalankar Institute of Technology, Mumbai. I submit "
        "this declaration truthfully and accept full responsibility for "
        "the same.", justify))

    from reportlab.platypus import Table, TableStyle

    story.append(Spacer(1, 14))
    story.append(Paragraph(
        f"Student Name: <b>{data['student_name']}</b>"
        f"&nbsp;&nbsp;&nbsp;&nbsp;Roll No.: <b>{data['roll_no']}</b>", normal))

    signature_path = data.get("signature_path")
    if signature_path and os.path.exists(signature_path):
        sig_img = Image(signature_path, width=1.4 * inch, height=0.5 * inch)
        sig_row = Table(
            [[Paragraph("Signature:", normal), sig_img, Paragraph(f"Date: <b>{data['date']}</b>", normal)]],
            colWidths=[0.9 * inch, 1.6 * inch, 2.5 * inch],
        )
        sig_row.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(sig_row)
    else:
        story.append(Paragraph(
            f"Signature&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            f"Date: <b>{data['date']}</b>", normal))

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        topMargin=0.9 * inch, bottomMargin=0.9 * inch,
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
        self.geometry("560x680")
        self.resizable(False, False)

        pad = {"padx": 12, "pady": 5}
        row = 0

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Declaration Form Generator", font=("Segoe UI", 14, "bold")) \
            .grid(row=row, column=0, columnspan=2, pady=(14, 2)); row += 1

        ttk.Label(
            container,
            text=f"DOCX \u2192 forms/docs   \u2022   PDF \u2192 forms/pdfs",
            foreground="#555555",
        ).grid(row=row, column=0, columnspan=2, pady=(0, 10)); row += 1

        # Vertical dropdown
        ttk.Label(container, text="Vertical:").grid(row=row, column=0, sticky="w", **pad)
        self.vertical_var = tk.StringVar()
        self.vertical_combo = ttk.Combobox(
            container, textvariable=self.vertical_var,
            values=list(VERTICALS.keys()), state="readonly")
        self.vertical_combo.grid(row=row, column=1, sticky="ew", **pad)
        self.vertical_combo.bind("<<ComboboxSelected>>", self.on_vertical_change)
        row += 1

        # Subject dropdown
        ttk.Label(container, text="Subject:").grid(row=row, column=0, sticky="w", **pad)
        self.subject_var = tk.StringVar()
        self.subject_combo = ttk.Combobox(
            container, textvariable=self.subject_var, values=[], state="readonly")
        self.subject_combo.grid(row=row, column=1, sticky="ew", **pad)
        self.subject_combo.bind("<<ComboboxSelected>>", self.on_subject_change)
        row += 1

        # Subject code (auto, read-only)
        ttk.Label(container, text="Subject Code:").grid(row=row, column=0, sticky="w", **pad)
        self.code_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.code_var, state="readonly") \
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
            .grid(row=row, column=0, columnspan=2, sticky="w", padx=12, pady=(14, 2)); row += 1

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
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Generate DOCX", command=lambda: self.generate(["docx"])) \
            .grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="Generate PDF", command=lambda: self.generate(["pdf"])) \
            .grid(row=0, column=1, padx=6)
        ttk.Button(btn_frame, text="Generate Both", command=lambda: self.generate(["docx", "pdf"])) \
            .grid(row=0, column=2, padx=6)

        self.status_var = tk.StringVar(value="")
        ttk.Label(container, textvariable=self.status_var, foreground="green") \
            .grid(row=row + 1, column=0, columnspan=2)

        # init subject list for default vertical
        if VERTICALS:
            self.vertical_combo.current(0)
            self.on_vertical_change()

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

    def on_vertical_change(self, event=None):
        vertical = self.vertical_var.get()
        subjects = [name for name, code in VERTICALS.get(vertical, [])]
        self.subject_combo["values"] = subjects
        if subjects:
            self.subject_combo.current(0)
            self.on_subject_change()
        else:
            self.subject_var.set("")
            self.code_var.set("")

    def on_subject_change(self, event=None):
        vertical = self.vertical_var.get()
        subject = self.subject_var.get()
        for name, code in VERTICALS.get(vertical, []):
            if name == subject:
                self.code_var.set(code)
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
    verticals = list(VERTICALS.keys())
    for i, v in enumerate(verticals, 1):
        print(f"  {i}. {v}")
    v_idx = int(input("Choose vertical number: ")) - 1
    vertical = verticals[v_idx]

    subjects = VERTICALS[vertical]
    for i, (name, code) in enumerate(subjects, 1):
        print(f"  {i}. {name} ({code})")
    s_idx = int(input("Choose subject number: ")) - 1
    subject_name, subject_code = subjects[s_idx]

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
            print("No display available for the GUI — falling back to CLI mode.\n")
            run_cli()