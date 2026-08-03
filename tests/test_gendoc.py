import os
import sys
import io
import unittest
from PIL import Image

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from gendoc import (
    process_signature_image,
    generate_docx_bytes,
    generate_pdf_bytes,
    merge_pdfs,
    merge_docxs,
)
from app import (
    DECLARATION_OPTIONS,
    ALL_SUBJECTS
)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

class TestGendoc(unittest.TestCase):

    def test_process_signature_image(self):
        img = Image.new("RGB", (200, 80), (255, 255, 255))
        img_bytes = io.BytesIO()
        img.save(img_bytes, format="PNG")
        img_bytes.seek(0)

        processed_io, target_w, target_h = process_signature_image(img_bytes)
        self.assertIsNotNone(processed_io)
        self.assertLessEqual(target_w, 0.80)
        self.assertLessEqual(target_h, 0.32)

    def test_generate_docx_bytes(self):
        form_data = {
            "vertical": "PC_PEC",
            "subject_name": "Machine Learning",
            "subject_code": "PCCC05T",
            "assignment_no": "Experiment 1",
            "student_name": "Test Student",
            "roll_no": "24102A0000",
            "branch": "CMPN",
            "semester": "V",
            "division": "A",
            "date": "20-07-2026",
            "selected_options": [1, 2, 3],
            "signature_bytes": None
        }

        docx_buf = generate_docx_bytes(form_data)
        self.assertIsNotNone(docx_buf)
        self.assertGreater(len(docx_buf.getvalue()), 1000)

        # Save sample file to tests/output/
        sample_path = os.path.join(OUTPUT_DIR, "sample_declaration.docx")
        try:
            with open(sample_path, "wb") as f:
                f.write(docx_buf.getvalue())
            self.assertTrue(os.path.exists(sample_path))
        except PermissionError:
            pass

    def test_generate_pdf_bytes(self):
        form_data = {
            "vertical": "PC_PEC",
            "subject_name": "Machine Learning",
            "subject_code": "PCCC05T",
            "assignment_no": "Experiment 1",
            "student_name": "Test Student",
            "roll_no": "24102A0000",
            "branch": "CMPN",
            "semester": "V",
            "division": "A",
            "date": "20-07-2026",
            "selected_options": [1, 2, 3],
            "signature_bytes": None
        }

        pdf_buf = generate_pdf_bytes(form_data)
        self.assertIsNotNone(pdf_buf)
        self.assertGreater(len(pdf_buf.getvalue()), 1000)

        # Save sample file to tests/output/
        sample_path = os.path.join(OUTPUT_DIR, "sample_declaration.pdf")
        with open(sample_path, "wb") as f:
            f.write(pdf_buf.getvalue())
        self.assertTrue(os.path.exists(sample_path))

    def test_merge_pdfs(self):
        form_data = {
            "vertical": "PC_PEC",
            "subject_name": "Machine Learning",
            "subject_code": "PCCC05T",
            "assignment_no": "Experiment 1",
            "student_name": "Test Student",
            "roll_no": "24102A0000",
            "branch": "CMPN",
            "semester": "V",
            "division": "A",
            "date": "20-07-2026",
            "selected_options": [1, 2, 3],
            "signature_bytes": None
        }
        dec_pdf_buf = generate_pdf_bytes(form_data)
        
        merged_buf = merge_pdfs(dec_pdf_buf.getvalue(), dec_pdf_buf.getvalue())
        self.assertIsNotNone(merged_buf)
        self.assertGreater(len(merged_buf.getvalue()), len(dec_pdf_buf.getvalue()))
        
        merged_path = os.path.join(OUTPUT_DIR, "merged_test_output.pdf")
        with open(merged_path, "wb") as f:
            f.write(merged_buf.getvalue())
        self.assertTrue(os.path.exists(merged_path))

    def test_merge_docxs(self):
        form_data = {
            "vertical": "PC_PEC",
            "subject_name": "Machine Learning",
            "subject_code": "PCCC05T",
            "assignment_no": "Experiment 1",
            "student_name": "Test Student",
            "roll_no": "24102A0000",
            "branch": "CMPN",
            "semester": "V",
            "division": "A",
            "date": "20-07-2026",
            "selected_options": [1, 2, 3],
            "signature_bytes": None
        }
        dec_docx_buf = generate_docx_bytes(form_data)
        
        merged_buf = merge_docxs(dec_docx_buf.getvalue(), dec_docx_buf.getvalue())
        self.assertIsNotNone(merged_buf)
        
        merged_path = os.path.join(OUTPUT_DIR, "merged_test_output.docx")
        with open(merged_path, "wb") as f:
            f.write(merged_buf.getvalue())
        self.assertTrue(os.path.exists(merged_path))

if __name__ == "__main__":
    unittest.main()

