import os
import sys
import json
import unittest

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from flask_app import app, normalize_supabase_url

class TestFlaskApp(unittest.TestCase):

    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()

    def test_normalize_supabase_url(self):
        # Direct DB connection string
        db_url = "postgresql://postgres:pass@db.fomsnceucuurjvsuzeja.supabase.co:5432/postgres"
        expected = "https://fomsnceucuurjvsuzeja.supabase.co"
        self.assertEqual(normalize_supabase_url(db_url), expected)

        # Standard HTTPS URL
        https_url = "https://fomsnceucuurjvsuzeja.supabase.co"
        self.assertEqual(normalize_supabase_url(https_url), expected)

    def test_index_route(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Ethical Practice Undertaking", response.data)
        self.assertIn(b"Customise Profile", response.data)

    def test_generate_docx_route(self):
        payload = {
            "subject_idx": "0",
            "assignment_no": "Experiment 1",
            "student_name": "Test Student",
            "roll_no": "24102A0000",
            "branch": "CMPN",
            "semester": "V",
            "division": "A",
            "date": "20-07-2026",
            "selected_options": ["1", "2"],
            "format": "docx"
        }
        response = self.client.post("/generate", data=payload)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    def test_generate_pdf_route(self):
        payload = {
            "subject_idx": "0",
            "assignment_no": "Experiment 1",
            "student_name": "Test Student",
            "roll_no": "24102A0000",
            "branch": "CMPN",
            "semester": "V",
            "division": "A",
            "date": "20-07-2026",
            "selected_options": ["1"],
            "format": "pdf"
        }
        response = self.client.post("/generate", data=payload)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/pdf")

    def test_api_student_invalid_roll(self):
        response = self.client.get("/api/student/9999999999NONEXISTENT")
        self.assertIn(response.status_code, [404, 200])

if __name__ == "__main__":
    unittest.main()
