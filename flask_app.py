import io
import os
import base64
import json
import datetime
import urllib.request
from flask import Flask, render_template, request, send_file, jsonify
from app import VERTICALS, ALL_SUBJECTS, DECLARATION_OPTIONS, generate_docx_bytes, generate_pdf_bytes

app = Flask(__name__)

# Load credentials from .env or env/.env
def load_env():
    for env_path in [".env", os.path.join("env", ".env")]:
        if os.path.exists(env_path):
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip().strip("'").strip('"'))

load_env()

import re

def normalize_supabase_url(url):
    if not url:
        return ""
    if "db." in url and ".supabase.co" in url:
        match = re.search(r"db\.([a-z0-9]+)\.supabase\.co", url)
        if match:
            return f"https://{match.group(1)}.supabase.co"
    if not url.startswith("http"):
        return f"https://{url}"
    return url

RAW_SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_URL = normalize_supabase_url(RAW_SUPABASE_URL)
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")

def fetch_profile_from_supabase(roll_no):
    if not SUPABASE_URL or not SUPABASE_ANON_KEY or not roll_no:
        return None
    try:
        endpoint = f"{SUPABASE_URL.rstrip('/')}/rest/v1/student_profiles?roll_no=eq.{roll_no}&select=*"
        req = urllib.request.Request(endpoint, headers={
            "apikey": SUPABASE_ANON_KEY,
            "Authorization": f"Bearer {SUPABASE_ANON_KEY}"
        })
        with urllib.request.urlopen(req, timeout=4) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data[0] if data else None
    except Exception:
        return None

def upsert_profile_to_supabase(profile_data):
    if not SUPABASE_URL or not SUPABASE_ANON_KEY or not profile_data.get("roll_no"):
        return False
    try:
        endpoint = f"{SUPABASE_URL.rstrip('/')}/rest/v1/student_profiles"
        payload = json.dumps([profile_data]).encode("utf-8")
        req = urllib.request.Request(endpoint, data=payload, headers={
            "apikey": SUPABASE_ANON_KEY,
            "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates"
        }, method="POST")
        with urllib.request.urlopen(req, timeout=4) as resp:
            return True
    except Exception:
        return False

@app.route("/", methods=["GET"])
def index():
    today = datetime.date.today().strftime("%d-%m-%Y")
    return render_template(
        "index.html",
        subjects=ALL_SUBJECTS,
        declaration_options=DECLARATION_OPTIONS,
        default_date=today
    )

@app.route("/api/student/<roll_no>", methods=["GET"])
def get_student_profile(roll_no):
    profile = fetch_profile_from_supabase(roll_no.strip())
    if profile:
        return jsonify({"success": True, "profile": profile})
    return jsonify({"success": False, "message": "Profile not found"}), 404

@app.route("/api/student", methods=["POST"])
def save_student_profile():
    data = request.json or {}
    if not data.get("roll_no"):
        return jsonify({"success": False, "message": "Roll No required"}), 400
    success = upsert_profile_to_supabase(data)
    return jsonify({"success": success})

@app.route("/generate", methods=["POST"])
def generate():
    s_idx = int(request.form.get("subject_idx", 0))
    if s_idx < len(ALL_SUBJECTS):
        subj_info = ALL_SUBJECTS[s_idx]
        vertical_val = subj_info["vertical"]
        subj_name_val = subj_info["name"]
        subj_code_val = subj_info["code"]
    else:
        vertical_val = request.form.get("vertical", "CUSTOM")
        subj_name_val = request.form.get("subject_name", "Custom Subject")
        subj_code_val = request.form.get("subject_code", "CUSTOM")

    assignment_no = request.form.get("assignment_no", "Experiment 1").strip()
    student_name = request.form.get("student_name", "").strip()
    roll_no = request.form.get("roll_no", "").strip()
    branch = request.form.get("branch", "").strip()
    semester = request.form.get("semester", "").strip()
    division = request.form.get("division", "").strip()
    date_str = request.form.get("date", "").strip()

    raw_options = request.form.getlist("selected_options")
    selected_options = [int(x) for x in raw_options if x.isdigit()] or [1]

    fmt = request.form.get("format", "docx").lower()

    # Check drawn signature data URL or uploaded signature file
    sig_bytes_io = None
    sig_draw_data = request.form.get("signature_draw_data", "").strip()

    if sig_draw_data and sig_draw_data.startswith("data:image"):
        header, encoded = sig_draw_data.split(",", 1)
        sig_bytes_io = io.BytesIO(base64.b64decode(encoded))
    else:
        sig_file = request.files.get("signature")
        if sig_file and sig_file.filename != "":
            sig_bytes_io = io.BytesIO(sig_file.read())

    # Upsert to Supabase securely on backend
    if roll_no:
        upsert_profile_to_supabase({
            "roll_no": roll_no,
            "student_name": student_name,
            "branch": branch,
            "semester": semester,
            "division": division,
            "last_subject_code": subj_code_val,
            "last_subject_name": subj_name_val,
            "updated_at": datetime.datetime.now().isoformat()
        })

    form_data = {
        "vertical": vertical_val,
        "subject_name": subj_name_val,
        "subject_code": subj_code_val,
        "assignment_no": assignment_no,
        "student_name": student_name,
        "roll_no": roll_no,
        "branch": branch,
        "semester": semester,
        "division": division,
        "date": date_str,
        "selected_options": selected_options,
        "signature_bytes": sig_bytes_io,
    }

    safe_assignment = assignment_no.replace(" ", "").replace("/", "-")
    base_filename = f"{subj_code_val}-{safe_assignment}-Declaration Form"

    if fmt == "pdf":
        buf = generate_pdf_bytes(form_data)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"{base_filename}.pdf",
            mimetype="application/pdf"
        )
    else:
        buf = generate_docx_bytes(form_data)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"{base_filename}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    app.run(debug=True, port=5000)
