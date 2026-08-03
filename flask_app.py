import io
import os
import base64
import json
import datetime
import urllib.request
from flask import Flask, render_template, request, send_file, jsonify
from gendoc import generate_docx_bytes, generate_pdf_bytes, merge_pdfs, merge_docxs
from app import VERTICALS, ALL_SUBJECTS, DECLARATION_OPTIONS

app = Flask(__name__, static_folder="public", static_url_path="")

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, apikey"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response

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
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            if data:
                profile = data[0]
                if isinstance(profile.get("custom_subjects"), str):
                    try:
                        profile["custom_subjects"] = json.loads(profile["custom_subjects"])
                    except Exception:
                        pass
                return profile
            return None
    except Exception as e:
        print(f"fetch_profile_from_supabase error: {e}")
        return None

def upsert_profile_to_supabase(profile_data):
    if not SUPABASE_URL or not SUPABASE_ANON_KEY or not profile_data.get("roll_no"):
        return False
    try:
        payload_data = dict(profile_data)
        if "custom_subjects" in payload_data:
            if isinstance(payload_data["custom_subjects"], str):
                try:
                    payload_data["custom_subjects"] = json.loads(payload_data["custom_subjects"])
                except Exception:
                    pass
        endpoint = f"{SUPABASE_URL.rstrip('/')}/rest/v1/student_profiles"
        payload = json.dumps([payload_data]).encode("utf-8")
        req = urllib.request.Request(endpoint, data=payload, headers={
            "apikey": SUPABASE_ANON_KEY,
            "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates"
        }, method="POST")
        with urllib.request.urlopen(req, timeout=5) as resp:
            resp.read()
            return True
    except Exception as e:
        print(f"upsert_profile_to_supabase error: {e}")
        return False

import traceback

@app.errorhandler(500)
def handle_500(e):
    return f"<pre>{traceback.format_exc()}</pre>", 500

@app.errorhandler(Exception)
def handle_exception(e):
    return f"<pre>{traceback.format_exc()}</pre>", 500

@app.route("/", methods=["GET"])
def index():
    today = datetime.date.today().strftime("%d-%m-%Y")
    return render_template(
        "index.html",
        subjects=ALL_SUBJECTS,
        declaration_options=DECLARATION_OPTIONS,
        default_date=today
    )

@app.route("/api/subjects", methods=["GET"])
def get_subjects():
    today = datetime.date.today().strftime("%d-%m-%Y")
    return jsonify({
        "success": True,
        "subjects": ALL_SUBJECTS,
        "declaration_options": DECLARATION_OPTIONS,
        "default_date": today
    })

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
@app.route("/api/generate", methods=["POST"])
def generate():
    if request.is_json:
        req_data = request.get_json() or {}
        s_idx = int(req_data.get("subject_idx", 0))
        if s_idx < len(ALL_SUBJECTS):
            subj_info = ALL_SUBJECTS[s_idx]
            vertical_val = subj_info["vertical"]
            subj_name_val = subj_info["name"]
            subj_code_val = subj_info["code"]
        else:
            vertical_val = req_data.get("vertical", "CUSTOM")
            subj_name_val = req_data.get("subject_name", "Custom Subject")
            subj_code_val = req_data.get("subject_code", "CUSTOM")

        assignment_no = str(req_data.get("assignment_no", "Experiment 1")).strip()
        student_name = str(req_data.get("student_name", "")).strip()
        roll_no = str(req_data.get("roll_no", "")).strip()
        branch = str(req_data.get("branch", "")).strip()
        semester = str(req_data.get("semester", "")).strip()
        division = str(req_data.get("division", "")).strip()
        date_str = str(req_data.get("date", "")).strip()
        raw_options = req_data.get("selected_options", [1])
        selected_options = [int(x) for x in raw_options if str(x).isdigit()] or [1]
        fmt = str(req_data.get("format", "docx")).lower()
        sig_draw_data = str(req_data.get("signature_draw_data", "")).strip()
        custom_subjects = req_data.get("custom_subjects", [])
        output_filename = str(req_data.get("output_filename", "")).strip()
    else:
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
        selected_options = [int(x) for x in raw_options if str(x).isdigit()] or [1]
        fmt = request.form.get("format", "docx").lower()
        sig_draw_data = request.form.get("signature_draw_data", "").strip()
        custom_subjects = request.form.get("custom_subjects", "[]")
        output_filename = request.form.get("output_filename", "").strip()

    # Check drawn signature data URL or uploaded signature file
    sig_bytes_io = None

    if sig_draw_data and sig_draw_data.startswith("data:image"):
        header, encoded = sig_draw_data.split(",", 1)
        sig_bytes_io = io.BytesIO(base64.b64decode(encoded))
    elif not request.is_json:
        sig_file = request.files.get("signature")
        if sig_file and sig_file.filename != "":
            sig_bytes_io = io.BytesIO(sig_file.read())

    # Check for user-uploaded file for merging
    user_file_bytes = None
    user_file_name = None
    if not request.is_json:
        user_file = request.files.get("user_file")
        if user_file and user_file.filename != "":
            user_file_bytes = user_file.read()
            user_file_name = user_file.filename

    # Validate file type and format match
    if user_file_name:
        ext = os.path.splitext(user_file_name.lower())[1]
        if ext == ".pdf" and fmt != "pdf":
            return "Error: Attached file is PDF but requested format is not PDF. Please download as PDF.", 400
        if ext == ".docx" and fmt != "docx":
            return "Error: Attached file is DOCX but requested format is not DOCX. Please download as Word (.docx).", 400

    # Upsert student profile to remote Supabase database
    if roll_no:
        upsert_profile_to_supabase({
            "roll_no": roll_no,
            "student_name": student_name,
            "branch": branch,
            "semester": semester,
            "division": division,
            "last_subject_code": subj_code_val,
            "last_subject_name": subj_name_val,
            "custom_subjects": custom_subjects,
            "signature_data": sig_draw_data if (sig_draw_data and sig_draw_data.startswith("data:image")) else "",
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

    if output_filename:
        # Strip extension if user included it
        if output_filename.lower().endswith(".pdf"):
            output_filename = output_filename[:-4]
        elif output_filename.lower().endswith(".docx"):
            output_filename = output_filename[:-5]
        base_filename = output_filename
    else:
        safe_assignment = assignment_no.replace(" ", "").replace("/", "-")
        base_filename = f"{subj_code_val}-{safe_assignment}-Declaration Form"

    if fmt == "pdf":
        buf = generate_pdf_bytes(form_data)
        if user_file_bytes and user_file_name.lower().endswith(".pdf"):
            buf = merge_pdfs(buf.getvalue(), user_file_bytes)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"{base_filename}.pdf",
            mimetype="application/pdf"
        )
    else:
        buf = generate_docx_bytes(form_data)
        if user_file_bytes and user_file_name.lower().endswith(".docx"):
            buf = merge_docxs(buf.getvalue(), user_file_bytes)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"{base_filename}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


def get_asset_path(filename):
    # Try static_assets first (packaged in Vercel lambda environment)
    path = os.path.join("static_assets", filename)
    if os.path.exists(path):
        return path
    # Fallback to public (used locally)
    return os.path.join("public", filename)

@app.route("/ads.txt")
def serve_ads():
    return send_file(get_asset_path("ads.txt"), mimetype="text/plain")

@app.route("/robots.txt")
def serve_robots():
    return send_file(get_asset_path("robots.txt"), mimetype="text/plain")

@app.route("/sitemap.xml")
def serve_sitemap():
    return send_file(get_asset_path("sitemap.xml"), mimetype="application/xml")

@app.route("/dg-logo.png")
def serve_logo():
    return send_file(get_asset_path("dg-logo.png"), mimetype="image/png")

@app.route("/OG-image-dg.png")
def serve_og_image():
    return send_file(get_asset_path("OG-image-dg.png"), mimetype="image/png")

@app.route("/favicon.ico")
def serve_favicon():
    return send_file(get_asset_path("favicon.ico"), mimetype="image/x-icon")

@app.route("/favicon.svg")
def serve_favicon_svg():
    return send_file(get_asset_path("favicon.svg"), mimetype="image/svg+xml")

@app.route("/favicon-96x96.png")
def serve_favicon_96():
    return send_file(get_asset_path("favicon-96x96.png"), mimetype="image/png")

@app.route("/apple-touch-icon.png")
def serve_apple_icon():
    return send_file(get_asset_path("apple-touch-icon.png"), mimetype="image/png")

@app.route("/site.webmanifest")
def serve_manifest():
    return send_file(get_asset_path("site.webmanifest"), mimetype="application/manifest+json")

@app.route("/web-app-manifest-192x192.png")
def serve_manifest_192():
    return send_file(get_asset_path("web-app-manifest-192x192.png"), mimetype="image/png")

@app.route("/web-app-manifest-512x512.png")
def serve_manifest_512():
    return send_file(get_asset_path("web-app-manifest-512x512.png"), mimetype="image/png")

if __name__ == "__main__":

    app.run(debug=True, port=5000)

