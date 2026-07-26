import os
import sys
import json
import urllib.request
import urllib.error

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, PROJECT_ROOT)

from flask_app import SUPABASE_URL, SUPABASE_ANON_KEY

def diagnose():
    print(f"URL: {SUPABASE_URL}")
    print(f"Key: {SUPABASE_ANON_KEY[:10]}...")

    endpoint = f"{SUPABASE_URL.rstrip('/')}/rest/v1/student_profiles"
    payload_data = [{
        "roll_no": "24000A0001",
        "student_name": "John Doe",
        "branch": "BRANCH",
        "semester": "V",
        "division": "A",
        "last_subject_code": "PCCE10T",
        "last_subject_name": "Artificial Intelligence",
        "updated_at": "2026-07-21T01:44:00Z"
    }]
    payload = json.dumps(payload_data).encode("utf-8")

    req = urllib.request.Request(endpoint, data=payload, headers={
        "apikey": SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates"
    }, method="POST")

    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            print(f"POST Status: {resp.status}")
            print(f"Response: {resp.read().decode()}")
    except urllib.error.HTTPError as e:
        print(f"HTTPError: {e.code} - {e.reason}")
        print(f"Error Body: {e.read().decode()}")
    except Exception as ex:
        print(f"Exception: {ex}")

if __name__ == "__main__":
    diagnose()
