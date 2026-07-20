import os
import sys
import json
import time

# Ensure project root directory is in sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, PROJECT_ROOT)

from flask_app import app, fetch_profile_from_supabase, upsert_profile_to_supabase, SUPABASE_URL, SUPABASE_ANON_KEY

def run_supabase_test():
    print("=" * 60)
    print("SUPABASE LIVE INTEGRATION TEST")
    print("=" * 60)
    print(f"Supabase Endpoint: {SUPABASE_URL}")
    print(f"Anon Key Present:  {'Yes' if SUPABASE_ANON_KEY else 'No'}")
    print("-" * 60)

    test_roll = "24102A9999"
    test_profile = {
        "roll_no": test_roll,
        "student_name": "Supabase Live Test Student",
        "branch": "CMPN",
        "semester": "V",
        "division": "B",
        "last_subject_code": "PCCE10T",
        "last_subject_name": "Artificial Intelligence",
        "custom_subjects": json.dumps([{"name": "Quantum Computing", "code": "QC101", "vertical": "CUSTOM"}]),
        "signature_data": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%SZ")
    }

    print(f"1. Attempting to UPSERT student profile for Roll No: {test_roll}...")
    success = upsert_profile_to_supabase(test_profile)

    if success:
        print("   SUCCESS: Profile upserted to Supabase table 'student_profiles'!")
    else:
        print("   FAILURE: Unable to upsert to Supabase. Check table existence or credentials.")

    print("\n2. Attempting to FETCH student profile from Supabase...")
    retrieved = fetch_profile_from_supabase(test_roll)

    if retrieved:
        print("   SUCCESS: Retrieved profile from Supabase REST API!")
        print("   Retrieved Data:")
        print(f"   - Roll No:       {retrieved.get('roll_no')}")
        print(f"   - Full Name:     {retrieved.get('student_name')}")
        print(f"   - Branch/Sem:    {retrieved.get('branch')} / {retrieved.get('semester')}")
        print(f"   - Last Subject:  {retrieved.get('last_subject_name')} ({retrieved.get('last_subject_code')})")
        print(f"   - Signature:     {retrieved.get('signature_data')[:40]}...")
    else:
        print("   FAILURE: Profile not returned from Supabase API.")

    print("=" * 60)

if __name__ == "__main__":
    run_supabase_test()
