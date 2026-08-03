import io
import os
import json
import datetime
import urllib.request
import streamlit as st
from PIL import Image

# --------------------------------------------------------------------------
# 1. VERTICAL / SUBJECT DATA
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
    "I have completed the assignment / academic activity with academic guidance or discussion from seniors / friends / peers, while ensuring the work is my own.",
    "I have used AI-based tools only as an aid and have appropriately acknowledged or cited their use as per the ethical guidelines of the Institute.",
    "I have used AI-based tools and directly copy-pasted content for the assignment / academic activity.",
    "I have copied the assignment / academic activity from peers / friends.",
]

LOGO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "college_logo.png")

ALL_SUBJECTS = []
for vert, subs in VERTICALS.items():
    for name, code in subs:
        ALL_SUBJECTS.append({"name": name, "code": code, "vertical": vert})

# --------------------------------------------------------------------------
# 2. SUPABASE REST HELPERS (Zero Extra Dependencies)
# --------------------------------------------------------------------------
def fetch_supabase_profile(url, key, roll_no):
    return None

def upsert_supabase_profile(url, key, profile):
    return

# --------------------------------------------------------------------------
# 3. DOCUMENT GENERATION ENGINE (Imported from gendoc.py)
# --------------------------------------------------------------------------
from gendoc import (
    process_signature_image,
    generate_docx_bytes,
    generate_pdf_bytes,
    merge_pdfs,
    merge_docxs,
)


# --------------------------------------------------------------------------
# 6. STREAMLIT WEB UI
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="VIT Declaration Form Generator",
    page_icon="📜",
    layout="centered"
)

# Custom Styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .stApp {
        background: linear-gradient(135deg, #0F172A 0%, #1E1B4B 50%, #0F172A 100%);
    }

    .main .block-container {
        max-width: 780px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    /* Force all labels, text and headings to bright white/slate */
    label, .stWidgetLabel, [data-testid="stWidgetLabel"] p, [data-testid="stMarkdownContainer"] p {
        color: #F8FAFC !important;
        font-weight: 600 !important;
    }

    h1, h2, h3, h4, h5, h6, .stSubheader {
        color: #F8FAFC !important;
        font-weight: 700 !important;
    }

    .vit-badge {
        display: inline-block;
        background: rgba(99, 102, 241, 0.2);
        color: #A5B4FC;
        font-weight: 700;
        font-size: 0.8rem;
        padding: 0.35rem 0.9rem;
        border-radius: 50px;
        letter-spacing: 0.5px;
        margin-bottom: 0.5rem;
    }

    .main-title {
        color: #FFFFFF;
        font-weight: 800;
        font-size: 2rem;
        letter-spacing: -0.5px;
        margin-bottom: 0.2rem;
    }

    .sub-title {
        color: #CBD5E1;
        font-size: 0.95rem;
        margin-bottom: 1.5rem;
    }

    /* Inputs background & text styling */
    input[type="text"], div[data-baseweb="select"] > div {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

header_col1, header_col2 = st.columns([0.82, 0.18])
with header_col1:
    st.markdown('<div class="vit-badge">🎓 VIDYALANKAR INSTITUTE OF TECHNOLOGY</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-title">Declaration Form Generator</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Generate and download official declaration forms (.docx & .pdf) formatted for VIT submissions.</div>', unsafe_allow_html=True)

with header_col2:
    with st.popover("⚙️", help="Customise Details & Add Subjects"):
        st.markdown("### ⚙️ Customise Profile & Subjects")
        
        # Section 1: Add Custom Subject
        st.caption("➕ **Add Custom Subject**")
        c_name = st.text_input("Subject Name", key="pop_c_name", placeholder="e.g. Robotics")
        c_code = st.text_input("Subject Code", key="pop_c_code", placeholder="e.g. ROB101")
        c_vert = st.text_input("Vertical", value="CUSTOM", key="pop_c_vert")
        
        if st.button("➕ Add Subject", key="btn_add_subj"):
            if c_name and c_code:
                if "custom_subjects" not in st.session_state:
                    st.session_state["custom_subjects"] = []
                st.session_state["custom_subjects"].append({"name": c_name, "code": c_code, "vertical": c_vert})
                st.toast(f"✅ Added subject: {c_name} ({c_code})")
                st.rerun()
            else:
                st.warning("Enter both Subject Name and Code.")

        st.markdown("---")
        
        # Section 2: Customise Profile Defaults
        st.caption("👤 **Customise Profile Defaults**")
        p_name = st.text_input("Full Name", value=st.session_state.get("student_name", ""), key="pop_p_name", placeholder="e.g. John Doe")
        p_branch = st.text_input("Branch", value=st.session_state.get("branch", ""), key="pop_p_branch", placeholder="e.g. CMPN")
        p_sem = st.text_input("Semester", value=st.session_state.get("semester", ""), key="pop_p_sem", placeholder="e.g. V")
        p_div = st.text_input("Division", value=st.session_state.get("division", ""), key="pop_p_div", placeholder="e.g. A")
        
        if st.button("💾 Save Profile Defaults", key="btn_save_prof"):
            st.session_state["student_name"] = p_name
            st.session_state["branch"] = p_branch
            st.session_state["semester"] = p_sem
            st.session_state["division"] = p_div
            
            # Sync to Supabase if configured
            sp_url = st.session_state.get("sb_url")
            sp_key = st.session_state.get("sb_key")
            curr_roll = st.session_state.get("roll_no", "")
            if sp_url and sp_key and curr_roll:
                upsert_supabase_profile(sp_url, sp_key, {
                    "roll_no": curr_roll,
                    "student_name": p_name,
                    "branch": p_branch,
                    "semester": p_sem,
                    "division": p_div,
                    "custom_subjects": json.dumps(st.session_state.get("custom_subjects", []))
                })
            st.toast("✅ Profile defaults saved and synced!")
            st.rerun()

# Supabase Config load from secrets / env
env_sb_url = os.environ.get("SUPABASE_URL", "")
env_sb_key = os.environ.get("SUPABASE_ANON_KEY", "")
try:
    if "SUPABASE_URL" in st.secrets:
        env_sb_url = st.secrets["SUPABASE_URL"]
    if "SUPABASE_ANON_KEY" in st.secrets:
        env_sb_key = st.secrets["SUPABASE_ANON_KEY"]
except Exception:
    pass

if env_sb_url and env_sb_key:
    st.session_state["sb_url"] = env_sb_url
    st.session_state["sb_key"] = env_sb_key

# Subject selection
with st.container():
    st.subheader("📚 Subject & Academic Details")
    all_subj_list = ALL_SUBJECTS + st.session_state.get("custom_subjects", [])
    subject_names = [s["name"] for s in all_subj_list]
    selected_subject_name = st.selectbox("Select Subject", options=subject_names)
    selected_subj_info = next((s for s in all_subj_list if s["name"] == selected_subject_name), ALL_SUBJECTS[0])

    col_sub1, col_sub2 = st.columns(2)
    with col_sub1:
        st.text_input("Subject Code", value=selected_subj_info["code"], disabled=True)
    with col_sub2:
        st.text_input("Vertical", value=selected_subj_info["vertical"], disabled=True)

    assignment_no = st.text_input("Assignment / Experiment No", value="Experiment 1")

st.markdown("---")

# Student Details
with st.container():
    st.subheader("👤 Student Details")
    
    # Roll number auto-fetch trigger
    roll_input = st.text_input("Roll No. (Enter to fetch profile)", value=st.session_state.get("roll_no", ""), key="roll_input_field", placeholder="e.g. 24102A0001")
    
    if roll_input != st.session_state.get("last_roll"):
        st.session_state["last_roll"] = roll_input
        # Fetch from Supabase if configured
        sp_url = st.session_state.get("sb_url")
        sp_key = st.session_state.get("sb_key")
        if sp_url and sp_key:
            remote = fetch_supabase_profile(sp_url, sp_key, roll_input.strip())
            if remote:
                st.session_state["student_name"] = remote.get("student_name", "")
                st.session_state["branch"] = remote.get("branch", "")
                st.session_state["semester"] = remote.get("semester", "")
                st.session_state["division"] = remote.get("division", "")
                if remote.get("custom_subjects"):
                    try:
                        c_sub = json.loads(remote["custom_subjects"]) if isinstance(remote["custom_subjects"], str) else remote["custom_subjects"]
                        if isinstance(c_sub, list):
                            st.session_state["custom_subjects"] = c_sub
                    except Exception:
                        pass
                st.toast(f"⚡ Synced profile for {roll_input} from Supabase!")

    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("Student Name", value=st.session_state.get("student_name", ""), placeholder="e.g. John Doe")
        branch = st.text_input("Branch", value=st.session_state.get("branch", ""))
        division = st.text_input("Division", value=st.session_state.get("division", ""))
    with col2:
        roll_no = roll_input
        semester = st.text_input("Semester", value=st.session_state.get("semester", ""))
        date_str = st.text_input("Date (DD-MM-YYYY)", value=datetime.date.today().strftime("%d-%m-%Y"))

st.markdown("---")

# Signature Section
with st.container():
    st.subheader("✍️ Signature Input")
    sig_method = st.radio("Choose Signature Method:", ["Draw Signature", "Upload Image File", "Skip"], horizontal=True)

    sig_bytes_io = None

    if sig_method == "Draw Signature":
        try:
            from streamlit_drawable_canvas import st_canvas
            st.caption("Use your mouse or touch pad to sign inside the box below:")
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)",
                stroke_width=3,
                stroke_color="#0F172A",
                background_color="#FFFFFF",
                height=140,
                width=500,
                drawing_mode="freedraw",
                key="canvas_sig",
            )
            if canvas_result.image_data is not None:
                img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                sig_bytes_io = buf
        except Exception as e:
            st.error(f"Canvas error: {e}")

    elif sig_method == "Upload Image File":
        uploaded_sig = st.file_uploader("Upload Signature (PNG / JPG)", type=["png", "jpg", "jpeg"])
        if uploaded_sig is not None:
            sig_bytes_io = io.BytesIO(uploaded_sig.getvalue())
            st.image(uploaded_sig, caption="Uploaded Signature Preview", width=180)

st.markdown("---")

# Declaration checkboxes
with st.container():
    st.subheader("✅ Ethical Declaration Statements")
    st.write("Tick applicable statement(s):")

    selected_options = []
    for idx, option_text in enumerate(DECLARATION_OPTIONS, start=1):
        checked = st.checkbox(f"({idx}) {option_text}", value=(idx == 1), key=f"opt_{idx}")
        if checked:
            selected_options.append(idx)

st.markdown("---")

with st.container():
    st.subheader("📁 Attachment & Output Options (Optional)")
    user_file = st.file_uploader("Upload your assignment file (PDF or DOCX)", type=["pdf", "docx"])
    safe_assignment = assignment_no.strip().replace(" ", "").replace("/", "-")
    default_filename = f"{selected_subj_info['code']}-{safe_assignment}-Declaration Form"
    custom_name = st.text_input("Custom Output Filename (without extension)", value=default_filename)

st.markdown("---")


# Generate buttons
if not selected_options:
    st.warning("⚠️ Please select at least one declaration statement above.")
else:
    form_data = {
        "vertical": selected_subj_info["vertical"],
        "subject_name": selected_subj_info["name"],
        "subject_code": selected_subj_info["code"],
        "assignment_no": assignment_no.strip(),
        "student_name": student_name.strip(),
        "roll_no": roll_no.strip(),
        "branch": branch.strip(),
        "semester": semester.strip(),
        "division": division.strip(),
        "date": date_str.strip(),
        "selected_options": selected_options,
        "signature_bytes": sig_bytes_io,
    }

    # Save to Supabase background task
    sp_url = st.session_state.get("sb_url")
    sp_key = st.session_state.get("sb_key")
    if sp_url and sp_key:
        upsert_supabase_profile(sp_url, sp_key, {
            "roll_no": roll_no.strip(),
            "student_name": student_name.strip(),
            "branch": branch.strip(),
            "semester": semester.strip(),
            "division": division.strip(),
            "last_subject_code": selected_subj_info["code"],
            "last_subject_name": selected_subj_info["name"],
            "updated_at": datetime.datetime.now().isoformat()
        })

    final_filename = custom_name.strip() if custom_name.strip() else default_filename
    if final_filename.lower().endswith(".pdf"):
        final_filename = final_filename[:-4]
    elif final_filename.lower().endswith(".docx"):
        final_filename = final_filename[:-5]

    st.subheader("📥 Download Generated Declaration Form")
    col_dl1, col_dl2 = st.columns(2)

    if user_file is not None:
        file_ext = user_file.name.split(".")[-1].lower()
        if file_ext == "pdf":
            with st.spinner("✨ Merging declaration with uploaded PDF..."):
                dec_pdf = generate_pdf_bytes(form_data)
                pdf_buffer = merge_pdfs(dec_pdf.getvalue(), user_file.getvalue())
            with col_dl2:
                st.download_button(
                    label="📄 Download Merged PDF (.pdf)",
                    data=pdf_buffer,
                    file_name=f"{final_filename}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    on_click=lambda: st.toast("✅ Merged PDF Downloaded!", icon="📄")
                )
            with col_dl1:
                st.info("💡 Attached file is a PDF. Output is PDF only.")
        elif file_ext == "docx":
            with st.spinner("✨ Merging declaration with uploaded DOCX..."):
                dec_docx = generate_docx_bytes(form_data)
                docx_buffer = merge_docxs(dec_docx.getvalue(), user_file.getvalue())
            with col_dl1:
                st.download_button(
                    label="📝 Download Merged Word (.docx)",
                    data=docx_buffer,
                    file_name=f"{final_filename}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    on_click=lambda: st.toast("✅ Merged DOCX Downloaded!", icon="📝")
                )
            with col_dl2:
                st.info("💡 Attached file is a Word document. Output is DOCX only.")
    else:
        with st.spinner("✨ Rendering official declaration layout & embedding signature..."):
            docx_buffer = generate_docx_bytes(form_data)
            pdf_buffer = generate_pdf_bytes(form_data)

        with col_dl1:
            st.download_button(
                label="📝 Download Word (.docx)",
                data=docx_buffer,
                file_name=f"{final_filename}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                on_click=lambda: st.toast("✅ .DOCX Declaration Form Downloaded!", icon="📝")
            )

        with col_dl2:
            st.download_button(
                label="📄 Download PDF (.pdf)",
                data=pdf_buffer,
                file_name=f"{final_filename}.pdf",
                mime="application/pdf",
                use_container_width=True,
                on_click=lambda: st.toast("✅ .PDF Declaration Form Downloaded!", icon="📄")
            )

