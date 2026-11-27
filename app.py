import streamlit as st
import pandas as pd
from supabase import create_client, Client
from datetime import datetime
import io
import zipfile
import os
import base64
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="KEA Comprehensive School Management System",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SUPABASE CONNECTION SETUP ---
@st.cache_resource
def init_supabase():
    try:
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except Exception as e:
        st.error(f"Failed to connect to Supabase. Check secrets.toml configuration. Details: {e}")
        return None

supabase = init_supabase()

# =========================================================
# --- SUBSCRIPTION & SYSTEM LICENSE GATEKEEPER ---
# =========================================================
def check_subscription():
    if not supabase:
        return
    try:
        response = (
            supabase.table("school_subscriptions")
            .select("*")
            .eq("school_id", "KEA_COMPREHENSIVE")
            .execute()
        )
        if response.data:
            sub_info = response.data[0]
            expiry_str = sub_info.get("expiry_date")
            status = sub_info.get("status", "INACTIVE").upper()
            
            if expiry_str:
                expiry_date = datetime.strptime(expiry_str, "%Y-%m-%d").date()
                today = datetime.now().date()
                
                if status != "ACTIVE" or today > expiry_date:
                    st.error("🔒 **System Access Expired**")
                    st.warning(
                        "The subscription license for KEA Comprehensive School SMS has ended or is inactive. "
                        "Please contact the system developer to renew your access."
                    )
                    st.stop()
    except Exception:
        pass

check_subscription()

# --- INITIALIZE SESSION STATE ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.full_name = ""

# --- CURRICULUM & LEARNING AREAS DEFINITIONS ---
LOWER_PRIMARY_SUBJECTS = [
    "MATHEMATICS", "ENGLISH", "KISWAHILI",
    "ENVIRONMENTAL", "RELIGIOUS EDUCATION", "CREATIVE ARTS"
]

UPPER_PRIMARY_SUBJECTS = [
    "MATHEMATICS", "ENGLISH", "KISWAHILI",
    "SCIENCE AND TECHNOLOGY", "AGRICULTURE AND NUTRITION",
    "RELIGIOUS EDUCATION", "CREATIVE ARTS", "SOCIAL STUDIES"
]

JUNIOR_SCHOOL_SUBJECTS = [
    "MATHEMATICS", "ENGLISH", "KISWAHILI", "INTEGRATED SCIENCE",
    "AGRICULTURE", "PRETECHNICAL STUDIES", "SOCIAL STUDIES",
    "CHRISTIAN RELIGIOUS EDUCATION", "CREATIVE ARTS SPORTS"
]

def get_learning_areas_for_grade(grade_str):
    g_lower = grade_str.strip().lower()
    if any(g in g_lower for g in ["grade 1", "grade 2", "grade 3"]):
        return LOWER_PRIMARY_SUBJECTS
    elif any(g in g_lower for g in ["grade 4", "grade 5", "grade 6"]):
        return UPPER_PRIMARY_SUBJECTS
    else:
        return JUNIOR_SCHOOL_SUBJECTS

SUBJECT_ABBR_MAP = {
    "MATHEMATICS": "MATHS",
    "ENGLISH": "ENG",
    "KISWAHILI": "KIS",
    "ENVIRONMENTAL": "ENV",
    "RELIGIOUS EDUCATION": "RE",
    "CREATIVE ARTS": "CA",
    "SCIENCE AND TECHNOLOGY": "SCI & TECH",
    "AGRICULTURE AND NUTRITION": "AGRI & NUTR",
    "SOCIAL STUDIES": "SOCIAL",
    "INTEGRATED SCIENCE": "INT SCI",
    "AGRICULTURE": "AGRIC",
    "PRETECHNICAL STUDIES": "PRETECH",
    "CHRISTIAN RELIGIOUS EDUCATION": "CRE",
    "CREATIVE ARTS SPORTS": "CAS"
}

SUBJECT_COLUMN_MAP = {
    "CREATIVE ARTS SPORTS": "creative_arts_sports",
    "CHRISTIAN RELIGIOUS EDUCATION": "christian_religious_education",
    "PRETECHNICAL STUDIES": "pretechnical_studies",
    "INTEGRATED SCIENCE": "integrated_science",
    "SOCIAL STUDIES": "social_studies",
    "SCIENCE AND TECHNOLOGY": "science_and_technology",
    "AGRICULTURE AND NUTRITION": "agriculture_and_nutrition",
    "RELIGIOUS EDUCATION": "religious_education",
    "CREATIVE ARTS": "creative_arts",
    "ENVIRONMENTAL": "environmental"
}

def get_db_col_name(subject_name):
    subject_clean = subject_name.strip().upper()
    if subject_clean in SUBJECT_COLUMN_MAP:
        return SUBJECT_COLUMN_MAP[subject_clean]
    return subject_name.lower().strip().replace(" ", "_").replace(".", "").replace("(", "").replace(")", "")

# --- TEACHER SUBJECT & GRADE ASSIGNMENTS ---
TEACHER_ASSIGNMENTS = {
    "Opondo": {
        "assignments": [
            {"grade": "Grade 7", "subjects": ["ENGLISH", "PRETECHNICAL STUDIES"]},
            {"grade": "Grade 8", "subjects": ["ENGLISH"]},
            {"grade": "Grade 9", "subjects": ["ENGLISH"]}
        ]
    },
    "Lucas": {
        "class_teacher": "Grade 9",
        "assignments": [
            {"grade": "Grade 9", "subjects": ["MATHEMATICS", "AGRICULTURE"]},
            {"grade": "Grade 8", "subjects": ["INTEGRATED SCIENCE", "AGRICULTURE"]}
        ]
    },
    "Vincent": {
        "class_teacher": "Grade 8",
        "assignments": [
            {"grade": "Grade 8", "subjects": ["SOCIAL STUDIES", "KISWAHILI"]},
            {"grade": "Grade 7", "subjects": ["SOCIAL STUDIES"]},
            {"grade": "Grade 9", "subjects": ["SOCIAL STUDIES", "KISWAHILI"]}
        ]
    },
    "Grace": {
        "class_teacher": "Grade 7",
        "assignments": [
            {"grade": "Grade 7", "subjects": ["MATHEMATICS", "AGRICULTURE"]},
            {"grade": "Grade 9", "subjects": ["INTEGRATED SCIENCE"]}
        ]
    },
    "Achiyo": {
        "assignments": [
            {"grade": "Grade 9", "subjects": ["CHRISTIAN RELIGIOUS EDUCATION"]}
        ]
    },
    "Balla": {
        "assignments": [
            {"grade": "Grade 8", "subjects": ["CREATIVE ARTS SPORTS"]}
        ]
    },
    "Faith": {
        "assignments": [
            {"grade": "Grade 7", "subjects": ["INTEGRATED SCIENCE", "KISWAHILI"]},
            {"grade": "Grade 8", "subjects": ["CHRISTIAN RELIGIOUS EDUCATION"]}
        ]
    },
    "Elijah": {
        "assignments": [
            {"grade": "Grade 7", "subjects": ["CREATIVE ARTS SPORTS"]},
            {"grade": "Grade 8", "subjects": ["MATHEMATICS", "PRETECHNICAL STUDIES"]},
            {"grade": "Grade 9", "subjects": ["PRETECHNICAL STUDIES", "CREATIVE ARTS SPORTS"]}
        ]
    }
}

# --- AUTHENTICATION & LOGIN SCREEN ---
def login_screen():
    st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>KEA COMPREHENSIVE SCHOOL</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #4B5563;'>School Management & Information System (2027)</h3>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=150)
        
        st.markdown("### Please Login to Continue")
        username = st.text_input("Username").strip()
        password = st.text_input("Password", type="password").strip()
        
        if st.button("Login", type="primary", use_container_width=True):
            if username == "Admin" and password == "janabieli@26!":
                st.session_state.logged_in = True
                st.session_state.username = "Admin"
                st.session_state.role = "DEVELOPER"
                st.session_state.full_name = "System Developer"
                st.success("Welcome back, System Developer!")
                st.rerun()
                
            elif supabase:
                try:
                    res = supabase.table("teachers").select("*").ilike("username", username).execute()
                    if res.data:
                        user = res.data[0]
                        if password == "kea@26" or password == user.get("password"):
                            st.session_state.logged_in = True
                            st.session_state.username = user["username"]
                            st.session_state.role = user["designation"]
                            st.session_state.full_name = user["full_name"]
                            st.success(f"Welcome back, {user['full_name']}!")
                            st.rerun()
                        else:
                            st.error("Incorrect password. Default teacher password is 'kea@26'.")
                    else:
                        st.error("Username not found in teachers directory.")
                except Exception as e:
                    st.error(f"Database error during login: {e}")
            else:
                st.error("Supabase connection not initialized.")

if not st.session_state.logged_in:
    login_screen()
    st.stop()

# --- ROLE & NAVIGATION SETUP ---
role = st.session_state.role
is_admin_role = role in ["HOI", "DHOI", "DEVELOPER"]

st.sidebar.markdown(f"**Logged in as:** {st.session_state.full_name}")
st.sidebar.markdown(f"**Role:** {role}")
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### Navigation")

if is_admin_role:
    nav_options = [
        "Dashboard",
        "Students Registration",
        "Marks Entry",
        "Results Analysis",
        "Teachers Portal",
        "Newsletter"
    ]
else:
    nav_options = [
        "Dashboard",
        "Marks Entry",
        "Newsletter"
    ]

page = st.sidebar.selectbox("Go to", nav_options)

# --- GRADING HELPER FUNCTIONS ---
def get_subject_performance(score, grade_str):
    if score is None or pd.isna(score) or float(score) < 0.0:
        return "-", "-", None
    score = float(score)
    g_lower = grade_str.strip().lower()
    is_junior = any(g in g_lower for g in ["grade 7", "grade 8", "grade 9"])

    if is_junior:
        if score <= 10: return "BE2", "Below Expectation 2", 1
        elif score <= 20: return "BE1", "Below Expectation 1", 2
        elif score <= 30: return "AE2", "Approaching Expectation 2", 3
        elif score <= 40: return "AE1", "Approaching Expectation 1", 4
        elif score <= 57: return "ME2", "Meeting Expectation 2", 5
        elif score <= 74: return "ME1", "Meeting Expectation 1", 6
        elif score <= 89: return "EE2", "Exceeding Expectation 2", 7
        else: return "EE1", "Exceeding Expectation 1", 8
    else:
        if score <= 24: return "BE", "Below Expectation", 1
        elif score <= 49: return "AE", "Approaching Expectation", 2
        elif score <= 74: return "ME", "Meeting Expectation", 3
        else: return "EE", "Exceeding Expectation", 4

def calculate_total_grade(total_points, grade_str):
    g_lower = grade_str.strip().lower()
    is_junior = any(g in g_lower for g in ["grade 7", "grade 8", "grade 9"])
    is_lower = any(g in g_lower for g in ["grade 1", "grade 2", "grade 3"])

    if is_junior:
        if total_points <= 13: return "BE2 (Below Expectation 2)"
        elif total_points <= 22: return "BE1 (Below Expectation 1)"
        elif total_points <= 31: return "AE2 (Approaching Expectation 2)"
        elif total_points <= 40: return "AE1 (Approaching Expectation 1)"
        elif total_points <= 49: return "ME2 (Meeting Expectation 2)"
        elif total_points <= 58: return "ME1 (Meeting Expectation 1)"
        elif total_points <= 67: return "EE2 (Exceeding Expectation 2)"
        else: return "EE1 (Exceeding Expectation 1)"
    elif is_lower:
        if total_points <= 8: return "BE (Below Expectation)"
        elif total_points <= 14: return "AE (Approaching Expectation)"
        elif total_points <= 20: return "ME (Meeting Expectation)"
        else: return "EE (Exceeding Expectation)"
    else:
        if total_points <= 11: return "BE (Below Expectation)"
        elif total_points <= 19: return "AE (Approaching Expectation)"
        elif total_points <= 27: return "ME (Meeting Expectation)"
        else: return "EE (Exceeding Expectation)"

# Precise single-page comments
def generate_teacher_comment(total_points, top_subject, low_subject, grade_str):
    overall_grade = calculate_total_grade(total_points, grade_str)
    return (
        f"Good effort overall ({overall_grade}). Strong in {top_subject}; "
        f"focus more practice on {low_subject}."
    )

def generate_hoi_comment(total_points, grade_str):
    overall_grade = calculate_total_grade(total_points, grade_str)
    return f"Steady term progress ({overall_grade}). Keep up discipline and aim higher."

def generate_dos_comment(total_points, grade_str):
    overall_grade = calculate_total_grade(total_points, grade_str)
    return f"Commendable work ({overall_grade}). Maintain consistent revision and focus."

def encode_photo_upload(uploaded_file):
    if uploaded_file is None:
        return None
    raw_bytes = uploaded_file.getvalue()
    ext = uploaded_file.type.split("/")[-1] if uploaded_file.type and "/" in uploaded_file.type else "png"
    return f"data:image/{ext};base64," + base64.b64encode(raw_bytes).decode()

# --- PAGE 1: DASHBOARD ---
if page == "Dashboard":
    if is_admin_role:
        st.markdown("""
            <style>
            .banner-card {
                background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
                border-radius: 12px;
                padding: 20px;
                color: white;
                border-left: 5px solid #3B82F6;
            }
            .support-card {
                background-color: #EA580C;
                border-radius: 12px;
                padding: 20px;
                color: white;
                text-align: center;
            }
            .metric-card {
                background-color: #1E293B;
                border-radius: 10px;
                padding: 15px;
                border: 1px solid #334155;
                color: white;
                text-align: center;
            }
            .course-card {
                background-color: #1E293B;
                border-radius: 10px;
                padding: 15px;
                color: white;
                border: 1px solid #334155;
                margin-bottom: 10px;
            }
            </style>
        """, unsafe_allow_html=True)

        total_students = 0
        male_students = 0
        female_students = 0
        total_teachers = 0

        if supabase:
            try:
                students_res = supabase.table("students").select("gender", count="exact").execute()
                teachers_res = supabase.table("teachers").select("*", count="exact").execute()
                
                all_st = students_res.data if students_res.data else []
                total_students = students_res.count if students_res.count is not None else len(all_st)
                male_students = sum(1 for s in all_st if str(s.get("gender", "")).strip().lower() == "male")
                female_students = sum(1 for s in all_st if str(s.get("gender", "")).strip().lower() == "female")
                total_teachers = teachers_res.count if teachers_res.count is not None else len(teachers_res.data)
            except Exception as e:
                st.error(f"Error fetching dashboard metrics: {e}")

        b_col1, b_col2 = st.columns([3, 1])
        with b_col1:
            st.markdown(f"""
                <div class="banner-card">
                    <h3>🎓 KEA COMPREHENSIVE SCHOOL — Management Portal (2027)</h3>
                    <p>Welcome back, <b>{st.session_state.full_name}</b> ({st.session_state.role})! Manage student registrations, subject marks entry, results analysis, and school newsletters seamlessly.</p>
                </div>
            """, unsafe_allow_html=True)
        with b_col2:
            st.markdown("""
                <div class="support-card">
                    <h4 style="margin:0;">📞 School Support</h4>
                    <p style="margin:5px 0 0 0; font-size:16px; font-weight:bold;">+254 726 502 706</p>
                    <small>Migori County, Kenya</small>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f"""
                <div class="metric-card">
                    <small style="color:#94A3B8;">TOTAL STUDENTS</small>
                    <h2 style="margin:5px 0; color:#60A5FA;">{total_students}</h2>
                    <small style="color:#34D399;">Active Enrolment (2027)</small>
                </div>
            """, unsafe_allow_html=True)
        with m2:
            st.markdown(f"""
                <div class="metric-card">
                    <small style="color:#94A3B8;">MALE STUDENTS</small>
                    <h2 style="margin:5px 0; color:#38BDF8;">{male_students}</h2>
                    <small style="color:#94A3B8;">Boys Enrolled</small>
                </div>
            """, unsafe_allow_html=True)
        with m3:
            st.markdown(f"""
                <div class="metric-card">
                    <small style="color:#94A3B8;">FEMALE STUDENTS</small>
                    <h2 style="margin:5px 0; color:#F472B6;">{female_students}</h2>
                    <small style="color:#94A3B8;">Girls Enrolled</small>
                </div>
            """, unsafe_allow_html=True)
        with m4:
            st.markdown(f"""
                <div class="metric-card">
                    <small style="color:#94A3B8;">TEACHING STAFF</small>
                    <h2 style="margin:5px 0; color:#A78BFA;">{total_teachers}</h2>
                    <small style="color:#34D399;">Assigned Educators</small>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        c_left, c_right = st.columns([2, 1])
        with c_left:
            st.markdown("### 📚 Comprehensive School Sections")
            cls_col1, cls_col2, cls_col3 = st.columns(3)
            with cls_col1:
                st.markdown("""
                    <div class="course-card">
                        <h4>📘 Lower Primary</h4>
                        <p style="color:#94A3B8; margin:0;">Grade 1 — Grade 3</p>
                        <small style="color:#60A5FA;">6 Learning Areas</small>
                    </div>
                """, unsafe_allow_html=True)
            with cls_col2:
                st.markdown("""
                    <div class="course-card">
                        <h4>📗 Upper Primary</h4>
                        <p style="color:#94A3B8; margin:0;">Grade 4 — Grade 6</p>
                        <small style="color:#60A5FA;">8 Learning Areas</small>
                    </div>
                """, unsafe_allow_html=True)
            with cls_col3:
                st.markdown("""
                    <div class="course-card">
                        <h4>📙 Junior School</h4>
                        <p style="color:#94A3B8; margin:0;">Grade 7 — Grade 9</p>
                        <small style="color:#60A5FA;">9 Learning Areas</small>
                    </div>
                """, unsafe_allow_html=True)
        with c_right:
            st.markdown("### 🗓️ School Calendar")
            st.date_input("Term Calendar View", value=datetime.now())
    else:
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border-radius: 12px; padding: 25px; color: white; border-left: 5px solid #10B981; margin-bottom: 20px;">
                <h2 style="margin:0; color:#34D399;">👋 Welcome, Teacher {st.session_state.full_name}!</h2>
                <p style="margin:5px 0 0 0; color:#94A3B8;">Professional Educator Portal | Teaching & Learning Area Analysis Dashboard (2027)</p>
            </div>
        """, unsafe_allow_html=True)

        user_key = st.session_state.username
        teacher_info = TEACHER_ASSIGNMENTS.get(user_key, {})
        assignments = teacher_info.get("assignments", [])
        class_teacher_grade = teacher_info.get("class_teacher", None)

        if class_teacher_grade:
            st.success(f"📌 **Class Teacher Responsibility:** You are the Class Teacher for **{class_teacher_grade}**.")

        if not assignments:
            st.warning("No specific teaching assignments are currently configured for your account.")
        else:
            st.markdown("### 📋 Your Assigned Learning Areas & Grades")
            cols = st.columns(min(len(assignments), 3))
            for idx, item in enumerate(assignments):
                with cols[idx % len(cols)]:
                    st.markdown(f"""
                        <div style="background-color: #1E293B; border-radius: 10px; padding: 15px; border: 1px solid #334155; color: white; margin-bottom: 10px;">
                            <h4 style="margin:0; color:#60A5FA;">{item['grade']}</h4>
                            <p style="margin:5px 0 0 0; font-size:13px; color:#CBD5E1;"><b>Subjects:</b> {', '.join(item['subjects'])}</p>
                        </div>
                    """, unsafe_allow_html=True)

# --- PAGE 2: STUDENTS REGISTRATION ---
elif page == "Students Registration":
    if not is_admin_role:
        st.error("Access Denied: Only HOI, DHOI, and Senior Teachers can register students.")
        st.stop()
        
    st.header("Students Registration Portal (2027)")
    tab_manual, tab_excel, tab_manage, tab_edit_learner = st.tabs([
        "Manual Registration",
        "Excel Spreadsheet Upload",
        "Manage Registered Students",
        "Edit Learner & Photo"
    ])
    
    with tab_manual:
        with st.form("manual_student_form"):
            grade = st.selectbox("Grade", [f"Grade {i}" for i in range(1, 10)])
            g_lower = grade.strip().lower()
            is_junior = any(g in g_lower for g in ["grade 7", "grade 8", "grade 9"])
            
            adm_no = st.text_input("Admission Number")
            if is_junior:
                assessment_no = st.text_input("Assessment Number (Mandatory for Junior School)")
            else:
                assessment_no = st.text_input("Assessment Number (Optional for Primary)")
                
            name = st.text_input("Full Name")
            gender = st.selectbox("Gender", ["Male", "Female"])
            photo_file = st.file_uploader("Upload Learner Photo (Optional - Free size)", type=["png", "jpg", "jpeg"])
            
            submit_manual = st.form_submit_button("Admit Student")
            
            if submit_manual:
                if adm_no and name and grade and (not is_junior or assessment_no):
                    try:
                        photo_b64 = encode_photo_upload(photo_file) or ""
                        supabase.table("students").upsert({
                            "adm_no": adm_no.strip(),
                            "assessment_no": assessment_no.strip(),
                            "name": name.strip(),
                            "gender": gender,
                            "grade": grade,
                            "photo": photo_b64
                        }, on_conflict="adm_no").execute()
                        st.success(f"Successfully admitted {name} to {grade}!")
                    except Exception as e:
                        st.error(f"Error saving student: {e}")
                else:
                    st.warning("Please fill in all mandatory fields (Assessment Number is required for Junior School).")

    with tab_excel:
        st.markdown("Upload an Excel spreadsheet with columns: **Adm No**, **Assessment No**, **Name**, **Gender**, **Grade**")
        uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx", "xls"])
        selected_grade_bulk = st.selectbox("Target Grade for Bulk Upload", [f"Grade {i}" for i in range(1, 10)])
        
        if uploaded_file and st.button("Process & Admit Students"):
            try:
                df = pd.read_excel(uploaded_file)
                df.columns = [str(c).strip().lower() for c in df.columns]
                
                adm_col = next((c for c in df.columns if "adm" in c), None)
                name_col = next((c for c in df.columns if "name" in c), None)
                assess_col = next((c for c in df.columns if "assess" in c), None)
                gender_col = next((c for c in df.columns if "gender" in c), None)
                grade_col = next((c for c in df.columns if "grade" in c), None)
                
                if not adm_col or not name_col:
                    st.error("Excel file must contain at least 'Adm No' and 'Name' columns.")
                    st.stop()

                records = []
                for _, row in df.iterrows():
                    adm_val = str(row[adm_col]).strip() if pd.notna(row[adm_col]) else None
                    name_val = str(row[name_col]).strip() if pd.notna(row[name_col]) else None
                    assess_val = str(row[assess_col]).strip() if assess_col and pd.notna(row[assess_col]) else ""
                    gender_val = str(row[gender_col]).strip() if gender_col and pd.notna(row[gender_col]) else "Male"
                    grade_val = str(row[grade_col]).strip() if grade_col and pd.notna(row[grade_col]) else selected_grade_bulk
                    
                    if adm_val and adm_val.lower() != "nan" and name_val and name_val.lower() != "nan":
                        records.append({
                            "adm_no": adm_val,
                            "assessment_no": assess_val,
                            "name": name_val,
                            "gender": gender_val,
                            "grade": grade_val
                        })

                if records:
                    supabase.table("students").upsert(records, on_conflict="adm_no").execute()
                    st.success(f"Successfully uploaded and admitted {len(records)} students!")
            except Exception as e:
                st.error(f"Error processing excel file: {e}")

    with tab_manage:
        st.markdown("### View, Edit & Deregister Students / Grades")
        filter_grade = st.selectbox("Filter by Grade", ["All Grades"] + [f"Grade {i}" for i in range(1, 10)])
        
        try:
            query = supabase.table("students").select("*")
            if filter_grade != "All Grades":
                query = query.eq("grade", filter_grade)
            
            res = query.execute()
            students_list = res.data if res.data else []
            
            if not students_list:
                st.info("No students found.")
            else:
                df_students = pd.DataFrame(students_list)
                cols = [c for c in ["adm_no", "assessment_no", "name", "gender", "grade"] if c in df_students.columns]
                df_students = df_students[cols]
                
                st.markdown("#### Registered Students Roster")
                edited_students_df = st.data_editor(
                    df_students,
                    column_config={
                        "adm_no": st.column_config.TextColumn("Adm No", disabled=True),
                        "assessment_no": st.column_config.TextColumn("Assessment No"),
                        "name": st.column_config.TextColumn("Full Name"),
                        "gender": st.column_config.SelectboxColumn("Gender", options=["Male", "Female"]),
                        "grade": st.column_config.SelectboxColumn("Grade", options=[f"Grade {i}" for i in range(1, 10)])
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="students_roster_editor"
                )
                
                if st.button("💾 Save Changes to Database", type="primary"):
                    try:
                        records_to_update = edited_students_df.to_dict(orient="records")
                        supabase.table("students").upsert(records_to_update, on_conflict="adm_no").execute()
                        st.success("Successfully updated student records!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error updating records: {e}")

                st.markdown("---")
                st.markdown("#### Deregister Options")
                d_mode = st.radio("Deregistration Mode", ["Individual Learner", "Whole Grade"], horizontal=True)
                
                if d_mode == "Individual Learner":
                    col_del1, col_del2 = st.columns([3, 1])
                    with col_del1:
                        student_to_delete = st.selectbox(
                            "Select Student to Deregister", 
                            students_list, 
                            format_func=lambda s: f"Adm No: {s['adm_no']} — {s['name']} ({s['grade']})"
                        )
                    with col_del2:
                        st.write("")
                        st.write("")
                        confirm_delete = st.button("🚨 Deregister Student", type="secondary")
                        
                    if confirm_delete and student_to_delete:
                        adm_del = student_to_delete["adm_no"]
                        try:
                            supabase.table("marks").delete().eq("adm_no", adm_del).execute()
                            supabase.table("students").delete().eq("adm_no", adm_del).execute()
                            st.success("Successfully deregistered student.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error deregistering student: {e}")
                else:
                    grade_to_clear = st.selectbox("Select Whole Grade to Deregister", [f"Grade {i}" for i in range(1, 10)])
                    confirm_grade_del = st.button("⚠️ Deregister All Learners in Grade", type="secondary")
                    if confirm_grade_del:
                        try:
                            grade_students = [s for s in students_list if str(s.get("grade", "")).strip().lower() == grade_to_clear.lower()]
                            for gs in grade_students:
                                supabase.table("marks").delete().eq("adm_no", gs["adm_no"]).execute()
                                supabase.table("students").delete().eq("adm_no", gs["adm_no"]).execute()
                            st.success(f"Successfully deregistered all learners in {grade_to_clear}.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error clearing grade: {e}")
        except Exception as e:
            st.error(f"Error loading students: {e}")

    with tab_edit_learner:
        st.markdown("### ✏️ Edit Learner Details & Photo")
        try:
            all_st_res = supabase.table("students").select("*").execute()
            all_st_list = sorted(all_st_res.data if all_st_res.data else [], key=lambda x: str(x.get("adm_no", "")))
            
            if not all_st_list:
                st.info("No registered learners available to edit.")
            else:
                selected_learner = st.selectbox(
                    "Choose Registered Learner",
                    all_st_list,
                    format_func=lambda s: f"{s.get('adm_no')} — {s.get('name')} ({s.get('grade')})"
                )
                
                if selected_learner:
                    current_photo = selected_learner.get("photo", "")
                    col_info, col_pic = st.columns([2, 1])
                    with col_pic:
                        st.markdown("**Current Photo:**")
                        if current_photo:
                            st.markdown(f"<img src='{current_photo}' style='max-width:140px; max-height:160px; border:1px solid #374151; border-radius:6px;' />", unsafe_allow_html=True)
                        else:
                            st.caption("No photo currently uploaded.")
                    
                    with col_info:
                        with st.form(key=f"edit_learner_form_{selected_learner.get('adm_no')}"):
                            new_adm_no = st.text_input("Admission Number", value=str(selected_learner.get("adm_no", "")))
                            new_assess_no = st.text_input("Assessment Number", value=str(selected_learner.get("assessment_no") or ""))
                            new_name = st.text_input("Full Name", value=str(selected_learner.get("name", "")))
                            
                            curr_gender = str(selected_learner.get("gender", "Male")).capitalize()
                            gender_idx = 0 if curr_gender != "Female" else 1
                            new_gender = st.selectbox("Gender", ["Male", "Female"], index=gender_idx)
                            
                            grade_opts = [f"Grade {i}" for i in range(1, 10)]
                            curr_grade = str(selected_learner.get("grade", "Grade 1"))
                            grade_idx = grade_opts.index(curr_grade) if curr_grade in grade_opts else 0
                            new_grade = st.selectbox("Grade", grade_opts, index=grade_idx)
                            
                            new_photo_file = st.file_uploader("Upload / Replace Learner Photo (Free size)", type=["png", "jpg", "jpeg"])
                            
                            if st.form_submit_button("💾 Save Learner Changes", type="primary"):
                                if not new_adm_no.strip() or not new_name.strip():
                                    st.warning("Admission Number and Full Name cannot be empty.")
                                else:
                                    try:
                                        final_photo = encode_photo_upload(new_photo_file)
                                        if final_photo is None:
                                            final_photo = current_photo or ""
                                        
                                        old_adm = str(selected_learner.get("adm_no"))
                                        updated_payload = {
                                            "adm_no": new_adm_no.strip(),
                                            "assessment_no": new_assess_no.strip(),
                                            "name": new_name.strip(),
                                            "gender": new_gender,
                                            "grade": new_grade,
                                            "photo": final_photo
                                        }
                                        
                                        if new_adm_no.strip() != old_adm:
                                            supabase.table("students").insert(updated_payload).execute()
                                            supabase.table("marks").update({"adm_no": new_adm_no.strip(), "grade": new_grade}).eq("adm_no", old_adm).execute()
                                            supabase.table("students").delete().eq("adm_no", old_adm).execute()
                                        else:
                                            supabase.table("students").upsert(updated_payload, on_conflict="adm_no").execute()
                                            supabase.table("marks").update({"grade": new_grade}).eq("adm_no", old_adm).execute()
                                            
                                        st.success(f"Updated details for {new_name.strip()}!")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Error saving learner updates: {e}")
        except Exception as e:
            st.error(f"Error loading learners: {e}")

# --- PAGE 3: MARKS ENTRY ---
elif page == "Marks Entry":
    st.header("Subject-Specific Marks Entry Portal (2027)")
    
    user_key = st.session_state.username
    teacher_data = TEACHER_ASSIGNMENTS.get(user_key)
    
    if is_admin_role or not teacher_data:
        allowed_grades = [f"Grade {i}" for i in range(1, 10)]
        st.info("Log-in Mode: Administrator Mode (Full Access)")
    else:
        teacher_grade_map = {item["grade"]: item["subjects"] for item in teacher_data["assignments"]}
        allowed_grades = list(teacher_grade_map.keys())

    if not allowed_grades:
        st.warning("No grades currently assigned to your profile.")
        st.stop()

    col_g, col_s = st.columns(2)
    with col_g:
        target_grade = st.selectbox("1. Select Grade", allowed_grades)
    
    grade_learning_areas = get_learning_areas_for_grade(target_grade)
    assigned_subs_for_grade = teacher_grade_map.get(target_grade, grade_learning_areas) if (not is_admin_role and teacher_data) else grade_learning_areas
    
    with col_s:
        target_subject = st.selectbox("2. Select Learning Area", assigned_subs_for_grade)

    st.markdown("---")
    st.subheader(f"Marking Sheet: {target_grade} — {target_subject}")
    
    grade_num = target_grade.replace("Grade", "").strip()
    try:
        students_res = supabase.table("students").select("adm_no, name, grade").execute()
        all_students = students_res.data if students_res.data else []
        students = [s for s in all_students if str(s.get("grade", "")).strip().lower() in [target_grade.lower(), grade_num, f"grade {grade_num}"]]
        students = sorted(students, key=lambda x: str(x.get("adm_no", "")))
    except Exception as e:
        students = []
        st.error(f"Error fetching students: {e}")

    if not students:
        st.warning(f"No students found under '{target_grade}'.")
    else:
        col_key = get_db_col_name(target_subject)
        existing_marks_map = {}
        existing_full_rows = {}
        try:
            m_res = supabase.table("marks").select("*").eq("grade", target_grade).execute()
            if m_res.data:
                for row in m_res.data:
                    adm = row["adm_no"]
                    existing_full_rows[adm] = row
                    val = row.get(col_key)
                    if val is not None and not pd.isna(val):
                        existing_marks_map[adm] = float(val)
        except Exception:
            pass

        unrecorded_students = [s for s in students if s["adm_no"] not in existing_marks_map]
        recorded_students = [s for s in students if s["adm_no"] in existing_marks_map]

        tab_manual_marks, tab_excel_marks = st.tabs(["Manual Sheet Input", "Upload Excel Marks"])

        with tab_manual_marks:
            if unrecorded_students:
                with st.form(key=f"unrecorded_form_{target_grade}_{col_key}"):
                    editor_rows = [{"Adm No": s["adm_no"], "Student Name": s["name"], "Score (%)": None, "Save & Submit": False} for s in unrecorded_students]
                    edited_df = st.data_editor(
                        pd.DataFrame(editor_rows),
                        column_config={
                            "Adm No": st.column_config.TextColumn("Adm No", disabled=True),
                            "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                            "Score (%)": st.column_config.NumberColumn("Score (%)", min_value=0.0, max_value=100.0, step=1.0, format="%.0f"),
                            "Save & Submit": st.column_config.CheckboxColumn("Save & Submit", default=False)
                        },
                        use_container_width=True, hide_index=True
                    )
                    if st.form_submit_button("💾 Save and Submit Marks", type="primary"):
                        records_to_upsert = []
                        grade_learning_list = get_learning_areas_for_grade(target_grade)
                        for _, row in edited_df.iterrows():
                            adm = str(row["Adm No"])
                            raw_score = row["Score (%)"]
                            if row["Save & Submit"] and raw_score is not None and not pd.isna(raw_score):
                                existing = existing_full_rows.get(adm, {})
                                rec = dict(existing) if existing else {"adm_no": adm, "grade": target_grade}
                                rec[col_key] = float(raw_score)
                                
                                tot_marks = 0.0
                                tot_pts = 0
                                for sub in grade_learning_list:
                                    s_key = get_db_col_name(sub)
                                    val = rec.get(s_key)
                                    if val is not None and not pd.isna(val) and float(val) >= 0:
                                        s_val = float(val)
                                        tot_marks += s_val
                                        _, _, pts = get_subject_performance(s_val, target_grade)
                                        if pts: tot_pts += pts
                                rec["total_marks"] = tot_marks
                                rec["total_points"] = tot_pts
                                records_to_upsert.append(rec)
                        if records_to_upsert:
                            supabase.table("marks").upsert(records_to_upsert, on_conflict="adm_no").execute()
                            st.success("Successfully saved marks!")
                            st.rerun()

            if recorded_students:
                st.markdown("### ✅ Recorded Students Table")
                recorded_rows = [{"Adm No": s["adm_no"], "Student Name": s["name"], "Score (%)": float(existing_marks_map.get(s["adm_no"])), "Edit & Save": False} for s in recorded_students]
                with st.form(key=f"recorded_edit_form_{target_grade}_{col_key}"):
                    edited_recorded_df = st.data_editor(
                        pd.DataFrame(recorded_rows),
                        column_config={
                            "Adm No": st.column_config.TextColumn("Adm No", disabled=True),
                            "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                            "Score (%)": st.column_config.NumberColumn("Score (%)", min_value=0.0, max_value=100.0, step=1.0, format="%.0f"),
                            "Edit & Save": st.column_config.CheckboxColumn("Edit & Save", default=False)
                        },
                        use_container_width=True, hide_index=True
                    )
                    if st.form_submit_button("💾 Save Edits", type="primary"):
                        records_to_update = []
                        grade_learning_list = get_learning_areas_for_grade(target_grade)
                        for _, row in edited_recorded_df.iterrows():
                            if row["Edit & Save"]:
                                adm = str(row["Adm No"])
                                existing = existing_full_rows.get(adm, {})
                                rec = dict(existing) if existing else {"adm_no": adm, "grade": target_grade}
                                rec[col_key] = float(row["Score (%)"])
                                tot_marks, tot_pts = 0.0, 0
                                for sub in grade_learning_list:
                                    s_key = get_db_col_name(sub)
                                    val = rec.get(s_key)
                                    if val is not None and not pd.isna(val) and float(val) >= 0:
                                        s_val = float(val)
                                        tot_marks += s_val
                                        _, _, pts = get_subject_performance(s_val, target_grade)
                                        if pts: tot_pts += pts
                                rec["total_marks"] = tot_marks
                                rec["total_points"] = tot_pts
                                records_to_update.append(rec)
                        if records_to_update:
                            supabase.table("marks").upsert(records_to_update, on_conflict="adm_no").execute()
                            st.success("Successfully updated marks!")
                            st.rerun()

# --- PAGE 4: RESULTS ANALYSIS & REPORT FORMS HUB ---
elif page == "Results Analysis":
    st.header("Results Analysis Hub & Report Forms (2027)")
    analysis_grade = st.selectbox("Select Grade for Analysis", [f"Grade {i}" for i in range(1, 10)], key="analysis_grade_select")
    current_grade_learning_areas = get_learning_areas_for_grade(analysis_grade)
    
    tab_overview, tab_ranking, tab_top_subjects, tab_reports = st.tabs([
        "Tab 1: General Performance Overview",
        "Tab 2: Whole Class Ranking",
        "Tab 3: Best Three per Learning Area",
        "Tab 4: Assessment Report Forms Hub"
    ])
    
    grade_num = analysis_grade.replace("Grade", "").strip()
    try:
        m_res = supabase.table("marks").select("*").execute()
        all_marks = m_res.data if m_res.data else []
        marks_data = [m for m in all_marks if str(m.get("grade", "")).strip().lower() in [analysis_grade.lower(), grade_num, f"grade {grade_num}"]]
        
        s_res = supabase.table("students").select("*").execute()
        students_dict = {str(s.get("adm_no", "")).strip(): s for s in s_res.data} if s_res.data else {}
    except Exception:
        marks_data, students_dict = [], {}

    active_ranked_students = []
    for m in marks_data:
        adm_str = str(m.get("adm_no", "")).strip()
        student_info = students_dict.get(adm_str, {})
        s_name = student_info.get("name", f"Student {adm_str}")
        
        total_m, total_p, valid_cnt = 0.0, 0, 0
        for sub in current_grade_learning_areas:
            k = get_db_col_name(sub)
            val = m.get(k)
            if val is not None and not pd.isna(val) and float(val) >= 0.0:
                s_val = float(val)
                total_m += s_val
                _, _, pts = get_subject_performance(s_val, analysis_grade)
                if pts: total_p += pts
                valid_cnt += 1
        
        if valid_cnt > 0:
            mean_pts = round(total_p / len(current_grade_learning_areas), 1)
            active_ranked_students.append({
                "adm_no": adm_str,
                "assessment_no": student_info.get("assessment_no", ""),
                "name": s_name,
                "gender": student_info.get("gender", ""),
                "photo": student_info.get("photo", ""),
                "total_marks": total_m,
                "total_points": total_p,
                "mean_points": mean_pts,
                "marks_row": m
            })

    active_ranked_students = sorted(active_ranked_students, key=lambda x: (x["total_points"], x["total_marks"]), reverse=True)
    for idx, item in enumerate(active_ranked_students):
        item["rank"] = idx + 1

    # --- TAB 1: GENERAL PERFORMANCE OVERVIEW + LEARNING AREA RANKING ---
    with tab_overview:
        if active_ranked_students:
            st.markdown("### 🏆 Top Ten Students")
            top_10 = [{
                "Rank": i["rank"], "Adm No": i["adm_no"], "Student Name": i["name"],
                "Total Marks": i["total_marks"], "Total Points": i["total_points"],
                "Mean Points": i["mean_points"], "Grade": calculate_total_grade(i["total_points"], analysis_grade)
            } for i in active_ranked_students[:10]]
            st.dataframe(pd.DataFrame(top_10), use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("### 📊 Learning Area Ranking (By Mean Score)")
            area_stats = []
            for sub in current_grade_learning_areas:
                col_k = get_db_col_name(sub)
                scores = [
                    float(item["marks_row"].get(col_k))
                    for item in active_ranked_students
                    if item["marks_row"].get(col_k) is not None
                    and not pd.isna(item["marks_row"].get(col_k))
                    and float(item["marks_row"].get(col_k)) >= 0.0
                ]
                if scores:
                    mean_s = sum(scores) / len(scores)
                    mean_p = sum(get_subject_performance(s, analysis_grade)[2] or 0 for s in scores) / len(scores)
                else:
                    mean_s, mean_p = 0.0, 0.0
                area_stats.append({
                    "Learning Area": sub,
                    "Entries": len(scores),
                    "Mean Score (%)": round(mean_s, 2),
                    "Mean Points": round(mean_p, 2)
                })
            area_stats.sort(key=lambda x: (x["Mean Score (%)"], x["Mean Points"]), reverse=True)
            for idx_a, a_row in enumerate(area_stats, 1):
                a_row["Area Rank"] = idx_a
            area_cols = ["Area Rank", "Learning Area", "Entries", "Mean Score (%)", "Mean Points"]
            st.dataframe(pd.DataFrame(area_stats)[area_cols], use_container_width=True, hide_index=True)
        else:
            st.info(f"No active student records with submitted marks found in {analysis_grade}.")

    # --- TAB 2: WHOLE CLASS RANKING ---
    with tab_ranking:
        if active_ranked_students:
            rows = [{
                "Rank": i["rank"], "Adm No": i["adm_no"], "Student Name": i["name"],
                "Total Marks": i["total_marks"], "Total Points": i["total_points"],
                "Mean Points": i["mean_points"], "Grade": calculate_total_grade(i["total_points"], analysis_grade)
            } for i in active_ranked_students]
            df_res = pd.DataFrame(rows)
            st.dataframe(df_res, use_container_width=True, hide_index=True)
            st.download_button("📥 Download Roll (.csv)", df_res.to_csv(index=False).encode('utf-8'), f"{analysis_grade}_Roll.csv", "text/csv")

    # --- TAB 3: BEST THREE PER LEARNING AREA + SINGLE EXCEL EXPORT ---
    with tab_top_subjects:
        st.markdown(f"### 🥇 Best Three Students Per Learning Area — {analysis_grade}")
        all_best_rows = []
        excel_buf = io.BytesIO()
        
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            for sub in current_grade_learning_areas:
                col_k = get_db_col_name(sub)
                sub_studs = []
                for item in active_ranked_students:
                    val = item["marks_row"].get(col_k)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        code, _, pts = get_subject_performance(float(val), analysis_grade)
                        sub_studs.append({
                            "Adm No": item["adm_no"],
                            "Student Name": item["name"],
                            "Score (%)": float(val),
                            "Perf. Level": code,
                            "Points": pts
                        })
                sub_studs = sorted(sub_studs, key=lambda x: x["Score (%)"], reverse=True)[:3]
                
                for pos_i, s_entry in enumerate(sub_studs, 1):
                    all_best_rows.append({
                        "Learning Area": sub,
                        "Position": f"Position {pos_i}",
                        "Adm No": s_entry["Adm No"],
                        "Student Name": s_entry["Student Name"],
                        "Score (%)": s_entry["Score (%)"],
                        "Perf. Level": s_entry["Perf. Level"],
                        "Points": s_entry["Points"]
                    })

                sheet_name = sub[:31].replace("/", "-").replace("\\", "-")
                df_sub_top = pd.DataFrame(sub_studs) if sub_studs else pd.DataFrame(columns=["Adm No", "Student Name", "Score (%)", "Perf. Level", "Points"])
                df_sub_top.to_excel(writer, sheet_name=sheet_name, index=False)

            df_combined_best = pd.DataFrame(all_best_rows) if all_best_rows else pd.DataFrame(columns=["Learning Area", "Position", "Adm No", "Student Name", "Score (%)", "Perf. Level", "Points"])
            df_combined_best.to_excel(writer, sheet_name="All_Best_Three", index=False)

        excel_buf.seek(0)
        st.download_button(
            label="📊 Download Best Per Learning Area (Single Excel .xlsx)",
            data=excel_buf,
            file_name=f"{analysis_grade.replace(' ', '_')}_Best_Per_Learning_Area.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        for sub in current_grade_learning_areas:
            sub_rows = [r for r in all_best_rows if r["Learning Area"] == sub]
            with st.expander(f"📚 {sub}", expanded=True):
                if sub_rows:
                    display_rows = [{"Position": r["Position"], "Adm No": r["Adm No"], "Student Name": r["Student Name"], "Score (%)": r["Score (%)"], "Perf. Level": r["Perf. Level"], "Points": r["Points"]} for r in sub_rows]
                    st.dataframe(pd.DataFrame(display_rows), use_container_width=True, hide_index=True)
                else:
                    st.caption("No scores recorded.")

    # --- TAB 4: ASSESSMENT REPORT FORMS HUB ---
    with tab_reports:
        st.markdown("### Assessment Report Form Generator (2027)")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            term_val = st.selectbox("Term", ["Term 1", "Term 2", "Term 3"], key="rep_term")
            opening_date = st.text_input("Opening Date", "11/01/2027", key="rep_open")
        with col_s2:
            closing_date = st.text_input("Closing Date", "02/04/2027", key="rep_close")
            stamp_upload = st.file_uploader("Upload Stamp Image (stamp.png)", type=["png", "jpg"], key="rep_stamp")

        st.markdown("---")
        if not active_ranked_students:
            st.info(f"No students with active exam scores found in {analysis_grade}.")
        else:
            g_lower = analysis_grade.strip().lower()
            is_junior_rep = any(g in g_lower for g in ["grade 7", "grade 8", "grade 9"])
            template_filename = (
                "report3.docx" if any(g in g_lower for g in ["grade 1", "grade 2", "grade 3"])
                else ("report2.docx" if any(g in g_lower for g in ["grade 4", "grade 5", "grade 6"]) else "report1.docx")
            )

            def build_report_data(student_obj):
                adm_str = student_obj["adm_no"]
                assess_str = student_obj.get("assessment_no", "")
                student_name = student_obj["name"]
                gender_str = student_obj.get("gender", "")
                s_row = student_obj["marks_row"]
                student_rank = student_obj["rank"]
                total_marks = student_obj["total_marks"]
                total_points = student_obj["total_points"]
                mean_points = student_obj["mean_points"]

                learning_list = get_learning_areas_for_grade(analysis_grade)
                top_sub, low_sub = learning_list[0], learning_list[0]
                best_s, worst_s = -1.0, 101.0

                subject_details = []
                for idx_s, sub in enumerate(learning_list, 1):
                    col_key = get_db_col_name(sub)
                    val = s_row.get(col_key)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        s_num = float(val)
                        if s_num > best_s: best_s = s_num; top_sub = sub
                        if s_num < worst_s: worst_s = s_num; low_sub = sub
                        code, _, pts = get_subject_performance(s_num, analysis_grade)

                        sub_scores_list = sorted([
                            (st_item["adm_no"], float(st_item["marks_row"].get(col_key)))
                            for st_item in active_ranked_students
                            if st_item["marks_row"].get(col_key) is not None
                            and not pd.isna(st_item["marks_row"].get(col_key))
                            and float(st_item["marks_row"].get(col_key)) >= 0.0
                        ], key=lambda x: x[1], reverse=True)
                        sub_rank = next((i + 1 for i, (a, _) in enumerate(sub_scores_list) if a == adm_str), 1)
                        pos_str = f"{sub_rank}/{len(sub_scores_list)}"
                        subject_details.append({
                            "no": str(idx_s),
                            "name": sub.upper(),
                            "marks": f"{s_num:.0f}",
                            "level": code,
                            "points": str(pts),
                            "position": pos_str,
                            "teacher": "TBA"
                        })
                    else:
                        subject_details.append({
                            "no": str(idx_s),
                            "name": sub.upper(),
                            "marks": "-",
                            "level": "-",
                            "points": "-",
                            "position": "-",
                            "teacher": "TBA"
                        })

                overall_g = calculate_total_grade(total_points, analysis_grade)
                teacher_comment = generate_teacher_comment(total_points, top_sub, low_sub, analysis_grade)
                hoi_comment = generate_hoi_comment(total_points, analysis_grade)
                dos_comment = generate_dos_comment(total_points, analysis_grade) if is_junior_rep else ""

                return {
                    "adm_str": adm_str,
                    "assess_str": assess_str,
                    "student_name": student_name,
                    "gender_str": gender_str,
                    "student_rank": student_rank,
                    "total_marks": total_marks,
                    "total_points": total_points,
                    "mean_points": mean_points,
                    "overall_g": overall_g,
                    "subject_details": subject_details,
                    "teacher_comment": teacher_comment,
                    "hoi_comment": hoi_comment,
                    "dos_comment": dos_comment
                }

            def generate_student_docx_object(student_obj):
                d = build_report_data(student_obj)
                # Build a clean single-page layout directly so columns never shift or overlap
                doc = Document()
                for section in doc.sections:
                    section.top_margin = Inches(0.25)
                    section.bottom_margin = Inches(0.25)
                    section.left_margin = Inches(0.35)
                    section.right_margin = Inches(0.35)

                # Header table: Logo | School Title | Photo
                hdr_tbl = doc.add_table(rows=1, cols=3)
                hdr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                c_logo, c_title, c_photo = hdr_tbl.rows[0].cells
                c_logo.width = Inches(0.9)
                c_title.width = Inches(5.8)
                c_photo.width = Inches(1.0)

                if os.path.exists("logo.png"):
                    p_l = c_logo.paragraphs[0]
                    p_l.paragraph_format.space_before = Pt(0)
                    p_l.paragraph_format.space_after = Pt(0)
                    p_l.add_run().add_picture("logo.png", width=Inches(0.75))

                p_t = c_title.paragraphs[0]
                p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_t.paragraph_format.space_before = Pt(0)
                p_t.paragraph_format.space_after = Pt(1)
                r1 = p_t.add_run("KEA COMPREHENSIVE SCHOOL\n")
                r1.bold = True
                r1.font.size = Pt(13)
                r2 = p_t.add_run("P.O. BOX 557-40400 SUNA MIGORI | Tel: 0726 502706\n")
                r2.font.size = Pt(8)
                r3 = p_t.add_run(f"LEARNER ASSESSMENT REPORT — {analysis_grade.upper()} ({term_val.upper()}, 2027)")
                r3.bold = True
                r3.font.size = Pt(9.5)

                p_ph = c_photo.paragraphs[0]
                p_ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_ph.paragraph_format.space_before = Pt(0)
                p_ph.paragraph_format.space_after = Pt(0)
                photo_str = student_obj.get("photo", "")
                photo_added = False
                if photo_str and "," in photo_str:
                    try:
                        img_bytes = io.BytesIO(base64.b64decode(photo_str.split(",", 1)[1]))
                        p_ph.add_run().add_picture(img_bytes, width=Inches(0.75))
                        photo_added = True
                    except Exception:
                        photo_added = False
                if not photo_added:
                    rp = p_ph.add_run("[PHOTO]")
                    rp.font.size = Pt(8)

                # Student bio table
                bio_tbl = doc.add_table(rows=2, cols=3)
                bio_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                bio_rows = [
                    [f"NAME: {d['student_name']}", f"GRADE: {analysis_grade}", f"GENDER: {d['gender_str']}"],
                    [f"ADM NO: {d['adm_str']}", f"ASSESS. NO: {d['assess_str']}", f"RANK: #{d['student_rank']} / {len(active_ranked_students)}"]
                ]
                for r_idx, b_row in enumerate(bio_rows):
                    for c_idx, b_val in enumerate(b_row):
                        cell = bio_tbl.rows[r_idx].cells[c_idx]
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(1)
                        p.paragraph_format.space_after = Pt(1)
                        run = p.add_run(b_val)
                        run.font.size = Pt(8.5)
                        run.bold = True

                # Subject marks table (exact 6 columns)
                subj_tbl = doc.add_table(rows=len(d["subject_details"]) + 2, cols=6)
                subj_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                headers = ["No.", "Learning Area", "Marks (%)", "Perf. Level", "Points", "Position"]
                for c_idx, h_text in enumerate(headers):
                    p = subj_tbl.rows[0].cells[c_idx].paragraphs[0]
                    p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.space_after = Pt(1)
                    run = p.add_run(h_text)
                    run.bold = True
                    run.font.size = Pt(8.5)

                for idx_sd, sd in enumerate(d["subject_details"], 1):
                    vals = [sd["no"], sd["name"], sd["marks"], sd["level"], sd["points"], sd["position"]]
                    for c_idx, v_text in enumerate(vals):
                        p = subj_tbl.rows[idx_sd].cells[c_idx].paragraphs[0]
                        p.paragraph_format.space_before = Pt(1)
                        p.paragraph_format.space_after = Pt(1)
                        run = p.add_run(v_text)
                        run.font.size = Pt(8.5)
                        if c_idx == 1:
                            run.bold = True

                # Totals summary row
                tot_row_cells = subj_tbl.rows[-1].cells
                tot_vals = [
                    "-",
                    f"OVERALL: {d['overall_g']}",
                    f"TOT: {d['total_marks']:.0f}",
                    "-",
                    f"PTS: {d['total_points']}",
                    f"MEAN: {d['mean_points']}"
                ]
                for c_idx, tv in enumerate(tot_vals):
                    p = tot_row_cells[c_idx].paragraphs[0]
                    p.paragraph_format.space_before = Pt(1.5)
                    p.paragraph_format.space_after = Pt(1.5)
                    run = p.add_run(tv)
                    run.bold = True
                    run.font.size = Pt(8.5)

                # Concise Comments & Term Dates
                p_comm = doc.add_paragraph()
                p_comm.paragraph_format.space_before = Pt(3)
                p_comm.paragraph_format.space_after = Pt(2)
                c_lines = [f"Class Teacher: {d['teacher_comment']}"]
                if is_junior_rep:
                    c_lines.append(f"Director of Studies: {d['dos_comment']}")
                c_lines.append(f"Head of Institution: {d['hoi_comment']}")
                c_lines.append(f"Term Closes: {closing_date}    |    Next Term Opens: {opening_date}")
                run_c = p_comm.add_run("\n".join(c_lines))
                run_c.font.size = Pt(8.2)

                # Sign & Stamp row
                sig_tbl = doc.add_table(rows=1, cols=2)
                sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                p_s1 = sig_tbl.rows[0].cells[0].paragraphs[0]
                p_s1.paragraph_format.space_before = Pt(2)
                p_s1.paragraph_format.space_after = Pt(0)
                r_s1 = p_s1.add_run("Class Teacher Sign: ____________________")
                r_s1.font.size = Pt(8.5)

                p_s2 = sig_tbl.rows[0].cells[1].paragraphs[0]
                p_s2.paragraph_format.space_before = Pt(2)
                p_s2.paragraph_format.space_after = Pt(0)
                r_s2 = p_s2.add_run("HOI Sign & Stamp: ")
                r_s2.font.size = Pt(8.5)

                if stamp_upload is not None:
                    p_s2.add_run().add_picture(io.BytesIO(stamp_upload.getvalue()), width=Inches(0.8))
                elif os.path.exists("stamp.png"):
                    p_s2.add_run().add_picture("stamp.png", width=Inches(0.8))
                else:
                    r_s2b = p_s2.add_run("____________________")
                    r_s2b.font.size = Pt(8.5)

                return doc

            st.subheader("1. Live Student Report Form Preview & Single Download")
            selected_student_obj = st.selectbox(
                "Select Student for Live Report Form Preview",
                active_ranked_students,
                format_func=lambda x: f"Rank #{x['rank']} | Adm No: {x['adm_no']} | Name: {x['name']}"
            )

            if selected_student_obj:
                d_prev = build_report_data(selected_student_obj)
                s_photo = selected_student_obj.get("photo", "")

                rows_html = ""
                for sd in d_prev["subject_details"]:
                    rows_html += f"""
                        <tr>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">{sd['no']}</td>
                            <td style="border: 1px solid #374151; padding: 5px; font-weight: bold;">{sd['name']}</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">{sd['marks']}</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">{sd['level']}</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">{sd['points']}</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">{sd['position']}</td>
                        </tr>
                    """

                dos_section = (
                    f"<p style='margin: 6px 0 2px 0;'><strong>DIRECTOR OF STUDIES:</strong> {d_prev['dos_comment']}</p>"
                    if is_junior_rep else ""
                )

                logo_b64 = ""
                if os.path.exists("logo.png"):
                    with open("logo.png", "rb") as f:
                        logo_b64 = base64.b64encode(f.read()).decode()

                stamp_b64 = ""
                if stamp_upload:
                    stamp_b64 = base64.b64encode(stamp_upload.getvalue()).decode()
                elif os.path.exists("stamp.png"):
                    with open("stamp.png", "rb") as f:
                        stamp_b64 = base64.b64encode(f.read()).decode()

                photo_html = (
                    f"<div style='width: 75px; height: 90px; border: 1px solid #374151; display: flex; align-items: center; justify-content: center; overflow: hidden;'><img src='{s_photo}' style='max-width: 100%; max-height: 100%; object-fit: cover;' /></div>"
                    if s_photo else
                    "<div style='width: 75px; height: 90px; border: 1px dashed #9CA3AF; display: flex; align-items: center; justify-content: center; font-size: 10px; color: #6B7280;'>PHOTO</div>"
                )

                html_preview = f"""
                <div style="border: 2px solid #1E3A8A; padding: 20px; border-radius: 8px; background-color: #FFFFFF; color: #111827; font-family: Arial, sans-serif;">
                    <div style="display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; margin-bottom: 10px;">
                        <div>{f'<img src="data:image/png;base64,{logo_b64}" style="height: 65px;" />' if logo_b64 else ''}</div>
                        <div style="text-align: center; flex-grow: 1;">
                            <h2 style="margin: 0; color: #1E3A8A; font-size: 18px;">KEA COMPREHENSIVE SCHOOL</h2>
                            <p style="margin: 2px 0; font-size: 11px;">P.O. BOX 557-40400 SUNA MIGORI | Tel: 0726 502706</p>
                            <h4 style="margin: 3px 0; font-size: 13px;">LEARNER ASSESSMENT REPORT — {analysis_grade.upper()} ({term_val.upper()}, 2027)</h4>
                        </div>
                        <div>{photo_html}</div>
                    </div>

                    <table style="width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 12px;">
                        <tr>
                            <td style="border: 1px solid #374151; padding: 5px;"><b>NAME:</b> {d_prev['student_name']}</td>
                            <td style="border: 1px solid #374151; padding: 5px;"><b>GRADE:</b> {analysis_grade}</td>
                            <td style="border: 1px solid #374151; padding: 5px;"><b>GENDER:</b> {d_prev['gender_str']}</td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid #374151; padding: 5px;"><b>ADM NO:</b> {d_prev['adm_str']}</td>
                            <td style="border: 1px solid #374151; padding: 5px;"><b>ASSESS. NO:</b> {d_prev['assess_str']}</td>
                            <td style="border: 1px solid #374151; padding: 5px;"><b>RANK:</b> #{d_prev['student_rank']} / {len(active_ranked_students)}</td>
                        </tr>
                    </table>

                    <table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 10px;">
                        <tr style="background-color: #1E3A8A; color: #FFFFFF;">
                            <th style="border: 1px solid #374151; padding: 5px;">No.</th>
                            <th style="border: 1px solid #374151; padding: 5px; text-align: left;">Learning Area</th>
                            <th style="border: 1px solid #374151; padding: 5px;">Marks (%)</th>
                            <th style="border: 1px solid #374151; padding: 5px;">Perf. Level</th>
                            <th style="border: 1px solid #374151; padding: 5px;">Points</th>
                            <th style="border: 1px solid #374151; padding: 5px;">Position</th>
                        </tr>
                        {rows_html}
                        <tr style="background-color: #E5E7EB; font-weight: bold;">
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">-</td>
                            <td style="border: 1px solid #374151; padding: 5px;">OVERALL: {d_prev['overall_g']}</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">TOT: {d_prev['total_marks']:.0f}</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">-</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">PTS: {d_prev['total_points']}</td>
                            <td style="border: 1px solid #374151; padding: 5px; text-align: center;">MEAN: {d_prev['mean_points']}</td>
                        </tr>
                    </table>

                    <div style="font-size: 11.5px; background-color: #F9FAFB; padding: 8px; border: 1px solid #D1D5DB; border-radius: 4px;">
                        <p style="margin: 2px 0;"><strong>CLASS TEACHER:</strong> {d_prev['teacher_comment']}</p>
                        {dos_section}
                        <p style="margin: 4px 0 2px 0;"><strong>HEAD OF INSTITUTION:</strong> {d_prev['hoi_comment']}</p>
                    </div>

                    <div style="display: flex; justify-content: space-between; margin-top: 10px; font-size: 11.5px; font-weight: bold;">
                        <div>CLOSING DATE: {closing_date}</div>
                        <div>OPENING DATE: {opening_date}</div>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 22px; font-size: 11.5px; font-weight: bold;">
                        <div>CLASS TEACHER SIGN: ____________________</div>
                        <div style="position: relative; text-align: right;">
                            {f'<img src="data:image/png;base64,{stamp_b64}" style="position: absolute; right: 10px; bottom: -8px; height: 60px; opacity: 0.85;" />' if stamp_b64 else ''}
                            HOI SIGN & STAMP: ____________________
                        </div>
                    </div>
                </div>
                """
                st.components.v1.html(html_preview, height=740, scrolling=True)

                doc_obj = generate_student_docx_object(selected_student_obj)
                out_stream = io.BytesIO()
                doc_obj.save(out_stream)
                out_stream.seek(0)
                st.download_button(
                    label=f"📥 Download Single Report for {d_prev['student_name']} (.docx)",
                    data=out_stream,
                    file_name=f"Rank_{d_prev['student_rank']}_{d_prev['adm_str']}_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

            st.markdown("---")
            st.subheader("2. Batch Download All Ranked Reports for Whole Grade (.zip)")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for s_item in active_ranked_students:
                    single_doc = generate_student_docx_object(s_item)
                    doc_bytes = io.BytesIO()
                    single_doc.save(doc_bytes)
                    doc_bytes.seek(0)
                    safe_name = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in s_item["name"])
                    file_filename = f"Rank_{s_item['rank']:02d}_{s_item['adm_no']}_{safe_name}_Report.docx"
                    zip_file.writestr(file_filename, doc_bytes.getvalue())
            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 Download All {len(active_ranked_students)} Ranked Reports for {analysis_grade} (.zip)",
                data=zip_buffer,
                file_name=f"{analysis_grade.replace(' ', '_')}_All_Ranked_Reports.zip",
                mime="application/zip",
                type="secondary"
            )

# --- PAGE 5: TEACHERS PORTAL ---
elif page == "Teachers Portal":
    st.header(f"Teacher Portal — {st.session_state.full_name}")
    user_key = st.session_state.username
    teacher_info = TEACHER_ASSIGNMENTS.get(user_key)
    if teacher_info:
        for item in teacher_info["assignments"]:
            st.success(f"📌 **{item['grade']}:** Teaching **{', '.join(item['subjects'])}**")
    elif is_admin_role:
        try:
            t_res = supabase.table("teachers").select("full_name", "username", "designation").execute()
            st.dataframe(pd.DataFrame(t_res.data), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"Could not load teachers directory: {e}")

# --- PAGE 6: NEWSLETTER ---
elif page == "Newsletter":
    st.header("School Newsletter & Memos (2027)")
    if is_admin_role:
        with st.form("newsletter_form"):
            news_title = st.text_input("Newsletter Title / Memo Subject")
            news_content = st.text_area("News Content & Announcements")
            if st.form_submit_button("Publish Newsletter"):
                try:
                    supabase.table("newsletter").insert({
                        "title": news_title,
                        "content": news_content,
                        "date": datetime.now().strftime("%Y-%m-%d")
                    }).execute()
                    st.success("Published successfully!")
                except Exception as e:
                    st.error(f"Error: {e}")
    try:
        n_res = supabase.table("newsletter").select("*").order("date", desc=True).execute()
        if n_res.data:
            for item in n_res.data:
                st.markdown(f"### {item['title']} (*{item['date']}*)")
                st.write(item['content'])
                st.download_button(
                    "📥 Download Memo",
                    f"KEA MEMO\nDate: {item['date']}\nTitle: {item['title']}\n\n{item['content']}".encode('utf-8'),
                    f"Memo_{item['date']}.txt",
                    "text/plain",
                    key=f"memo_{item.get('id', item['title'])}_{item['date']}"
                )
                st.markdown("---")
    except Exception:
        pass