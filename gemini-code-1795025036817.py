import streamlit as st
import pandas as pd
from supabase import create_client, Client
from datetime import datetime
import io
import zipfile
import os
import base64
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from PIL import Image

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
    except Exception as e:
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
    "Balla":{
        "assignments":[
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

def generate_teacher_comment(total_points, top_subject, low_subject, grade_str):
    overall_grade = calculate_total_grade(total_points, grade_str)
    return (
        f"An impressive academic effort! Outstanding performance was demonstrated in {top_subject}, "
        f"reflecting dedication and grasp of concepts. To elevate overall standing further, "
        f"more focused revision and practice should be channeled into {low_subject}. "
        f"Overall Achievement Level: {overall_grade}."
    )

def generate_hoi_comment(total_points, grade_str):
    overall_grade = calculate_total_grade(total_points, grade_str)
    return (
        f"The administration notes steady academic progress and commitment shown during the term "
        f"(Overall Level: {overall_grade}). Learners are encouraged to maintain consistency, "
        f"embrace diligence in all learning areas, and put forth greater effort towards academic excellence."
    )

def generate_dos_comment(total_points, grade_str):
    overall_grade = calculate_total_grade(total_points, grade_str)
    return (
        f"The Director of Studies commends the learner's determination and discipline ({overall_grade}). "
        f"Continued intellectual curiosity, active participation in school programs, "
        f"and rigorous personal discipline will unlock even higher potential in future assessments."
    )

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
    tab_manual, tab_excel, tab_manage = st.tabs(["Manual Registration", "Excel Spreadsheet Upload", "Manage Registered Students"])
    
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
                        photo_b64 = ""
                        if photo_file is not None:
                            photo_bytes = photo_file.read()
                            photo_b64 = f"data:image/{photo_file.type.split('/')[-1]};base64," + base64.b64encode(photo_bytes).decode()

                        supabase.table("students").upsert({
                            "adm_no": adm_no,
                            "assessment_no": assessment_no,
                            "name": name,
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
                        supabase.table("students").upsert(records_to_update).execute()
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
                            st.success(f"Successfully deregistered student.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error deregistering student: {e}")
                else:
                    grade_to_clear = st.selectbox("Select Whole Grade to Deregister", [f"Grade {i}" for i in range(1, 10)])
                    confirm_grade_del = st.button("⚠️ Deregister All Learners in Grade", type="secondary")
                    if confirm_grade_del:
                        try:
                            grade_students = [s for s in students_list if str(s.get("grade","")).strip().lower() == grade_to_clear.lower()]
                            for gs in grade_students:
                                supabase.table("marks").delete().eq("adm_no", gs["adm_no"]).execute()
                                supabase.table("students").delete().eq("adm_no", gs["adm_no"]).execute()
                            st.success(f"Successfully deregistered all learners in {grade_to_clear}.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error clearing grade: {e}")
        except Exception as e:
            st.error(f"Error loading students: {e}")

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
        except Exception as e:
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
    except Exception as e:
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

    with tab_overview:
        if active_ranked_students:
            st.markdown("### 🏆 Top Ten Students")
            top_10 = [{
                "Rank": i["rank"], "Adm No": i["adm_no"], "Student Name": i["name"], 
                "Total Marks": i["total_marks"], "Total Points": i["total_points"], 
                "Mean Points": i["mean_points"], "Grade": calculate_total_grade(i["total_points"], analysis_grade)
            } for i in active_ranked_students[:10]]
            st.dataframe(pd.DataFrame(top_10), use_container_width=True)

    with tab_ranking:
        if active_ranked_students:
            rows = [{
                "Rank": i["rank"], "Adm No": i["adm_no"], "Student Name": i["name"], 
                "Total Marks": i["total_marks"], "Total Points": i["total_points"], 
                "Mean Points": i["mean_points"], "Grade": calculate_total_grade(i["total_points"], analysis_grade)
            } for i in active_ranked_students]
            df_res = pd.DataFrame(rows)
            st.dataframe(df_res, use_container_width=True)
            st.download_button("📥 Download Roll (.csv)", df_res.to_csv(index=False).encode('utf-8'), f"{analysis_grade}_Roll.csv", "text/csv")

    with tab_top_subjects:
        for sub in current_grade_learning_areas:
            col_k = get_db_col_name(sub)
            sub_studs = []
            for item in active_ranked_students:
                val = item["marks_row"].get(col_k)
                if val is not None and not pd.isna(val) and float(val) >= 0.0:
                    code, _, pts = get_subject_performance(float(val), analysis_grade)
                    sub_studs.append({"Adm No": item["adm_no"], "Student Name": item["name"], "Score (%)": float(val), "Perf. Level": code, "Points": pts})
            sub_studs = sorted(sub_studs, key=lambda x: x["Score (%)"], reverse=True)[:3]
            with st.expander(f"📚 {sub}", expanded=True):
                if sub_studs: st.dataframe(pd.DataFrame(sub_studs), use_container_width=True, hide_index=True)
                else: st.caption("No scores recorded.")

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
            template_filename = "report3.docx" if "grade 1" in g_lower or "grade 2" in g_lower or "grade 3" in g_lower else ("report2.docx" if "grade 4" in g_lower or "grade 5" in g_lower or "grade 6" in g_lower else "report1.docx")

            def generate_student_docx_object(student_obj):
                adm_str = student_obj["adm_no"]
                assess_str = student_obj.get("assessment_no", "")
                student_name = student_obj["name"]
                gender_str = student_obj.get("gender", "")
                s_row = student_obj["marks_row"]
                student_rank = student_obj["rank"]
                total_marks = student_obj["total_marks"]
                total_points = student_obj["total_points"]
                mean_points = student_obj["mean_points"]

                doc = Document(template_filename) if os.path.exists(template_filename) else (Document("report.docx") if os.path.exists("report.docx") else Document())
                for section in doc.sections:
                    section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = Inches(0.3), Inches(0.3), Inches(0.4), Inches(0.4)

                if os.path.exists("logo.png"):
                    p_logo = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
                    p_logo.add_run().add_picture("logo.png", width=Inches(0.9))

                learning_list = get_learning_areas_for_grade(analysis_grade)
                top_sub, low_sub = learning_list[0], learning_list[0]
                best_s, worst_s = -1, 101
                
                subject_details = []
                for sub in learning_list:
                    col_key = get_db_col_name(sub)
                    val = s_row.get(col_key)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        s_num = float(val)
                        if s_num > best_s: best_s = s_num; top_sub = sub
                        if s_num < worst_s: worst_s = s_num; low_sub = sub
                        code, _, pts = get_subject_performance(s_num, analysis_grade)
                        
                        sub_scores_list = sorted([(st_item["adm_no"], float(st_item["marks_row"].get(col_key))) for st_item in active_ranked_students if st_item["marks_row"].get(col_key) is not None], key=lambda x: x[1], reverse=True)
                        sub_rank = next((i+1 for i, (a, _) in enumerate(sub_scores_list) if a == adm_str), 1)
                        pos_str = f"{sub_rank}/{len(sub_scores_list)}"
                        subject_details.append({"name": sub.upper(), "marks": f"{s_num:.0f}", "level": code, "points": str(pts), "position": pos_str, "teacher": "TBA"})
                    else:
                        subject_details.append({"name": sub.upper(), "marks": "-", "level": "-", "points": "-", "position": "-", "teacher": "TBA"})

                overall_g = calculate_total_grade(total_points, analysis_grade)
                teacher_comment = generate_teacher_comment(total_points, top_sub, low_sub, analysis_grade)
                hoi_comment = generate_hoi_comment(total_points, analysis_grade)
                dos_comment = generate_dos_comment(total_points, analysis_grade) if any(g in analysis_grade.lower() for g in ["grade 7", "grade 8", "grade 9"]) else ""

                for t in doc.tables:
                    for row in t.rows:
                        row_text = " ".join([cell.text for cell in row.cells]).upper()
                        for cell in row.cells:
                            cu = cell.text.upper()
                            if "NAME:" in cu: cell.text = f"NAME: {student_name}"
                            elif "ADM NO:" in cu: cell.text = f"ADM NO: {adm_str}"
                            elif "ASSES" in cu: cell.text = f"ASSES.NO: {assess_str}"
                            elif "GENDER:" in cu: cell.text = f"GENDER: {gender_str}"
                            elif "TERM:" in cu: cell.text = f"TERM: {term_val}"
                            elif "POSITION:" in cu: cell.text = f"POSITION: #{student_rank} / {len(active_ranked_students)}"
                        for sd in subject_details:
                            if sd["name"] in row_text and len(row.cells) >= 6:
                                row.cells[1].text = sd["marks"]
                                row.cells[2].text = sd["level"]
                                row.cells[3].text = sd["points"]
                                row.cells[4].text = sd["position"]
                                row.cells[5].text = sd["teacher"]

                if stamp_upload or os.path.exists("stamp.png"):
                    stamp_source = stamp_upload if stamp_upload else "stamp.png"
                    for p in doc.paragraphs:
                        if "STAMP" in p.text.upper() or "SIGN" in p.text.upper():
                            p.add_run().add_picture(stamp_source, width=Inches(1.1))
                            break
                return doc

            st.subheader("1. Live Student Report Form Preview")
            selected_student_obj = st.selectbox("Select Student for Live Report Form Preview", active_ranked_students, format_func=lambda x: f"Rank #{x['rank']} | Adm No: {x['adm_no']} | Name: {x['name']}")

            if selected_student_obj:
                doc_obj = generate_student_docx_object(selected_student_obj)
                s_name, s_adm, s_assess, s_gender, s_photo = selected_student_obj['name'], selected_student_obj['adm_no'], selected_student_obj.get('assessment_no', ''), selected_student_obj.get('gender', ''), selected_student_obj.get('photo', '')
                s_rank, s_total_m, s_total_p, s_mean_p, s_row = selected_student_obj['rank'], selected_student_obj['total_marks'], selected_student_obj['total_points'], selected_student_obj['mean_points'], selected_student_obj['marks_row']
                overall_g = calculate_total_grade(s_total_p, analysis_grade)

                learning_list = get_learning_areas_for_grade(analysis_grade)
                top_sub, low_sub = learning_list[0], learning_list[0]
                best_s, worst_s = -1, 101

                rows_html = ""
                for sub in learning_list:
                    col_key = get_db_col_name(sub)
                    val = s_row.get(col_key)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        s_num = float(val)
                        if s_num > best_s: best_s = s_num; top_sub = sub
                        if s_num < worst_s: worst_s = s_num; low_sub = sub
                        code, _, pts = get_subject_performance(s_num, analysis_grade)
                        sub_scores_list = sorted([(st_item["adm_no"], float(st_item["marks_row"].get(col_key))) for st_item in active_ranked_students if st_item["marks_row"].get(col_key) is not None], key=lambda x: x[1], reverse=True)
                        sub_rank = next((i+1 for i, (a, _) in enumerate(sub_scores_list) if a == s_adm), 1)
                        pos_str = f"{sub_rank}/{len(sub_scores_list)}"
                        m_str, lvl_str, pts_str = f"{s_num:.0f}", code, str(pts)
                    else:
                        m_str, lvl_str, pts_str, pos_str = "-", "-", "-", "-"

                    rows_html += f"""
                        <tr>
                            <td style="border: 1px solid #374151; padding: 6px; font-weight: bold;">{sub.upper()}</td>
                            <td style="border: 1px solid #374151; padding: 6px; text-align: center;">{m_str}</td>
                            <td style="border: 1px solid #374151; padding: 6px; text-align: center;">{lvl_str}</td>
                            <td style="border: 1px solid #374151; padding: 6px; text-align: center;">{pts_str}</td>
                            <td style="border: 1px solid #374151; padding: 6px; text-align: center;">{pos_str}</td>
                            <td style="border: 1px solid #374151; padding: 6px; text-align: center;">TBA</td>
                        </tr>
                    """

                teacher_comment = generate_teacher_comment(s_total_p, top_sub, low_sub, analysis_grade)
                hoi_comment = generate_hoi_comment(s_total_p, analysis_grade)
                dos_section = f"""
                    <p style="margin: 8px 0 2px 0;"><strong>DIRECTOR OF STUDIES COMMENT:</strong></p>
                    <div style="border: 1px dashed #374151; padding: 6px; font-style: italic; background-color: #F9FAFB;">"{generate_dos_comment(s_total_p, analysis_grade)}"</div>
                """ if any(g in analysis_grade.lower() for g in ["grade 7", "grade 8", "grade 9"]) else ""

                logo_b64 = ""
                if os.path.exists("logo.png"):
                    with open("logo.png", "rb") as f: logo_b64 = base64.b64encode(f.read()).decode()

                stamp_b64 = ""
                if stamp_upload: stamp_b64 = base64.b64encode(stamp_upload.getvalue()).decode()
                elif os.path.exists("stamp.png"):
                    with open("stamp.png", "rb") as f: stamp_b64 = base64.b64encode(f.read()).decode()

                photo_html = f'<div style="width: 80px; height: 95px; border: 1px solid #374151; display: flex; align-items: center; justify-content: center; background: #E5E7EB;"><img src="{s_photo}" style="max-width: 100%; max-height: 100%; object-fit: cover;" /></div>' if s_photo else '<div style="width: 80px; height: 95px; border: 1px dashed #9CA3AF; display: flex; align-items: center; justify-content: center; font-size: 10px; color: #6B7280; text-align: center;">PHOTO</div>'

                st.markdown("#### 👁️ Live Report Form Document View")
                html_preview = f"""
                <div style="border: 2px solid #1E3A8A; padding: 24px; border-radius: 8px; background-color: #FFFFFF; color: #111827; font-family: Arial, sans-serif;">
                    <div style="display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px; margin-bottom: 12px;">
                        <div>{f'<img src="data:image/png;base64,{logo_b64}" style="height: 70px;" />' if logo_b64 else ''}</div>
                        <div style="text-align: center; flex-grow: 1;">
                            <h2 style="margin: 0; color: #1E3A8A; font-size: 20px;">KEA COMPREHENSIVE SCHOOL</h2>
                            <p style="margin: 2px 0; font-size: 12px;">P.O. BOX 557-40400 SUNA MIGORI</p>
                            <h4 style="margin: 4px 0; text-decoration: underline; font-size: 14px;">STUDENT ASSESSMENT REPORT ({term_val} - 2027)</h4>
                        </div>
                        <div>{photo_html}</div>
                    </div>

                    <table style="width: 100%; border-collapse: collapse; margin-bottom: 14px; font-size: 12px;">
                        <tr>
                            <td style="border: 1px solid #374151; padding: 6px;"><b>NAME:</b> {s_name}</td>
                            <td style="border: 1px solid #374151; padding: 6px;"><b>GRADE:</b> {analysis_grade}</td>
                            <td style="border: 1px solid #374151; padding: 6px;"><b>GENDER:</b> {s_gender}</td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid #374151; padding: 6px;"><b>ADM NO:</b> {s_adm}</td>
                            <td style="border: 1px solid #374151; padding: 6px;"><b>ASSESSMENT NO:</b> {s_assess}</td>
                            <td style="border: 1px solid #374151; padding: 6px;"><b>OVERALL RANK:</b> #{s_rank} / {len(active_ranked_students)}</td>
                        </tr>
                    </table>

                    <table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 14px;">
                        <tr style="background-color: #1E3A8A; color: #FFFFFF;">
                            <th style="border: 1px solid #374151; padding: 6px; text-align: left;">LEARNING AREAS</th>
                            <th style="border: 1px solid #374151; padding: 6px;">MARKS</th>
                            <th style="border: 1px solid #374151; padding: 6px;">PERF. LEVEL</th>
                            <th style="border: 1px solid #374151; padding: 6px;">POINTS</th>
                            <th style="border: 1px solid #374151; padding: 6px;">POSITION</th>
                            <th style="border: 1px solid #374151; padding: 6px;">TEACHER</th>
                        </tr>
                        {rows_html}
                    </table>

                    <div style="background-color: #F3F4F6; padding: 10px; border: 1px solid #374151; font-size: 12px; margin-bottom: 12px;">
                        <b>TOTAL MARKS:</b> {s_total_m:.0f} &nbsp;|&nbsp; 
                        <b>TOTAL POINTS:</b> {s_total_p} &nbsp;|&nbsp; 
                        <b>MEAN POINTS:</b> {s_mean_p} &nbsp;|&nbsp; 
                        <b>OVERALL PERFORMANCE LEVEL:</b> <span style="color: #1E3A8A; font-weight: bold;">{overall_g}</span>
                    </div>

                    <div style="font-size: 12px;">
                        <p style="margin: 6px 0 2px 0;"><strong>CLASS TEACHER'S COMMENT:</strong></p>
                        <div style="border: 1px dashed #374151; padding: 6px; font-style: italic; background-color: #F9FAFB;">"{teacher_comment}"</div>
                        {dos_section}
                        <p style="margin: 8px 0 2px 0;"><strong>HEAD OF INSTITUTION COMMENT:</strong></p>
                        <div style="border: 1px dashed #374151; padding: 6px; font-style: italic; background-color: #F9FAFB;">"{hoi_comment}"</div>
                    </div>

                    <div style="display: flex; justify-content: space-between; margin-top: 15px; font-size: 12px; font-weight: bold;">
                        <div>CLOSING DATE: {closing_date}</div>
                        <div>OPENING DATE: {opening_date}</div>
                    </div>

                    <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 30px; font-size: 12px; font-weight: bold;">
                        <div>CLASS TEACHER SIGN: ____________________</div>
                        <div style="position: relative; text-align: right;">
                            {f'<img src="data:image/png;base64,{stamp_b64}" style="position: absolute; right: 15px; bottom: -10px; height: 70px; opacity: 0.85;" />' if stamp_b64 else ''}
                            SIGN & STAMP: ____________________
                        </div>
                    </div>
                </div>
                """
                st.components.v1.html(html_preview, height=800, scrolling=True)

                out_stream = io.BytesIO()
                doc_obj.save(out_stream)
                out_stream.seek(0)
                st.download_button("📥 Download Assessment Report (.docx)", out_stream, file_name=f"Rank_{s_rank}_{s_adm}_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")

# --- PAGE 5: TEACHERS PORTAL ---
elif page == "Teachers Portal":
    st.header(f"Teacher Portal — {st.session_state.full_name}")
    user_key = st.session_state.username
    teacher_info = TEACHER_ASSIGNMENTS.get(user_key)
    if teacher_info:
        for item in teacher_info["assignments"]:
            st.success(f"📌 **{item['grade']}:** Teaching **{', '.join(item['subjects'])}**")

# --- PAGE 6: NEWSLETTER ---
elif page == "Newsletter":
    st.header("School Newsletter & Memos (2027)")
    if is_admin_role:
        with st.form("newsletter_form"):
            news_title = st.text_input("Newsletter Title / Memo Subject")
            news_content = st.text_area("News Content & Announcements")
            if st.form_submit_button("Publish Newsletter"):
                try:
                    supabase.table("newsletter").insert({"title": news_title, "content": news_content, "date": datetime.now().strftime("%Y-%m-%d")}).execute()
                    st.success("Published successfully!")
                except Exception as e:
                    st.error(f"Error: {e}")
    try:
        n_res = supabase.table("newsletter").select("*").order("date", desc=True).execute()
        if n_res.data:
            for item in n_res.data:
                st.markdown(f"### {item['title']} (*{item['date']}*)")
                st.write(item['content'])
                st.download_button("📥 Download Memo", f"KEA MEMO\nDate: {item['date']}\nTitle: {item['title']}\n\n{item['content']}".encode('utf-8'), f"Memo_{item['date']}.txt", "text/plain")
                st.markdown("---")
    except:
        pass