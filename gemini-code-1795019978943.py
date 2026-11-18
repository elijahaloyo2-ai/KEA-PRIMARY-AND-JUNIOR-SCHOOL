import streamlit as st
import pandas as pd
from supabase import create_client, Client
from datetime import datetime
import io
import zipfile
import os
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

# Union of all possible learning areas for database mapping / migration
ALL_LEARNING_AREAS = list(set(LOWER_PRIMARY_SUBJECTS + UPPER_PRIMARY_SUBJECTS + JUNIOR_SCHOOL_SUBJECTS))

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
        # Junior School grading (out of 100%)
        if score <= 10: return "BE2", "Below Expectation 2", 1
        elif score <= 20: return "BE1", "Below Expectation 1", 2
        elif score <= 30: return "AE2", "Approaching Expectation 2", 3
        elif score <= 40: return "AE1", "Approaching Expectation 1", 4
        elif score <= 57: return "ME2", "Meeting Expectation 2", 5
        elif score <= 74: return "ME1", "Meeting Expectation 1", 6
        elif score <= 89: return "EE2", "Exceeding Expectation 2", 7
        else: return "EE1", "Exceeding Expectation 1", 8
    else:
        # Lower & Upper Primary grading (out of 100%)
        if score <= 24: return "BE", "Below Expectation", 1
        elif score <= 49: return "AE", "Approaching Expectation", 2
        elif score <= 74: return "ME", "Meeting Expectation", 3
        else: return "EE", "Exceeding Expectation", 4

def calculate_total_grade(total_points, grade_str):
    g_lower = grade_str.strip().lower()
    is_junior = any(g in g_lower for g in ["grade 7", "grade 8", "grade 9"])
    is_lower = any(g in g_lower for g in ["grade 1", "grade 2", "grade 3"])

    if is_junior:
        # Total points out of 72 (9 learning areas @ 8 points max)
        if total_points <= 13: return "BE2 (Below Expectation 2)"
        elif total_points <= 22: return "BE1 (Below Expectation 1)"
        elif total_points <= 31: return "AE2 (Approaching Expectation 2)"
        elif total_points <= 40: return "AE1 (Approaching Expectation 1)"
        elif total_points <= 49: return "ME2 (Meeting Expectation 2)"
        elif total_points <= 58: return "ME1 (Meeting Expectation 1)"
        elif total_points <= 67: return "EE2 (Exceeding Expectation 2)"
        else: return "EE1 (Exceeding Expectation 1)"
    elif is_lower:
        # Total points out of 24 (6 learning areas @ 4 points max)
        if total_points <= 8: return "BE (Below Expectation)"
        elif total_points <= 14: return "AE (Approaching Expectation)"
        elif total_points <= 20: return "ME (Meeting Expectation)"
        else: return "EE (Exceeding Expectation)"
    else:
        # Upper primary out of 32 (8 learning areas @ 4 points max)
        if total_points <= 11: return "BE (Below Expectation)"
        elif total_points <= 19: return "AE (Approaching Expectation)"
        elif total_points <= 27: return "ME (Meeting Expectation)"
        else: return "EE (Exceeding Expectation)"

# --- COMMENT GENERATOR HELPERS ---
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
        total_teachers = 0

        if supabase:
            try:
                students_res = supabase.table("students").select("*", count="exact").execute()
                teachers_res = supabase.table("teachers").select("*", count="exact").execute()
                total_students = students_res.count if students_res.count is not None else len(students_res.data)
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

        m1, m2 = st.columns(2)
        with m1:
            st.markdown(f"""
                <div class="metric-card">
                    <small style="color:#94A3B8;">TOTAL REGISTERED STUDENTS</small>
                    <h2 style="margin:5px 0; color:#60A5FA;">{total_students}</h2>
                    <small style="color:#34D399;">Active Enrolment (2027)</small>
                </div>
            """, unsafe_allow_html=True)
        with m2:
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
        # Teacher Dashboard
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

            st.markdown("---")
            st.markdown("### 📈 Learning Area Submission & Performance Analytics")
            analytics_data = []
            try:
                students_res = supabase.table("students").select("adm_no, grade").execute()
                all_students = students_res.data if students_res.data else []
                marks_res = supabase.table("marks").select("*").execute()
                all_marks = marks_res.data if marks_res.data else []
                marks_map_by_grade_and_adm = {f"{m.get('grade')}_{m.get('adm_no')}": m for m in all_marks}

                for assignment in assignments:
                    grade = assignment["grade"]
                    grade_num = grade.replace("Grade", "").strip()
                    grade_students = [
                        s for s in all_students 
                        if str(s.get("grade", "")).strip().lower() in [grade.lower(), grade_num, f"grade {grade_num}"]
                    ]
                    total_grade_students = len(grade_students)

                    for subj in assignment["subjects"]:
                        db_col = get_db_col_name(subj)
                        submitted_count = 0
                        scores_list = []
                        for student in grade_students:
                            adm = student["adm_no"]
                            m_row = marks_map_by_grade_and_adm.get(f"{grade}_{adm}", {})
                            val = m_row.get(db_col)
                            if val is not None and not pd.isna(val) and float(val) >= 0.0:
                                submitted_count += 1
                                scores_list.append(float(val))

                        pending_count = max(0, total_grade_students - submitted_count)
                        mean_score = (sum(scores_list) / len(scores_list)) if scores_list else 0.0

                        analytics_data.append({
                            "Grade": grade,
                            "Learning Area": subj,
                            "Enrolled": total_grade_students,
                            "Submitted": submitted_count,
                            "Pending": pending_count,
                            "Mean Score (%)": round(mean_score, 1)
                        })

                if analytics_data:
                    df_analytics = pd.DataFrame(analytics_data)
                    tot_enrolled = sum(d["Enrolled"] for d in analytics_data)
                    tot_submitted = sum(d["Submitted"] for d in analytics_data)
                    tot_pending = sum(d["Pending"] for d in analytics_data)
                    overall_avg = sum(d["Mean Score (%)"] * d["Submitted"] for d in analytics_data) / tot_submitted if tot_submitted > 0 else 0.0

                    tm1, tm2, tm3, tm4 = st.columns(4)
                    tm1.metric("Total Enrolled Students", tot_enrolled)
                    tm2.metric("Marks Submitted", tot_submitted, delta=f"{round((tot_submitted/tot_enrolled)*100 if tot_enrolled>0 else 0)}% completion")
                    tm3.metric("Pending Submissions", tot_pending, delta_color="inverse")
                    tm4.metric("Average Subject Mean", f"{round(overall_avg, 1)}%")

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.dataframe(df_analytics, use_container_width=True, hide_index=True)
                else:
                    st.info("No submission data available for your assigned classes yet.")
            except Exception as e:
                st.error(f"Error compiling teacher analytics: {e}")

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
            
            submit_manual = st.form_submit_button("Admit Student")
            
            if submit_manual:
                if adm_no and name and grade and (not is_junior or assessment_no):
                    try:
                        supabase.table("students").insert({
                            "adm_no": adm_no,
                            "assessment_no": assessment_no,
                            "name": name,
                            "gender": gender,
                            "grade": grade
                        }).execute()
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

                if not records:
                    st.warning("No valid student records found in the uploaded file.")
                else:
                    unique_records_dict = {r["adm_no"]: r for r in records}
                    deduplicated_records = list(unique_records_dict.values())
                    supabase.table("students").upsert(deduplicated_records, on_conflict="adm_no").execute()
                    st.success(f"Successfully uploaded and admitted {len(deduplicated_records)} students!")
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
                st.markdown("#### Deregister Options (Select Learner or Whole Grade)")
                
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
                        name_del = student_to_delete["name"]
                        try:
                            supabase.table("marks").delete().eq("adm_no", adm_del).execute()
                            supabase.table("students").delete().eq("adm_no", adm_del).execute()
                            st.success(f"Successfully deregistered {name_del} (Adm No: {adm_del}).")
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
        teacher_grade_map = {}
        for item in teacher_data["assignments"]:
            teacher_grade_map[item["grade"]] = item["subjects"]
        allowed_grades = list(teacher_grade_map.keys())
        st.info(f"Welcome, {st.session_state.full_name}! Select an assigned grade and learning area below.")

    if not allowed_grades:
        st.warning("No grades currently assigned to your profile.")
        st.stop()

    col_g, col_s = st.columns(2)
    with col_g:
        target_grade = st.selectbox("1. Select Grade", allowed_grades)
    
    grade_learning_areas = get_learning_areas_for_grade(target_grade)
    if not is_admin_role and teacher_data:
        assigned_subs_for_grade = teacher_grade_map.get(target_grade, grade_learning_areas)
    else:
        assigned_subs_for_grade = grade_learning_areas
    
    with col_s:
        target_subject = st.selectbox("2. Select Learning Area", assigned_subs_for_grade)

    st.markdown("---")
    st.subheader(f"Marking Sheet: {target_grade} — {target_subject}")
    st.caption("Enter scores out of 100%. Unrecorded students appear for entry, while saved entries appear in the 'Recorded' table with direct 'Edit & Save' capabilities.")

    grade_num = target_grade.replace("Grade", "").strip()
    try:
        students_res = supabase.table("students").select("adm_no, name, grade").execute()
        all_students = students_res.data if students_res.data else []
        students = [
            s for s in all_students 
            if str(s.get("grade", "")).strip().lower() in [target_grade.lower(), grade_num, f"grade {grade_num}"]
        ]
        students = sorted(students, key=lambda x: str(x.get("adm_no", "")))
    except Exception as e:
        students = []
        st.error(f"Error fetching students from database: {e}")

    if not students:
        st.warning(f"No students found under '{target_grade}' in the database.")
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
            st.caption(f"Note: Could not load prior marks automatically: {e}")

        unrecorded_students = [s for s in students if s["adm_no"] not in existing_marks_map]
        recorded_students = [s for s in students if s["adm_no"] in existing_marks_map]

        tab_manual_marks, tab_excel_marks = st.tabs(["Manual Sheet Input", "Upload Excel Marks"])

        with tab_manual_marks:
            st.markdown("### 📝 Active Marking Table (Unrecorded Students)")
            if not unrecorded_students:
                st.success("🎉 All students in this grade have their marks recorded for this learning area!")
            else:
                with st.form(key=f"unrecorded_form_{target_grade}_{col_key}"):
                    editor_rows = []
                    for s in unrecorded_students:
                        editor_rows.append({
                            "Adm No": s["adm_no"],
                            "Student Name": s["name"],
                            "Score (%)": None,
                            "Save & Submit": False
                        })

                    df_unrecorded = pd.DataFrame(editor_rows)
                    edited_df = st.data_editor(
                        df_unrecorded,
                        column_config={
                            "Adm No": st.column_config.TextColumn("Adm No", disabled=True),
                            "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                            "Score (%)": st.column_config.NumberColumn("Score (%)", min_value=0.0, max_value=100.0, step=1.0, format="%.0f"),
                            "Save & Submit": st.column_config.CheckboxColumn("Save & Submit", default=False)
                        },
                        use_container_width=True,
                        hide_index=True,
                        key=f"editor_table_{target_grade}_{col_key}"
                    )

                    col_b1, col_b2 = st.columns([1, 1])
                    with col_b1:
                        submit_all = st.form_submit_button("💾 Save and Submit All Entered Marks", type="primary")
                    with col_b2:
                        submit_checked = st.form_submit_button("✅ Save Selected (Checked Rows)", type="secondary")

                    if submit_all or submit_checked:
                        records_to_upsert = []
                        submitted_count = 0
                        grade_learning_list = get_learning_areas_for_grade(target_grade)
                        
                        for _, row in edited_df.iterrows():
                            adm = str(row["Adm No"])
                            raw_score = row["Score (%)"]
                            is_checked = row["Save & Submit"]
                            
                            should_save = (submit_all and raw_score is not None and not pd.isna(raw_score)) or \
                                          (submit_checked and is_checked and raw_score is not None and not pd.isna(raw_score))

                            if should_save:
                                existing = existing_full_rows.get(adm, {})
                                rec = dict(existing) if existing else {"adm_no": adm, "grade": target_grade}
                                rec[col_key] = float(raw_score)
                                
                                # Recalculate total marks and total points
                                tot_marks = 0.0
                                tot_pts = 0
                                for sub in grade_learning_list:
                                    s_key = get_db_col_name(sub)
                                    val = rec.get(s_key)
                                    if val is not None and not pd.isna(val) and float(val) >= 0:
                                        s_val = float(val)
                                        tot_marks += s_val
                                        _, _, pts = get_subject_performance(s_val, target_grade)
                                        if pts:
                                            tot_pts += pts

                                rec["total_marks"] = tot_marks
                                rec["total_points"] = tot_pts
                                records_to_upsert.append(rec)
                                submitted_count += 1

                        if records_to_upsert:
                            try:
                                supabase.table("marks").upsert(records_to_upsert, on_conflict="adm_no").execute()
                                st.success(f"Successfully saved marks for {submitted_count} student(s)!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Database error saving marks: {e}")

            st.markdown("---")
            st.markdown("### ✅ Recorded Students Table (Edit & Save)")
            if not recorded_students:
                st.info("No marks recorded yet for this learning area.")
            else:
                recorded_rows = []
                for s in recorded_students:
                    adm = s["adm_no"]
                    score_val = existing_marks_map.get(adm)
                    code, perf, pts = get_subject_performance(score_val, target_grade)
                    recorded_rows.append({
                        "Adm No": adm,
                        "Student Name": s["name"],
                        "Score (%)": float(score_val) if score_val is not None else 0.0,
                        "Perf. Level": code,
                        "Points": pts if pts else 0,
                        "Edit & Save": False
                    })
                
                df_recorded = pd.DataFrame(recorded_rows)
                with st.form(key=f"recorded_edit_form_{target_grade}_{col_key}"):
                    edited_recorded_df = st.data_editor(
                        df_recorded,
                        column_config={
                            "Adm No": st.column_config.TextColumn("Adm No", disabled=True),
                            "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                            "Score (%)": st.column_config.NumberColumn("Score (%)", min_value=0.0, max_value=100.0, step=1.0, format="%.0f"),
                            "Perf. Level": st.column_config.TextColumn("Perf. Level", disabled=True),
                            "Points": st.column_config.NumberColumn("Points", disabled=True),
                            "Edit & Save": st.column_config.CheckboxColumn("Edit & Save", default=False)
                        },
                        use_container_width=True,
                        hide_index=True,
                        key=f"editor_recorded_table_{target_grade}_{col_key}"
                    )

                    col_re1, col_re2 = st.columns([1, 1])
                    with col_re1:
                        submit_edited_recorded = st.form_submit_button("💾 Save All Checked Edits", type="primary")
                    with col_re2:
                        reset_selected_recorded = st.form_submit_button("🔄 Move Selected Back to Unrecorded", type="secondary")

                    if submit_edited_recorded:
                        records_to_update = []
                        updated_cnt = 0
                        grade_learning_list = get_learning_areas_for_grade(target_grade)
                        
                        for _, row in edited_recorded_df.iterrows():
                            if row["Edit & Save"]:
                                adm = str(row["Adm No"])
                                new_score = row["Score (%)"]
                                existing = existing_full_rows.get(adm, {})
                                rec = dict(existing) if existing else {"adm_no": adm, "grade": target_grade}
                                rec[col_key] = float(new_score) if new_score is not None and not pd.isna(new_score) else None
                                
                                tot_marks = 0.0
                                tot_pts = 0
                                for sub in grade_learning_list:
                                    s_key = get_db_col_name(sub)
                                    val = rec.get(s_key)
                                    if val is not None and not pd.isna(val) and float(val) >= 0:
                                        s_val = float(val)
                                        tot_marks += s_val
                                        _, _, pts = get_subject_performance(s_val, target_grade)
                                        if pts:
                                            tot_pts += pts

                                rec["total_marks"] = tot_marks
                                rec["total_points"] = tot_pts
                                records_to_update.append(rec)
                                updated_cnt += 1

                        if records_to_update:
                            try:
                                supabase.table("marks").upsert(records_to_update, on_conflict="adm_no").execute()
                                st.success(f"Successfully updated marks for {updated_cnt} student(s)!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error updating recorded marks: {e}")

                    if reset_selected_recorded:
                        records_to_reset = []
                        reset_cnt = 0
                        grade_learning_list = get_learning_areas_for_grade(target_grade)
                        
                        for _, row in edited_recorded_df.iterrows():
                            if row["Edit & Save"]:
                                adm = str(row["Adm No"])
                                existing = existing_full_rows.get(adm, {})
                                if existing:
                                    existing[col_key] = None
                                    tot_marks = 0.0
                                    tot_pts = 0
                                    for sub in grade_learning_list:
                                        s_key = get_db_col_name(sub)
                                        val = existing.get(s_key)
                                        if val is not None and not pd.isna(val) and float(val) >= 0:
                                            s_val = float(val)
                                            tot_marks += s_val
                                            _, _, pts = get_subject_performance(s_val, target_grade)
                                            if pts:
                                                tot_pts += pts
                                    existing["total_marks"] = tot_marks
                                    existing["total_points"] = tot_pts
                                    records_to_reset.append(existing)
                                    reset_cnt += 1

                        if records_to_reset:
                            try:
                                supabase.table("marks").upsert(records_to_reset, on_conflict="adm_no").execute()
                                st.success(f"Successfully moved {reset_cnt} student(s) back to unrecorded table.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error resetting marks: {e}")

        with tab_excel_marks:
            st.markdown(f"Upload an Excel sheet for **{target_grade} — {target_subject}**.")
            marks_file = st.file_uploader("Choose Excel File", type=["xlsx", "xls"], key=f"uploader_{target_grade}_{col_key}")
            if marks_file and st.button("Upload Excel Marks Sheet"):
                try:
                    df_upload = pd.read_excel(marks_file)
                    records_to_upsert = []
                    score_col = next((c for c in df_upload.columns if c.lower() in ["score", "marks", target_subject.lower()]), None)
                    if not score_col and len(df_upload.columns) >= 2:
                        score_col = df_upload.columns[1]

                    grade_learning_list = get_learning_areas_for_grade(target_grade)
                    for _, row in df_upload.iterrows():
                        adm = str(row["Adm No"]).strip()
                        raw_val = row[score_col]
                        try:
                            score_num = float(raw_val)
                        except:
                            score_num = 0.0

                        existing = {}
                        try:
                            ex_res = supabase.table("marks").select("*").eq("adm_no", adm).execute()
                            if ex_res.data:
                                existing = ex_res.data[0]
                        except:
                            pass

                        rec = existing if existing else {"adm_no": adm, "grade": target_grade}
                        rec[col_key] = score_num if score_num >= 0 else None
                        
                        tot_marks = 0.0
                        tot_pts = 0
                        for sub in grade_learning_list:
                            s_key = get_db_col_name(sub)
                            val = rec.get(s_key)
                            if val is not None and not pd.isna(val) and float(val) >= 0:
                                s_val = float(val)
                                tot_marks += s_val
                                _, _, pts = get_subject_performance(s_val, target_grade)
                                if pts:
                                    tot_pts += pts
                        rec["total_marks"] = tot_marks
                        rec["total_points"] = tot_pts
                        records_to_upsert.append(rec)

                    supabase.table("marks").upsert(records_to_upsert, on_conflict="adm_no").execute()
                    st.success(f"Successfully processed marks for {len(records_to_upsert)} students!")
                except Exception as e:
                    st.error(f"Error processing uploaded file: {e}")

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
        marks_data = [
            m for m in all_marks 
            if str(m.get("grade", "")).strip().lower() in [analysis_grade.lower(), grade_num, f"grade {grade_num}"]
        ]
        
        s_res = supabase.table("students").select("*").execute()
        students_dict = {}
        if s_res.data:
            for s in s_res.data:
                clean_adm = str(s.get("adm_no", "")).strip()
                students_dict[clean_adm] = s
    except Exception as e:
        marks_data, students_dict = [], {}
        st.error(f"Error fetching analysis data: {e}")

    def process_student_scores(m_row, grade_str):
        total_m = 0.0
        total_p = 0
        valid_count = 0
        learning_list = get_learning_areas_for_grade(grade_str)
        for sub in learning_list:
            k = get_db_col_name(sub)
            val = m_row.get(k)
            if val is not None and not pd.isna(val) and float(val) >= 0.0:
                s_val = float(val)
                total_m += s_val
                _, _, pts = get_subject_performance(s_val, grade_str)
                if pts:
                    total_p += pts
                valid_count += 1
        return total_m, total_p, valid_count

    active_ranked_students = []
    for m in marks_data:
        adm_str = str(m.get("adm_no", "")).strip()
        student_info = students_dict.get(adm_str, {})
        s_name = student_info.get("name", f"Student {adm_str}")
        assess_num = student_info.get("assessment_no", "")
        gender_val = student_info.get("gender", "")
        
        tot_marks, tot_points, valid_cnt = process_student_scores(m, analysis_grade)
        if valid_cnt > 0:
            mean_pts = round(tot_points / len(current_grade_learning_areas), 1)
            active_ranked_students.append({
                "adm_no": adm_str,
                "assessment_no": assess_num,
                "name": s_name,
                "gender": gender_val,
                "total_marks": tot_marks,
                "total_points": tot_points,
                "mean_points": mean_pts,
                "marks_row": m
            })

    # Overall Rank based on total points / total marks for active learners
    active_ranked_students = sorted(active_ranked_students, key=lambda x: (x["total_points"], x["total_marks"]), reverse=True)
    for idx, item in enumerate(active_ranked_students):
        item["rank"] = idx + 1

    # =========================================================
    # TAB 1: GENERAL PERFORMANCE OVERVIEW
    # =========================================================
    with tab_overview:
        if not active_ranked_students:
            st.info(f"No active student records with submitted marks found in {analysis_grade}.")
        else:
            st.markdown("### 🏆 Top Ten Students")
            top_10_rows = []
            for item in active_ranked_students[:10]:
                top_10_rows.append({
                    "Rank": item["rank"],
                    "Adm No": item["adm_no"],
                    "Student Name": item["name"],
                    "Total Marks": item["total_marks"],
                    "Total Points": item["total_points"],
                    "Mean Points": item["mean_points"],
                    "Overall Grade": calculate_total_grade(item["total_points"], analysis_grade)
                })
            st.dataframe(pd.DataFrame(top_10_rows), use_container_width=True)
            
            st.markdown("---")
            subject_stats = []
            df_marks = pd.DataFrame(marks_data)
            
            for sub in current_grade_learning_areas:
                col_name = get_db_col_name(sub)
                if col_name in df_marks.columns:
                    valid_scores = [float(v) for v in df_marks[col_name] if v is not None and not pd.isna(v) and float(v) >= 0.0]
                    valid_count = len(valid_scores)
                    if valid_count > 0:
                        mean_score = sum(valid_scores) / valid_count
                        mean_pts = sum([get_subject_performance(s, analysis_grade)[2] for s in valid_scores]) / valid_count
                    else:
                        mean_score, mean_pts = 0.0, 0.0

                    subject_stats.append({
                        "Learning Area": sub,
                        "Valid Submissions": valid_count,
                        "Mean Score (%)": round(mean_score, 2),
                        "Mean Points": round(mean_pts, 2)
                    })
            st.markdown("### 📊 Learning Area Performance Statistics")
            st.dataframe(pd.DataFrame(subject_stats), use_container_width=True)

    # =========================================================
    # TAB 2: WHOLE CLASS RANKING
    # =========================================================
    with tab_ranking:
        st.markdown(f"### 📋 Whole Class Ranking & Roll — {analysis_grade}")
        if not active_ranked_students:
            st.info(f"No student records found for {analysis_grade}.")
        else:
            ranking_view = st.radio(
                "Select Ranking View Phase:",
                ["Concise Summary", "Full Subject Marking Matrix"],
                horizontal=True
            )
            st.markdown("---")
            if "Concise" in ranking_view:
                rows = []
                for item in active_ranked_students:
                    rows.append({
                        "Rank": item["rank"],
                        "Adm No": item["adm_no"],
                        "Student Name": item["name"],
                        "Total Marks": item["total_marks"],
                        "Total Points": item["total_points"],
                        "Mean Points": item["mean_points"],
                        "Performance Level": calculate_total_grade(item["total_points"], analysis_grade)
                    })
                df_res = pd.DataFrame(rows)
                st.dataframe(df_res, use_container_width=True)
                st.download_button("📥 Download Concise Roll (.csv)", df_res.to_csv(index=False).encode('utf-8'), f"{analysis_grade}_Concise_Roll.csv", "text/csv")
            else:
                rows = []
                for item in active_ranked_students:
                    s_row = item["marks_row"]
                    row_dict = {"Rank": item["rank"], "Student Name": item["name"]}
                    for sub in current_grade_learning_areas:
                        abbr = SUBJECT_ABBR_MAP.get(sub, sub)
                        db_k = get_db_col_name(sub)
                        val = s_row.get(db_k)
                        row_dict[abbr] = f"{float(val):.0f}" if val is not None and not pd.isna(val) and float(val) >= 0 else "-"
                    row_dict["Total Marks"] = item["total_marks"]
                    row_dict["Total Points"] = item["total_points"]
                    row_dict["Mean Points"] = item["mean_points"]
                    row_dict["Overall Grade"] = calculate_total_grade(item["total_points"], analysis_grade)
                    rows.append(row_dict)
                df_res = pd.DataFrame(rows)
                st.dataframe(df_res, use_container_width=True)
                st.download_button("📥 Download Full Matrix Roll (.csv)", df_res.to_csv(index=False).encode('utf-8'), f"{analysis_grade}_Matrix_Roll.csv", "text/csv")

    # =========================================================
    # TAB 3: BEST THREE PER LEARNING AREA
    # =========================================================
    with tab_top_subjects:
        st.markdown(f"### 🥇 Best Three Students Per Learning Area — {analysis_grade}")
        if not active_ranked_students:
            st.info(f"No student records found for {analysis_grade}.")
        else:
            for sub in current_grade_learning_areas:
                col_k = get_db_col_name(sub)
                valid_sub_students = []
                for item in active_ranked_students:
                    val = item["marks_row"].get(col_k)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        code, perf, pts = get_subject_performance(float(val), analysis_grade)
                        valid_sub_students.append({
                            "Student Name": item["name"],
                            "Adm No": item["adm_no"],
                            "Score (%)": float(val),
                            "Perf. Level": code,
                            "Points": pts
                        })
                valid_sub_students = sorted(valid_sub_students, key=lambda x: x["Score (%)"], reverse=True)[:3]
                with st.expander(f"📚 {sub}", expanded=True):
                    if valid_sub_students:
                        sub_top_rows = []
                        for pos_idx, top_s in enumerate(valid_sub_students, 1):
                            medal = f"Position {pos_idx}"
                            sub_top_rows.append({
                                "Position": medal,
                                "Adm No": top_s["Adm No"],
                                "Student Name": top_s["Student Name"],
                                "Score (%)": top_s["Score (%)"],
                                "Perf. Level": top_s["Perf. Level"],
                                "Points": top_s["Points"]
                            })
                        st.dataframe(pd.DataFrame(sub_top_rows), use_container_width=True, hide_index=True)
                    else:
                        st.caption(f"No submitted scores recorded yet for {sub}.")

    # =========================================================
    # TAB 4: ASSESSMENT REPORT FORMS HUB
    # =========================================================
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
            # Determine template file based on section
            g_lower = analysis_grade.strip().lower()
            if any(g in g_lower for g in ["grade 1", "grade 2", "grade 3"]):
                template_filename = "report3.docx"
            elif any(g in g_lower for g in ["grade 4", "grade 5", "grade 6"]):
                template_filename = "report2.docx"
            else:
                template_filename = "report1.docx"

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

                if os.path.exists(template_filename):
                    doc = Document(template_filename)
                elif os.path.exists("report.docx"):
                    doc = Document("report.docx")
                else:
                    doc = Document()

                for section in doc.sections:
                    section.top_margin = Inches(0.3)
                    section.bottom_margin = Inches(0.3)
                    section.left_margin = Inches(0.4)
                    section.right_margin = Inches(0.4)

                if os.path.exists("logo.png"):
                    p_logo = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run_logo = p_logo.add_run()
                    run_logo.add_picture("logo.png", width=Inches(0.9))

                # Subject-wise calculation and ranking
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
                        code, perf, pts = get_subject_performance(s_num, analysis_grade)
                        
                        # Calculate position per learning area among active students
                        sub_scores_list = []
                        for st_item in active_ranked_students:
                            v_sub = st_item["marks_row"].get(col_key)
                            if v_sub is not None and not pd.isna(v_sub) and float(v_sub) >= 0.0:
                                sub_scores_list.append((st_item["adm_no"], float(v_sub)))
                        sub_scores_list = sorted(sub_scores_list, key=lambda x: x[1], reverse=True)
                        sub_rank = next((i+1 for i, (a, _) in enumerate(sub_scores_list) if a == adm_str), 1)
                        sub_total_active = len(sub_scores_list)
                        pos_str = f"{sub_rank}/{sub_total_active}"
                        
                        subject_details.append({
                            "name": sub.upper(),
                            "marks": f"{s_num:.0f}",
                            "level": code,
                            "points": str(pts),
                            "position": pos_str,
                            "teacher": "TBA"
                        })
                    else:
                        subject_details.append({
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
                dos_comment = generate_dos_comment(total_points, analysis_grade) if any(g in analysis_grade.lower() for g in ["grade 7", "grade 8", "grade 9"]) else ""

                # Populate doc tables and paragraphs
                for t in doc.tables:
                    t.alignment = WD_TABLE_ALIGNMENT.CENTER
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
                            elif "YEAR:" in cu: cell.text = f"YEAR: 2027"

                        # Subject table row filling if table matches learning areas
                        for sd in subject_details:
                            if sd["name"] in row_text and len(row.cells) >= 6:
                                row.cells[1].text = sd["marks"]
                                row.cells[2].text = sd["level"]
                                row.cells[3].text = sd["points"]
                                row.cells[4].text = sd["position"]
                                row.cells[5].text = sd["teacher"]

                replacements = {
                    "TOTAL MARKS": f"TOTAL MARKS: {total_marks:.0f}",
                    "TOTAL POINTS": f"TOTAL POINTS: {total_points}",
                    "MEAN POINTS": f"MEAN POINTS: {mean_points}",
                    "GENERAL PERFORMANCE": f"GENERAL PERFORMANCE LEVEL: {overall_g}",
                    "TEACHER": f"CLASS TEACHER'S COMMENT: {teacher_comment}",
                    "HOI": f"HEAD OF INSTITUTION COMMENT: {hoi_comment}",
                    "DOS": f"DIRECTOR OF STUDIES COMMENT: {dos_comment}",
                    "CLOSING DATE": f"CLOSING DATE: {closing_date}",
                    "OPENING DATE": f"OPENING DATE: {opening_date}",
                }

                for p in doc.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for k, v in replacements.items():
                        if k in p.text.upper():
                            p.text = v

                if stamp_upload or os.path.exists("stamp.png"):
                    stamp_source = stamp_upload if stamp_upload else "stamp.png"
                    for p in doc.paragraphs:
                        if "STAMP" in p.text.upper() or "SIGN" in p.text.upper():
                            run = p.add_run()
                            run.add_picture(stamp_source, width=Inches(1.1))
                            break

                return doc

            st.subheader("1. Live Student Report Form Preview")
            selected_student_obj = st.selectbox(
                "Select Student for Live Report Form Preview", 
                active_ranked_students, 
                format_func=lambda x: f"Rank #{x['rank']} | Adm No: {x['adm_no']} | Name: {x['name']} (Points: {x['total_points']})"
            )

            if selected_student_obj:
                doc_obj = generate_student_docx_object(selected_student_obj)
                out_stream = io.BytesIO()
                doc_obj.save(out_stream)
                out_stream.seek(0)
                
                st.download_button(
                    label=f"📥 Download Assessment Report for {selected_student_obj['name']} (.docx)",
                    data=out_stream,
                    file_name=f"Rank_{selected_student_obj['rank']}_{selected_student_obj['adm_no']}_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

            st.markdown("---")
            st.subheader("2. Ranked Batch Reports Generator (Entire Grade)")
            if st.button(f"📦 Generate All Ranked Reports for {analysis_grade} (.zip)", type="secondary"):
                with st.spinner("Compiling ranked report forms... Please wait."):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for s_item in active_ranked_students:
                            single_doc = generate_student_docx_object(s_item)
                            doc_bytes = io.BytesIO()
                            single_doc.save(doc_bytes)
                            doc_bytes.seek(0)
                            file_filename = f"Rank_{s_item['rank']:02d}_{s_item['adm_no']}_{s_item['name'].replace(' ', '_')}_Report.docx"
                            zip_file.writestr(file_filename, doc_bytes.getvalue())

                    zip_buffer.seek(0)
                    st.success(f"Successfully compiled {len(active_ranked_students)} ranked report forms!")
                    st.download_button(
                        label=f"📥 Download All {len(active_ranked_students)} Ranked Reports (.zip)",
                        data=zip_buffer,
                        file_name=f"{analysis_grade.replace(' ', '_')}_Ranked_Assessment_Reports.zip",
                        mime="application/zip"
                    )

# --- PAGE 5: TEACHERS PORTAL ---
elif page == "Teachers Portal":
    st.header(f"Teacher Portal — {st.session_state.full_name}")
    user_key = st.session_state.username
    teacher_info = TEACHER_ASSIGNMENTS.get(user_key)
    
    if is_admin_role and not teacher_info:
        st.markdown("### Administrator Overview")
        st.info("You are logged in with Administrative privileges.")
        try:
            t_res = supabase.table("teachers").select("full_name", "username", "designation").execute()
            st.dataframe(pd.DataFrame(t_res.data), use_container_width=True)
        except Exception as e:
            st.error(f"Could not load teachers directory: {e}")
    elif teacher_info:
        st.markdown("### Your Assigned Learning Areas & Responsibilities")
        for item in teacher_info["assignments"]:
            grade = item["grade"]
            subs = ", ".join(item["subjects"])
            st.success(f"📌 **{grade}:** Teaching **{subs}**")
            try:
                st_res = supabase.table("students").select("adm_no", "assessment_no", "name", "gender").eq("grade", grade).execute()
                if st_res.data:
                    st.markdown(f"**Registered Students in {grade}:**")
                    st.dataframe(pd.DataFrame(st_res.data), use_container_width=True)
            except Exception as e:
                st.error(f"Could not load students for {grade}: {e}")
            st.markdown("---")
    else:
        st.warning("No teaching assignments found for your account.")

# --- PAGE 6: NEWSLETTER ---
elif page == "Newsletter":
    st.header("School Newsletter & Memos (2027)")
    
    if is_admin_role:
        with st.form("newsletter_form"):
            news_title = st.text_input("Newsletter Title / Memo Subject")
            news_content = st.text_area("News Content & Announcements")
            submit_news = st.form_submit_button("Publish Newsletter")
            if submit_news:
                try:
                    supabase.table("newsletter").insert({
                        "title": news_title,
                        "content": news_content,
                        "date": datetime.now().strftime("%Y-%m-%d")
                    }).execute()
                    st.success("Newsletter published successfully!")
                except Exception as e:
                    st.error(f"Error publishing: {e}")

    try:
        n_res = supabase.table("newsletter").select("*").order("date", desc=True).execute()
        if n_res.data:
            for item in n_res.data:
                st.markdown(f"### {item['title']}")
                st.markdown(f"*{item['date']}*")
                st.write(item['content'])
                
                # Downloadable / Printable text format
                memo_bytes = f"KEA COMPREHENSIVE SCHOOL MEMO\nDate: {item['date']}\nTitle: {item['title']}\n\n{item['content']}".encode('utf-8')
                st.download_button(
                    label=f"📥 Download / Print Memo: {item['title']}",
                    data=memo_bytes,
                    file_name=f"Memo_{item['date']}.txt",
                    mime="text/plain",
                    key=f"memo_dl_{item['title']}"
                )
                st.markdown("---")
        else:
            st.info("No newsletters posted yet.")
    except Exception as e:
        st.info("No newsletters posted yet.")