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
        "assignments": [
            {"grade": "Grade 8", "subjects": ["INTEGRATED SCIENCE", "AGRICULTURE"]},
            {"grade": "Grade 9", "subjects": ["MATHEMATICS", "AGRICULTURE"]}
        ]
    },
    "Vincent": {
        "assignments": [
            {"grade": "Grade 7", "subjects": ["SOCIAL STUDIES"]},
            {"grade": "Grade 8", "subjects": ["SOCIAL STUDIES", "KISWAHILI"]},
            {"grade": "Grade 9", "subjects": ["SOCIAL STUDIES", "KISWAHILI"]}
        ]
    },
    "Grace": {
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

LEARNING_AREAS = [
    "MATHEMATICS", "ENGLISH", "KISWAHILI", "INTEGRATED SCIENCE", 
    "AGRICULTURE", "PRETECHNICAL STUDIES", "SOCIAL STUDIES", 
    "CHRISTIAN RELIGIOUS EDUCATION", "CREATIVE ARTS SPORTS"
]

# Abbreviation map matching the column layout requirements
SUBJECT_ABBR_MAP = {
    "MATHEMATICS": "MATHS",
    "ENGLISH": "ENG",
    "KISWAHILI": "KIS",
    "INTEGRATED SCIENCE": "INT SCI",
    "AGRICULTURE": "AGRIC",
    "PRETECHNICAL STUDIES": "PRETECH",
    "CHRISTIAN RELIGIOUS EDUCATION": "CRE",
    "SOCIAL STUDIES": "SOCIAL",
    "CREATIVE ARTS SPORTS": "CAS"
}

# --- SUBJECT TO POSTGRESQL COLUMN MAPPING HELPER ---
SUBJECT_COLUMN_MAP = {
    "CREATIVE ARTS SPORTS": "creative_arts_sports",
    "CHRISTIAN RELIGIOUS EDUCATION": "christian_religious_education",
    "PRETECHNICAL STUDIES": "pretechnical_studies",
    "INTEGRATED SCIENCE": "integrated_science",
    "SOCIAL STUDIES": "social_studies"
}

def get_db_col_name(subject_name):
    subject_clean = subject_name.strip().upper()
    if subject_clean in SUBJECT_COLUMN_MAP:
        return SUBJECT_COLUMN_MAP[subject_clean]
    return subject_name.lower().strip().replace(" ", "_").replace(".", "").replace("(", "").replace(")", "")

# --- AUTHENTICATION & LOGIN SCREEN ---
def login_screen():
    st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>KEA COMPREHENSIVE SCHOOL</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #4B5563;'>School Management & Information System</h3>", unsafe_allow_html=True)
    
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
        "Fee Payment", 
        "Teachers Portal", 
        "Teacher Time Login", 
        "Newsletter", 
        "School Contact"
    ]
else:
    nav_options = [
        "Dashboard", 
        "Marks Entry", 
        "Newsletter", 
        "School Contact"
    ]

page = st.sidebar.selectbox("Go to", nav_options)

# --- HELPER FUNCTION FOR SUBJECT GRADING ---
def get_subject_performance(score):
    if score is None or pd.isna(score) or float(score) < 0.0:
        return "-", "-", None
    score = float(score)
    if score <= 10: return "BE2", "Below Expectation 2", 1
    elif score <= 20: return "BE1", "Below Expectation 1", 2
    elif score <= 30: return "AE2", "Approaching Expectation 2", 3
    elif score <= 40: return "AE1", "Approaching Expectation 1", 4
    elif score <= 56: return "ME2", "Meeting Expectation 2", 5
    elif score <= 73: return "ME1", "Meeting Expectation 1", 6
    elif score <= 88: return "EE2", "Exceeding Expectation 2", 7
    else: return "EE1", "Exceeding Expectation 1", 8

def calculate_total_grade(total_marks):
    if total_marks <= 112: return "BE2 (Below Expectation 2)"
    elif total_marks <= 224: return "BE1 (Below Expectation 1)"
    elif total_marks <= 336: return "AE2 (Approaching Expectation 2)"
    elif total_marks <= 449: return "AE1 (Approaching Expectation 1)"
    elif total_marks <= 560: return "ME2 (Meeting Expectation 2)"
    elif total_marks <= 672: return "ME1 (Meeting Expectation 1)"
    elif total_marks <= 785: return "EE2 (Exceeding Expectation 2)"
    else: return "EE1 (Exceeding Expectation 1)"

# --- HELPER FUNCTION: FORMULATE TEACHER'S COMMENT ---
def generate_teacher_comment(total_score, top_subject, low_subject):
    overall_grade = calculate_total_grade(total_score)
    if "EE" in overall_grade:
        return f"An excellent performance! Demonstrates outstanding mastery across learning areas, especially in {top_subject}. Keep up the high standard!"
    elif "ME" in overall_grade:
        return f"Good performance. Shows consistent effort and understanding, particularly in {top_subject}. Aim to put more effort into {low_subject} for better results."
    elif "AE" in overall_grade:
        return f"Fair performance. Shows potential in {top_subject}, but requires more focus and practice in {low_subject} to improve overall standing."
    else:
        return f"Below expectations. Needs active academic support and dedicated revision, particularly in {low_subject}. Steady practice in {top_subject} will build confidence."

# --- PAGE 1: DASHBOARD ---
if page == "Dashboard":
    # CSS Custom Styling for Modern Dark EduAdmin Layout UI
    st.markdown("""
        <style>
        .dashboard-header {
            background-color: #111827;
            padding: 15px 25px;
            border-radius: 12px;
            color: white;
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
        }
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

    # Fetch dynamic backend statistics
    total_students = 0
    total_teachers = 0
    total_collected = 0.0
    expected_fee_per_student = 300.0
    total_expected = 0.0
    deficit = 0.0

    if supabase:
        try:
            students_res = supabase.table("students").select("*", count="exact").execute()
            teachers_res = supabase.table("teachers").select("*", count="exact").execute()
            fees_res = supabase.table("fee_payments").select("amount_paid").execute()
            
            total_students = students_res.count if students_res.count is not None else len(students_res.data)
            total_teachers = teachers_res.count if teachers_res.count is not None else len(teachers_res.data)
            
            if fees_res.data:
                total_collected = sum(float(item.get("amount_paid", 0)) for item in fees_res.data)
                
            total_expected = total_students * expected_fee_per_student
            deficit = max(0.0, total_expected - total_collected)
        except Exception as e:
            st.error(f"Error fetching dashboard metrics: {e}")

    # Top Hero Banners
    b_col1, b_col2 = st.columns([3, 1])
    with b_col1:
        st.markdown(f"""
            <div class="banner-card">
                <h3>🎓 KEA COMPREHENSIVE SCHOOL — EduAdmin Dashboard</h3>
                <p>Welcome back, <b>{st.session_state.full_name}</b> ({st.session_state.role})! Access termly results analysis, registered student rosters, fee payments, and teaching schedules seamlessly.</p>
            </div>
        """, unsafe_allow_html=True)
    with b_col2:
        st.markdown("""
            <div class="support-card">
                <h4 style="margin:0;">📞 School Support</h4>
                <p style="margin:5px 0 0 0; font-size:16px; font-weight:bold;">+254 700 000 000</p>
                <small>Migori County, Kenya</small>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Key Statistics Metric Grid
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f"""
            <div class="metric-card">
                <small style="color:#94A3B8;">TOTAL REGISTERED STUDENTS</small>
                <h2 style="margin:5px 0; color:#60A5FA;">{total_students}</h2>
                <small style="color:#34D399;">Active Enrolment</small>
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
    with m3:
        st.markdown(f"""
            <div class="metric-card">
                <small style="color:#94A3B8;">FEE COLLECTED (KSH)</small>
                <h2 style="margin:5px 0; color:#34D399;">{total_collected:,.2f}</h2>
                <small style="color:#34D399;">Term Revenue Received</small>
            </div>
        """, unsafe_allow_html=True)
    with m4:
        st.markdown(f"""
            <div class="metric-card">
                <small style="color:#94A3B8;">FEE DEFICIT (KSH)</small>
                <h2 style="margin:5px 0; color:#F87171;">{deficit:,.2f}</h2>
                <small style="color:#F87171;">Pending Collections</small>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Middle Content: Course Classes & Performance Statistics
    c_left, c_right = st.columns([2, 1])

    with c_left:
        st.markdown("### 📚 Primary & Junior Secondary Classes")
        cls_col1, cls_col2, cls_col3 = st.columns(3)
        with cls_col1:
            st.markdown("""
                <div class="course-card">
                    <h4>📘 Lower Primary</h4>
                    <p style="color:#94A3B8; margin:0;">Grade 1 — Grade 3</p>
                    <small style="color:#60A5FA;">CBC Curriculum</small>
                </div>
            """, unsafe_allow_html=True)
        with cls_col2:
            st.markdown("""
                <div class="course-card">
                    <h4>📗 Upper Primary</h4>
                    <p style="color:#94A3B8; margin:0;">Grade 4 — Grade 6</p>
                    <small style="color:#60A5FA;">CBC Curriculum</small>
                </div>
            """, unsafe_allow_html=True)
        with cls_col3:
            st.markdown("""
                <div class="course-card">
                    <h4>📙 Junior Secondary</h4>
                    <p style="color:#94A3B8; margin:0;">Grade 7 — Grade 9</p>
                    <small style="color:#60A5FA;">CBC Curriculum</small>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("### 📊 Performance & Revenue Statistics")
        st.write("**Termly Financial Progress Tracker**")
        progress_val = (total_collected / total_expected) if total_expected > 0 else 0.0
        progress_val = min(1.0, max(0.0, float(progress_val)))
        st.progress(progress_val)
        st.info(f"Financial Progress: Ksh {total_collected:,.2f} collected out of expected Ksh {total_expected:,.2f} total revenue.")

    with c_right:
        st.markdown("### 🗓️ School Calendar & Notice")
        st.date_input("Term Calendar View", value=datetime.now())
        
        st.markdown("""
            <div class="course-card" style="border-left: 4px solid #F59E0B;">
                <h5 style="margin:0; color:#F59E0B;">📢 Important Reminder</h5>
                <p style="font-size:12px; color:#CBD5E1; margin-top:5px;">
                    Ensure all subject-specific exam scores for Term 3 are saved and submitted under <b>Marks Entry</b> before final ranking generation.
                </p>
            </div>
        """, unsafe_allow_html=True)

# --- PAGE 2: STUDENTS REGISTRATION ---
elif page == "Students Registration":
    if not is_admin_role:
        st.error("Access Denied: Only HOI, DHOI, and Senior Teachers can register students.")
        st.stop()
        
    st.header("Students Registration Portal")
    tab_manual, tab_excel, tab_manage = st.tabs(["Manual Registration", "Excel Spreadsheet Upload", "Manage Registered Students"])
    
    with tab_manual:
        with st.form("manual_student_form"):
            adm_no = st.text_input("Admission Number")
            assessment_no = st.text_input("Assessment Number")
            name = st.text_input("Full Name")
            grade = st.selectbox("Grade", [f"Grade {i}" for i in range(1, 10)])
            submit_manual = st.form_submit_button("Admit Student")
            
            if submit_manual:
                if adm_no and name and grade:
                    try:
                        supabase.table("students").insert({
                            "adm_no": adm_no,
                            "assessment_no": assessment_no,
                            "name": name,
                            "grade": grade
                        }).execute()
                        st.success(f"Successfully admitted {name} to {grade}!")
                    except Exception as e:
                        st.error(f"Error saving student: {e}")
                else:
                    st.warning("Please fill in all mandatory fields.")

    with tab_excel:
        st.markdown("Upload an Excel spreadsheet with columns: **Adm No**, **Assessment No**, **Name**, **Grade**")
        uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx", "xls"])
        selected_grade_bulk = st.selectbox("Target Grade for Bulk Upload", [f"Grade {i}" for i in range(1, 10)])
        
        if uploaded_file and st.button("Process & Admit Students"):
            try:
                df = pd.read_excel(uploaded_file)
                df.columns = [str(c).strip().lower() for c in df.columns]
                
                adm_col = next((c for c in df.columns if "adm" in c), None)
                name_col = next((c for c in df.columns if "name" in c), None)
                assess_col = next((c for c in df.columns if "assess" in c), None)
                grade_col = next((c for c in df.columns if "grade" in c), None)
                
                if not adm_col or not name_col:
                    st.error("Excel file must contain at least 'Adm No' and 'Name' columns.")
                    st.stop()

                records = []
                for _, row in df.iterrows():
                    adm_val = str(row[adm_col]).strip() if pd.notna(row[adm_col]) else None
                    name_val = str(row[name_col]).strip() if pd.notna(row[name_col]) else None
                    assess_val = str(row[assess_col]).strip() if assess_col and pd.notna(row[assess_col]) else ""
                    grade_val = str(row[grade_col]).strip() if grade_col and pd.notna(row[grade_col]) else selected_grade_bulk
                    
                    if adm_val and adm_val.lower() != "nan" and name_val and name_val.lower() != "nan":
                        records.append({
                            "adm_no": adm_val,
                            "assessment_no": assess_val,
                            "name": name_val,
                            "grade": grade_val
                        })

                if not records:
                    st.warning("No valid student records found in the uploaded file.")
                else:
                    unique_records_dict = {r["adm_no"]: r for r in records}
                    deduplicated_records = list(unique_records_dict.values())
                    
                    duplicates_removed = len(records) - len(deduplicated_records)
                    if duplicates_removed > 0:
                        st.info(f"Removed {duplicates_removed} duplicate admission number entry/entries from batch.")

                    supabase.table("students").upsert(deduplicated_records, on_conflict="adm_no").execute()
                    st.success(f"Successfully uploaded and admitted {len(deduplicated_records)} students!")

            except Exception as e:
                st.error(f"Error processing excel file: {e}")

    with tab_manage:
        st.markdown("### View, Edit & Deregister Students")
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
                cols = [c for c in ["adm_no", "assessment_no", "name", "grade"] if c in df_students.columns]
                df_students = df_students[cols]
                
                st.markdown("#### Registered Students Roster")
                st.caption("Click on any cell to edit details directly, then click 'Save Changes' below.")
                
                edited_students_df = st.data_editor(
                    df_students,
                    column_config={
                        "adm_no": st.column_config.TextColumn("Adm No", disabled=True),
                        "assessment_no": st.column_config.TextColumn("Assessment No"),
                        "name": st.column_config.TextColumn("Full Name"),
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
                st.markdown("#### Deregister / Remove Student")
                
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

        except Exception as e:
            st.error(f"Error loading students: {e}")

# --- PAGE 3: MARKS ENTRY ---
elif page == "Marks Entry":
    st.header("Subject-Specific Marks Entry Portal")
    
    user_key = st.session_state.username
    teacher_data = TEACHER_ASSIGNMENTS.get(user_key)
    
    if is_admin_role or not teacher_data:
        allowed_grades = [f"Grade {i}" for i in range(1, 10)]
        teacher_grade_map = {g: LEARNING_AREAS for g in allowed_grades}
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
    
    assigned_subjects = teacher_grade_map.get(target_grade, LEARNING_AREAS)
    
    with col_s:
        target_subject = st.selectbox("2. Select Learning Area", assigned_subjects)

    st.markdown("---")
    st.subheader(f"Marking Sheet: {target_grade} — {target_subject}")
    st.caption("Enter scores out of 100%. Blank cells remain null (actual 0 scores are supported). Unrecorded students appear for entry, while successfully saved entries move to the 'Recorded' table.")

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

        st.markdown("### 🔍 Quick Student Search")
        search_query = st.selectbox(
            "Search student by Admission No or Name to jump/filter:",
            options=["-- All Students --"] + [f"{s['adm_no']} - {s['name']}" for s in students],
            key=f"search_bar_{target_grade}_{col_key}"
        )
        
        filtered_unrecorded = unrecorded_students
        if search_query and search_query != "-- All Students --":
            selected_adm = search_query.split(" - ")[0]
            filtered_unrecorded = [s for s in unrecorded_students if s["adm_no"] == selected_adm]
            if not filtered_unrecorded:
                st.info(f"Student {search_query} is already recorded in the database below.")

        st.markdown("---")

        tab_manual_marks, tab_excel_marks = st.tabs(["Manual Sheet Input", "Upload Excel Marks"])

        with tab_manual_marks:
            st.markdown("### 📝 Active Marking Table (Unrecorded Students)")
            
            if not unrecorded_students:
                st.success("🎉 All students in this grade have their marks recorded for this learning area!")
            else:
                session_marks_key = f"input_scores_{target_grade}_{col_key}"
                if session_marks_key not in st.session_state:
                    st.session_state[session_marks_key] = {}

                with st.form(key=f"unrecorded_form_{target_grade}_{col_key}"):
                    editor_rows = []
                    for s in filtered_unrecorded:
                        adm = s["adm_no"]
                        default_val = st.session_state[session_marks_key].get(adm, None)
                        
                        editor_rows.append({
                            "Adm No": adm,
                            "Student Name": s["name"],
                            "Score (%)": float(default_val) if default_val is not None and not pd.isna(default_val) else None,
                            "Save & Submit": False
                        })

                    df_unrecorded = pd.DataFrame(editor_rows)

                    edited_df = st.data_editor(
                        df_unrecorded,
                        column_config={
                            "Adm No": st.column_config.TextColumn("Adm No", disabled=True),
                            "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                            "Score (%)": st.column_config.NumberColumn(
                                "Score (%)", 
                                min_value=0.0, 
                                max_value=100.0, 
                                step=1.0,
                                format="%.0f",
                                help="Leave blank for null. 0 is fully supported."
                            ),
                            "Save & Submit": st.column_config.CheckboxColumn(
                                "Save & Submit",
                                help="Check to submit individual student mark",
                                default=False
                            )
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
                        
                        for _, row in edited_df.iterrows():
                            adm = str(row["Adm No"])
                            raw_score = row["Score (%)"]
                            is_checked = row["Save & Submit"]
                            
                            should_save = False
                            if submit_all and raw_score is not None and not pd.isna(raw_score):
                                should_save = True
                            elif submit_checked and is_checked and raw_score is not None and not pd.isna(raw_score):
                                should_save = True

                            if should_save:
                                existing = existing_full_rows.get(adm, {})
                                rec = dict(existing) if existing else {"adm_no": adm, "grade": target_grade}
                                
                                rec[col_key] = float(raw_score) if raw_score is not None and not pd.isna(raw_score) else None
                                
                                total = 0.0
                                for sub in LEARNING_AREAS:
                                    s_key = get_db_col_name(sub)
                                    val = rec.get(s_key)
                                    if val is not None and not pd.isna(val) and float(val) >= 0:
                                        total += float(val)
                                
                                rec["total_marks"] = total
                                records_to_upsert.append(rec)
                                submitted_count += 1

                        if records_to_upsert:
                            try:
                                supabase.table("marks").upsert(records_to_upsert, on_conflict="adm_no").execute()
                                st.success(f"Successfully saved and submitted marks for {submitted_count} student(s)! Moved to Recorded table.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Database error saving marks: {e}")
                        else:
                            st.warning("Please enter valid scores and check/select items to submit.")

            st.markdown("---")

            st.markdown("### ✅ Recorded Students Table")
            if not recorded_students:
                st.info("No marks recorded yet for this learning area.")
            else:
                recorded_rows = []
                for s in recorded_students:
                    adm = s["adm_no"]
                    score_val = existing_marks_map.get(adm)
                    grade_code, perf_level, _ = get_subject_performance(score_val)
                    
                    recorded_rows.append({
                        "Adm No": adm,
                        "Student Name": s["name"],
                        "Score (%)": score_val if score_val is not None else "-",
                        "Grade Code": grade_code,
                        "Performance Level": perf_level
                    })
                
                df_recorded = pd.DataFrame(recorded_rows)
                st.dataframe(df_recorded, use_container_width=True, hide_index=True)
                
                st.caption("Need to update a recorded mark? Select student below to clear and re-enter:")
                col_ed1, col_ed2 = st.columns([2, 1])
                with col_ed1:
                    student_to_reset = st.selectbox(
                        "Select Recorded Student to Edit",
                        recorded_students,
                        format_func=lambda x: f"{x['adm_no']} - {x['name']}",
                        key=f"reset_select_{target_grade}_{col_key}"
                    )
                with col_ed2:
                    st.write("")
                    st.write("")
                    if st.button("🔄 Move Back to Unrecorded", type="secondary"):
                        try:
                            adm_r = student_to_reset["adm_no"]
                            existing = existing_full_rows.get(adm_r, {})
                            if existing:
                                existing[col_key] = None
                                total = 0.0
                                for sub in LEARNING_AREAS:
                                    s_key = get_db_col_name(sub)
                                    val = existing.get(s_key)
                                    if val is not None and not pd.isna(val) and float(val) >= 0:
                                        total += float(val)
                                existing["total_marks"] = total
                                supabase.table("marks").upsert(existing, on_conflict="adm_no").execute()
                                st.success(f"Moved {student_to_reset['name']} back to unrecorded marking list.")
                                st.rerun()
                        except Exception as e:
                            st.error(f"Error resetting mark: {e}")

        with tab_excel_marks:
            st.markdown(f"Upload an Excel sheet for **{target_grade} — {target_subject}**.")
            st.caption("Ensure the file contains an **'Adm No'** column and a column named **'Score'** or matching the subject.")
            
            marks_file = st.file_uploader("Choose Excel File", type=["xlsx", "xls"], key=f"uploader_{target_grade}_{col_key}")
            
            if marks_file and st.button("Upload Excel Marks Sheet"):
                try:
                    df_upload = pd.read_excel(marks_file)
                    records_to_upsert = []
                    
                    score_col = None
                    for c in df_upload.columns:
                        if c.lower() in ["score", "marks", target_subject.lower()]:
                            score_col = c
                            break
                    
                    if not score_col and len(df_upload.columns) >= 2:
                        score_col = df_upload.columns[1]

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
                        
                        total = 0.0
                        for sub in LEARNING_AREAS:
                            s_key = get_db_col_name(sub)
                            val = rec.get(s_key)
                            if val is not None and not pd.isna(val) and float(val) >= 0:
                                total += float(val)
                        rec["total_marks"] = total
                        records_to_upsert.append(rec)

                    supabase.table("marks").upsert(records_to_upsert, on_conflict="adm_no").execute()
                    st.success(f"Successfully processed and uploaded marks for {len(records_to_upsert)} students!")
                except Exception as e:
                    st.error(f"Error processing uploaded file: {e}")

# --- PAGE 4: RESULTS ANALYSIS & REPORT FORMS HUB ---
elif page == "Results Analysis":
    st.header("Results Analysis Hub & Report Forms")
    
    analysis_grade = st.selectbox("Select Grade for Analysis", [f"Grade {i}" for i in range(1, 10)], key="analysis_grade_select")
    
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

    def process_student_scores(m_row):
        total = 0.0
        valid_count = 0
        for sub in LEARNING_AREAS:
            k = get_db_col_name(sub)
            val = m_row.get(k)
            if val is not None and not pd.isna(val) and float(val) >= 0.0:
                total += float(val)
                valid_count += 1
        return total, valid_count

    active_ranked_students = []
    for m in marks_data:
        adm_str = str(m.get("adm_no", "")).strip()
        assess_num = students_dict.get(adm_str, {}).get("assessment_no", "")
        s_name = students_dict.get(adm_str, {}).get("name", f"Student {adm_str}")
        tot_marks, valid_cnt = process_student_scores(m)
        
        if valid_cnt > 0:
            active_ranked_students.append({
                "adm_no": adm_str,
                "assessment_no": assess_num,
                "name": s_name,
                "total_marks": tot_marks,
                "marks_row": m
            })

    active_ranked_students = sorted(active_ranked_students, key=lambda x: x["total_marks"], reverse=True)
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
                    "Overall Grade": calculate_total_grade(item["total_marks"])
                })
            st.dataframe(pd.DataFrame(top_10_rows), use_container_width=True)
            
            st.markdown("---")

            subject_stats = []
            df_marks = pd.DataFrame(marks_data)
            
            for sub in LEARNING_AREAS:
                col_name = get_db_col_name(sub)
                if col_name in df_marks.columns:
                    valid_scores = [float(v) for v in df_marks[col_name] if v is not None and not pd.isna(v) and float(v) >= 0.0]
                    valid_count = len(valid_scores)
                    
                    if valid_count > 0:
                        total_sub_marks = sum(valid_scores)
                        mean_score = total_sub_marks / valid_count
                        total_points = sum([get_subject_performance(s)[2] for s in valid_scores if get_subject_performance(s)[2] is not None])
                        mean_points = total_points / valid_count
                    else:
                        total_sub_marks, mean_score, mean_points = 0.0, 0.0, 0.0

                    subject_stats.append({
                        "Learning Area": sub,
                        "Valid Students": valid_count,
                        "Total Marks": round(total_sub_marks, 1),
                        "Mean Score (%)": round(mean_score, 2),
                        "Aggregate Mean Points": round(mean_points, 2)
                    })

            df_subject_stats = pd.DataFrame(subject_stats)

            if not df_subject_stats.empty and df_subject_stats["Mean Score (%)"].max() > 0:
                best_subject = df_subject_stats.sort_values(by="Mean Score (%)", ascending=False).iloc[0]
                st.markdown(f"### ⭐ Best Performed Learning Area")
                st.success(
                    f"**{best_subject['Learning Area']}** with a Mean Score of **{best_subject['Mean Score (%)']}%** "
                    f"and Aggregate Points of **{best_subject['Aggregate Mean Points']}** (based on {best_subject['Valid Students']} valid submissions)."
                )
            
            st.markdown("---")
            st.markdown("### 📊 Performance Summary for Each Learning Area")
            st.dataframe(df_subject_stats, use_container_width=True)

    # =========================================================
    # TAB 2: WHOLE CLASS RANKING (TWO PHASES)
    # =========================================================
    with tab_ranking:
        st.markdown(f"### 📋 Whole Class Ranking & Roll — {analysis_grade}")
        if not active_ranked_students:
            st.info(f"No student records found for {analysis_grade}.")
        else:
            ranking_view = st.radio(
                "Select Ranking View Phase:",
                [
                    "Phase 1: Concise Rank Summary (Rank, Name, Total Score & Performance Level)", 
                    "Phase 2: Complete Subject Marking Roll (Subject Abbreviations Matrix)"
                ],
                horizontal=True
            )
            
            st.markdown("---")

            if "Phase 1" in ranking_view:
                st.markdown("#### Phase 1: Concise Class Ranking Summary")
                phase1_rows = []
                for item in active_ranked_students:
                    phase1_rows.append({
                        "Rank": item["rank"],
                        "Students Name": item["name"],
                        "Total Score": round(item["total_marks"], 1),
                        "Performance Level": calculate_total_grade(item["total_marks"])
                    })
                
                df_phase1 = pd.DataFrame(phase1_rows)
                st.dataframe(df_phase1, use_container_width=True)

                csv_p1 = df_phase1.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Download Phase 1 Concise Roll for {analysis_grade} (.csv)",
                    data=csv_p1,
                    file_name=f"{analysis_grade.replace(' ', '_')}_Concise_Class_Roll.csv",
                    mime="text/csv",
                    type="primary"
                )

            else:
                st.markdown("#### Phase 2: Full Subject Marking Roll Matrix")
                phase2_rows = []
                for item in active_ranked_students:
                    s_row = item["marks_row"]
                    row_dict = {
                        "Rank": item["rank"],
                        "Students Name": item["name"]
                    }
                    
                    for sub in LEARNING_AREAS:
                        abbr_col = SUBJECT_ABBR_MAP.get(sub, sub)
                        db_key = get_db_col_name(sub)
                        score_val = s_row.get(db_key)
                        
                        if score_val is not None and not pd.isna(score_val) and float(score_val) >= 0.0:
                            row_dict[abbr_col] = f"{float(score_val):.0f}"
                        else:
                            row_dict[abbr_col] = "-"
                    
                    row_dict["Total score"] = round(item["total_marks"], 1)
                    row_dict["General performance level"] = calculate_total_grade(item["total_marks"])
                    
                    phase2_rows.append(row_dict)

                df_phase2 = pd.DataFrame(phase2_rows)
                st.dataframe(df_phase2, use_container_width=True)

                csv_p2 = df_phase2.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Download Phase 2 Full Matrix Roll for {analysis_grade} (.csv)",
                    data=csv_p2,
                    file_name=f"{analysis_grade.replace(' ', '_')}_Full_Matrix_Class_Roll.csv",
                    mime="text/csv",
                    type="primary"
                )

    # =========================================================
    # TAB 3: BEST THREE PER LEARNING AREA (SEPARATE TABLES)
    # =========================================================
    with tab_top_subjects:
        st.markdown(f"### 🥇 Best Three Students Per Learning Area — {analysis_grade}")
        if not active_ranked_students:
            st.info(f"No student records found for {analysis_grade}.")
        else:
            df_m_all = pd.DataFrame(marks_data)
            
            for sub in LEARNING_AREAS:
                col_k = get_db_col_name(sub)
                
                valid_sub_students = []
                for item in active_ranked_students:
                    val = item["marks_row"].get(col_k)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        valid_sub_students.append({
                            "Student Name": item["name"],
                            "Adm No": item["adm_no"],
                            "Score (%)": float(val),
                            "Performance Level": get_subject_performance(float(val))[1]
                        })
                
                valid_sub_students = sorted(valid_sub_students, key=lambda x: x["Score (%)"], reverse=True)[:3]
                
                with st.expander(f"📚 {sub}", expanded=True):
                    if valid_sub_students:
                        sub_top_rows = []
                        for pos_idx, top_s in enumerate(valid_sub_students, 1):
                            medal = "🥇 Position 1" if pos_idx == 1 else ("🥈 Position 2" if pos_idx == 2 else "🥉 Position 3")
                            sub_top_rows.append({
                                "Position": medal,
                                "Adm No": top_s["Adm No"],
                                "Student Name": top_s["Student Name"],
                                "Score (%)": top_s["Score (%)"],
                                "Performance Level": top_s["Performance Level"]
                            })
                        
                        st.dataframe(pd.DataFrame(sub_top_rows), use_container_width=True, hide_index=True)
                    else:
                        st.caption(f"No submitted scores recorded yet for {sub}.")

    # =========================================================
    # TAB 4: ASSESSMENT REPORT FORMS HUB
    # =========================================================
    with tab_reports:
        st.markdown("### Assessment Report Form Generator")
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            term_val = st.selectbox("Term", ["Term 1", "Term 2", "Term 3"], key="rep_term")
            opening_date = st.text_input("Opening Date", "14/09/2026", key="rep_open")
        with col_s2:
            closing_date = st.text_input("Closing Date", "04/12/2026", key="rep_close")
            stamp_upload = st.file_uploader("Upload Stamp Image (stamp.png)", type=["png", "jpg"], key="rep_stamp")

        st.markdown("---")
        
        if not active_ranked_students:
            st.info(f"No students with active exam scores found in {analysis_grade}. Report forms are suppressed for students without marks.")
        else:
            def generate_student_docx_object(student_obj):
                adm_str = student_obj["adm_no"]
                assess_str = student_obj.get("assessment_no", "")
                student_name = student_obj["name"]
                s_row = student_obj["marks_row"]
                student_rank = student_obj["rank"]

                if os.path.exists("report.docx"):
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
                    p_logo.paragraph_format.space_before = Pt(0)
                    p_logo.paragraph_format.space_after = Pt(0)
                    run_logo = p_logo.add_run()
                    run_logo.add_picture("logo.png", width=Inches(0.95))

                total_score = 0.0
                top_sub, low_sub = "Mathematics", "English"
                best_s, worst_s = -1, 101

                subject_scores = {}
                for sub in LEARNING_AREAS:
                    col_key = get_db_col_name(sub)
                    val = s_row.get(col_key)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        s_num = float(val)
                        total_score += s_num
                        if s_num > best_s: best_s = s_num; top_sub = sub
                        if s_num < worst_s: worst_s = s_num; low_sub = sub
                        g_code, p_lvl, pts = get_subject_performance(s_num)
                        subject_scores[sub.upper()] = (str(s_num), p_lvl)
                    else:
                        subject_scores[sub.upper()] = ("-", "-")

                overall_g = calculate_total_grade(total_score)
                teacher_comment = generate_teacher_comment(total_score, top_sub, low_sub)

                for t in doc.tables:
                    t.alignment = WD_TABLE_ALIGNMENT.CENTER
                    for row in t.rows:
                        row_text = " ".join([cell.text for cell in row.cells]).upper()
                        
                        for cell in row.cells:
                            cell_upper = cell.text.upper()
                            if "NAME:" in cell_upper:
                                cell.text = f"NAME: {student_name}"
                            elif "ADM NO:" in cell_upper or "ADM.NO:" in cell_upper:
                                cell.text = f"ADM NO: {adm_str}"
                            elif "ASSES.NO:" in cell_upper or "ASSESSMENT NO:" in cell_upper or "ASSES NO:" in cell_upper:
                                cell.text = f"ASSES.NO: {assess_str}"
                            elif "TERM:" in cell_upper:
                                cell.text = f"TERM: {term_val}"
                            elif "POSITION:" in cell_upper or "RANK:" in cell_upper:
                                cell.text = f"POSITION: #{student_rank}"
                            elif "YEAR:" in cell_upper:
                                cell.text = f"YEAR: 2026"
                                
                        for sub_name, (score_str, perf_str) in subject_scores.items():
                            if sub_name in row_text and len(row.cells) >= 5:
                                row.cells[3].text = score_str
                                row.cells[4].text = perf_str

                        for cell in row.cells:
                            for p in cell.paragraphs:
                                p.paragraph_format.space_before = Pt(3)
                                p.paragraph_format.space_after = Pt(3)

                replacements = {
                    "TOTAL MARKS: _____": f"TOTAL MARKS: {total_score:.0f}",
                    "GENERAL PERFORMANCE LEVEL:": f"GENERAL PERFORMANCE LEVEL: {overall_g}",
                    "TEACHERS GENERAL COMMENT:": f"TEACHERS GENERAL COMMENT:\n\"{teacher_comment}\"",
                    "CLOSING DATE: __________": f"CLOSING DATE: {closing_date}",
                    "OPENING DATE: __________": f"OPENING DATE: {opening_date}",
                }

                for p in doc.paragraphs:
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(4)
                    for k, v in replacements.items():
                        if k in p.text:
                            p.text = p.text.replace(k, v)

                if stamp_upload or os.path.exists("stamp.png"):
                    stamp_source = stamp_upload if stamp_upload else "stamp.png"
                    for p in doc.paragraphs:
                        if "H.O.I STAMP" in p.text or "SIGNATURE:" in p.text:
                            run = p.add_run()
                            run.add_picture(stamp_source, width=Inches(1.2))
                            break

                return doc

            st.subheader("1. Live Student Report Form Preview")
            
            selected_student_obj = st.selectbox(
                "Select Student for Live Report Form Preview", 
                active_ranked_students, 
                format_func=lambda x: f"Rank #{x['rank']} | Adm No: {x['adm_no']} | Name: {x['name']} (Total: {x['total_marks']:.0f})"
            )

            if selected_student_obj:
                doc_obj = generate_student_docx_object(selected_student_obj)
                
                s_name = selected_student_obj['name']
                s_adm = selected_student_obj['adm_no']
                s_assess = selected_student_obj.get('assessment_no', '')
                s_rank = selected_student_obj['rank']
                s_total = selected_student_obj['total_marks']
                s_row = selected_student_obj['marks_row']
                overall_g = calculate_total_grade(s_total)
                
                top_s, low_s = "Mathematics", "English"
                b_val, w_val = -1, 101
                for sub in LEARNING_AREAS:
                    col_key = get_db_col_name(sub)
                    val = s_row.get(col_key)
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        v = float(val)
                        if v > b_val: b_val = v; top_s = sub
                        if v < w_val: w_val = v; low_s = sub
                teacher_comment = generate_teacher_comment(s_total, top_s, low_s)

                st.markdown("#### 👁️ Live Report Form Document View")
                
                html_preview = f"""
                <div style="border: 2px solid #374151; padding: 20px; border-radius: 8px; background-color: #FFFFFF; color: #111827; font-family: Arial, sans-serif; min-height: 900px; display: flex; flex-direction: column; justify-content: space-between;">
                    
                    <div>
                        <div style="text-align: center; margin-bottom: 10px;">
                            {'<img src="logo.png" style="height: 65px; margin-bottom: 5px;" /><br/>' if os.path.exists("logo.png") else ''}
                            <h2 style="margin: 0; color: #1E3A8A; font-size: 20px;">KEA COMPREHENSIVE SCHOOL</h2>
                            <p style="margin: 2px 0; font-size: 13px;">P.O. BOX 557-40400 SUNA MIGORI</p>
                            <h4 style="margin: 4px 0; text-decoration: underline; font-size: 14px;">STUDENT ASSESSMENT REPORT</h4>
                        </div>

                        <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 12px; border: 1px solid #374151;">
                            <tr>
                                <td style="border: 1px solid #374151; padding: 5px; font-weight: bold; background-color: #F9FAFB;">NAME: {s_name}</td>
                                <td style="border: 1px solid #374151; padding: 5px; font-weight: bold;">YEAR: 2026</td>
                            </tr>
                            <tr>
                                <td style="border: 1px solid #374151; padding: 5px; font-weight: bold;">TERM: {term_val}</td>
                                <td style="border: 1px solid #374151; padding: 5px; font-weight: bold;">POSITION: #{s_rank}</td>
                            </tr>
                            <tr>
                                <td style="border: 1px solid #374151; padding: 5px; font-weight: bold; background-color: #F9FAFB;">ADM NO: {s_adm}</td>
                                <td style="border: 1px solid #374151; padding: 5px; font-weight: bold; background-color: #F9FAFB;">ASSES.NO: {s_assess}</td>
                            </tr>
                            <tr>
                                <td colspan="2" style="border: 1px solid #374151; padding: 5px; font-weight: bold;">GRADE: {analysis_grade}</td>
                            </tr>
                        </table>

                        <table style="width: 100%; border-collapse: collapse; font-size: 12px; border: 1px solid #374151;">
                            <tr style="background-color: #1E3A8A; color: #FFFFFF; font-weight: bold; text-align: left;">
                                <th style="border: 1px solid #374151; padding: 6px;">S/N</th>
                                <th style="border: 1px solid #374151; padding: 6px;">CODE</th>
                                <th style="border: 1px solid #374151; padding: 6px;">LEARNING AREA</th>
                                <th style="border: 1px solid #374151; padding: 6px; text-align: center;">MARKS SCORED</th>
                                <th style="border: 1px solid #374151; padding: 6px;">PERFORMANCE LEVEL</th>
                            </tr>
                """

                for idx, sub in enumerate(LEARNING_AREAS, 1):
                    col_key = get_db_col_name(sub)
                    val = s_row.get(col_key)
                    code_val = f"90{idx}" if idx < 10 else f"9{idx}"
                    
                    if val is not None and not pd.isna(val) and float(val) >= 0.0:
                        s_num = float(val)
                        g_code, p_lvl, pts = get_subject_performance(s_num)
                        m_str = f"{s_num:.1f}"
                        p_str = p_lvl
                    else:
                        m_str = "-"
                        p_str = "-"

                    html_preview += f"""
                            <tr>
                                <td style="border: 1px solid #374151; padding: 5px;">{idx}</td>
                                <td style="border: 1px solid #374151; padding: 5px;">{code_val}</td>
                                <td style="border: 1px solid #374151; padding: 5px; font-weight: bold;">{sub.upper()}</td>
                                <td style="border: 1px solid #374151; padding: 5px; text-align: center; font-weight: bold; color: #1E3A8A;">{m_str}</td>
                                <td style="border: 1px solid #374151; padding: 5px;">{p_str}</td>
                            </tr>
                    """

                html_preview += f"""
                        </table>
                    </div>

                    <div style="margin-top: 15px; font-size: 13px; border-top: 2px solid #374151; padding-top: 10px;">
                        <p style="margin: 6px 0;"><strong>TOTAL MARKS:</strong> <span style="font-size: 15px; color: #1E3A8A; font-weight: bold;">{s_total:.0f}</span> &nbsp;|&nbsp; <strong>GENERAL PERFORMANCE LEVEL:</strong> {overall_g}</p>
                        
                        <p style="margin: 8px 0 2px 0;"><strong>TEACHER'S GENERAL COMMENT:</strong></p>
                        <div style="border-bottom: 1px dashed #374151; padding: 6px; font-style: italic; background-color: #F9FAFB;">"{teacher_comment}"</div>
                        
                        <div style="display: flex; justify-content: space-between; margin-top: 12px; font-weight: bold;">
                            <div>CLOSING DATE: {closing_date}</div>
                            <div>OPENING DATE: {opening_date}</div>
                        </div>

                        <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 25px; font-weight: bold; position: relative;">
                            <div>CLASS TEACHER SIGNATURE: ____________</div>
                            <div style="text-align: right; position: relative;">
                                {'<img src="stamp.png" style="position: absolute; right: 20px; bottom: -10px; height: 75px; opacity: 0.85;" />' if os.path.exists("stamp.png") else ''}
                                H.O.I STAMP & SIGNATURE: ____________
                            </div>
                        </div>
                    </div>

                </div>
                """

                st.components.v1.html(html_preview, height=620, scrolling=True)

                out_stream = io.BytesIO()
                doc_obj.save(out_stream)
                out_stream.seek(0)
                
                st.download_button(
                    label=f"📥 Download Assessment Report Form for {s_name} (.docx)",
                    data=out_stream,
                    file_name=f"Rank_{s_rank}_{s_adm}_{s_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

            st.markdown("---")

            st.subheader("2. Ranked Batch Reports Generator (Entire Grade)")
            st.caption(f"Generates report forms ordered from Rank #1 to Rank #{len(active_ranked_students)}. Students without marks are excluded.")

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
                    st.success(f"Successfully compiled {len(active_ranked_students)} ranked report forms into ZIP archive!")
                    
                    st.download_button(
                        label=f"📥 Download All {len(active_ranked_students)} Ranked Reports (.zip)",
                        data=zip_buffer,
                        file_name=f"{analysis_grade.replace(' ', '_')}_Ranked_Assessment_Reports.zip",
                        mime="application/zip"
                    )

# --- PAGE 5: FEE PAYMENT PORTAL ---
elif page == "Fee Payment":
    if not is_admin_role:
        st.error("Access Denied: Only HOI, DHOI, and Senior Teachers have access to Financial Records.")
        st.stop()
        
    st.header("Fee Payment & Financial Portal")
    tab_fee_struct, tab_make_pay, tab_history, tab_balances = st.tabs(["Fee Structure", "Make Payment", "Receipt & History", "Termly Balances"])
    
    with tab_fee_struct:
        st.markdown("### KEA Comprehensive School Fee Structure")
        st.info("""
        * **Continuing Students:** Ksh 300 per term (Tuition: Ksh 100 [50/month], Exams: Ksh 200 [100 Mid Term, 100 End Term]). Total Annual = Ksh 900
        * **Newly Admitted Students:** Pay an extra Ksh 200 admission fee in Term 1 (Total Term 1 = Ksh 500). Terms 2 & 3 revert to standard Ksh 300.
        """)
        
    with tab_make_pay:
        try:
            st_res = supabase.table("students").select("*").execute()
            all_students = st_res.data
        except:
            all_students = []
            
        if all_students:
            with st.form("payment_form"):
                payer_student = st.selectbox("Search / Select Student", all_students, format_func=lambda x: f"{x['adm_no']} - {x['name']} ({x['grade']})")
                amount = st.number_input("Amount Paid (Ksh)", min_value=0.0, value=550.0)
                channel = st.selectbox("Payment Channel", ["Cash Payment", "Bank Payment", "M-Pesa Payment"])
                payer_name = st.text_input("Name of Payer")
                category = st.selectbox("Payment Category", ["Whole Term Payment", "Tuition Fee", "Exams Fee", "Admission Fee"])
                
                submit_pay = st.form_submit_button("Record Payment")
                if submit_pay:
                    try:
                        supabase.table("fee_payments").insert({
                            "adm_no": payer_student["adm_no"],
                            "student_name": payer_student["name"],
                            "grade": payer_student["grade"],
                            "amount_paid": amount,
                            "channel": channel,
                            "payer_name": payer_name,
                            "category": category,
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }).execute()
                        st.success(f"Successfully recorded payment of Ksh {amount} for {payer_student['name']}!")
                    except Exception as e:
                        st.error(f"Error recording payment: {e}")

    with tab_history:
        st.markdown("### Payment History & Receipt Generation")
        try:
            pay_res = supabase.table("fee_payments").select("*").order("timestamp", desc=True).execute()
            payments = pay_res.data
            if payments:
                df_pay = pd.DataFrame(payments)
                st.dataframe(df_pay, use_container_width=True)
            else:
                st.info("No payment history recorded yet.")
        except Exception as e:
            st.error(f"Could not load payment history: {e}")

    with tab_balances:
        st.markdown("### Termly Balances Status")
        try:
            st_res = supabase.table("students").select("*").execute()
            pay_res = supabase.table("fee_payments").select("*").execute()
            
            student_paid = {}
            for p in pay_res.data:
                adm = p["adm_no"]
                student_paid[adm] = student_paid.get(adm, 0) + p["amount_paid"]
                
            balance_rows = []
            for s in st_res.data:
                paid = student_paid.get(s["adm_no"], 0)
                required = 300
                diff = paid - required
                if diff == 0:
                    status = "🟢 Nil Balance"
                elif diff < 0:
                    status = f"🔴 Pending Balance (Deficit: {-diff})"
                else:
                    status = f"🔵 Overpayment (Credit: +{diff})"
                    
                balance_rows.append({
                    "Adm No": s["adm_no"],
                    "Name": s["name"],
                    "Grade": s["grade"],
                    "Paid (Ksh)": paid,
                    "Status": status
                })
            
            if balance_rows:
                st.dataframe(pd.DataFrame(balance_rows), use_container_width=True)
        except Exception as e:
            st.error(f"Error calculating balances: {e}")

# --- PAGE 6: TEACHERS PORTAL ---
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
                st_res = supabase.table("students").select("adm_no", "assessment_no", "name").eq("grade", grade).execute()
                if st_res.data:
                    st.markdown(f"**Registered Students in {grade}:**")
                    st.dataframe(pd.DataFrame(st_res.data), use_container_width=True)
                else:
                    st.caption(f"No students currently registered under {grade}.")
            except Exception as e:
                st.error(f"Could not load students for {grade}: {e}")
            st.markdown("---")
    else:
        st.warning("No teaching assignments found for your account.")

# --- PAGE 7: TEACHER TIME LOGIN ---
elif page == "Teacher Time Login":
    st.header("Teacher Attendance & Time Log")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Clock In (Time In)", type="primary"):
            st.session_state.show_in_popup = True
            
        if st.session_state.get("show_in_popup", False):
            with st.form("time_in_form"):
                t_name = st.text_input("Enter Your First Name")
                if st.form_submit_button("Submit Time In"):
                    try:
                        supabase.table("teacher_time_logs").insert({
                            "teacher_name": t_name,
                            "time_in": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "day_of_week": datetime.now().strftime("%A")
                        }).execute()
                        st.success(f"Successfully clocked IN for {t_name}!")
                        st.session_state.show_in_popup = False
                    except Exception as e:
                        st.error(f"Error recording time in: {e}")

    with col2:
        if st.button("Clock Out (Time Out)", type="secondary"):
            st.session_state.show_out_popup = True
            
        if st.session_state.get("show_out_popup", False):
            with st.form("time_out_form"):
                t_name_out = st.text_input("Enter Your First Name for Clock Out")
                if st.form_submit_button("Submit Time Out"):
                    try:
                        supabase.table("teacher_time_logs").update({
                            "time_out": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }).eq("teacher_name", t_name_out).is_("time_out", "null").execute()
                        st.success(f"Successfully clocked OUT for {t_name_out}!")
                        st.session_state.show_out_popup = False
                    except Exception as e:
                        st.error(f"Error recording time out: {e}")

    st.markdown("### Weekly Attendance Logs (Mon - Fri)")
    try:
        logs_res = supabase.table("teacher_time_logs").select("*").execute()
        if logs_res.data:
            st.dataframe(pd.DataFrame(logs_res.data), use_container_width=True)
    except:
        st.info("No attendance logs found.")

# --- PAGE 8: NEWSLETTER ---
elif page == "Newsletter":
    st.header("School Newsletter & Upcoming Events")
    
    if is_admin_role:
        with st.form("newsletter_form"):
            news_title = st.text_input("Newsletter Title")
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
    else:
        st.info("Note: You are viewing in read-only mode. Only administrators can publish updates.")

    try:
        n_res = supabase.table("newsletter").select("*").order("date", desc=True).execute()
        for item in n_res.data:
            st.markdown(f"### {item['title']}")
            st.markdown(f"*{item['date']}*")
            st.write(item['content'])
            st.markdown("---")
    except:
        st.info("No newsletters posted yet.")

# --- PAGE 9: SCHOOL CONTACT ---
elif page == "School Contact":
    st.header("School Contact & Location Information")
    
    if is_admin_role:
        with st.form("contact_form"):
            address = st.text_input("Physical Address", "Migori County, Kenya")
            phone = st.text_input("School Contact Number", "+254 700 000 000")
            email = st.text_input("Email Address", "info@keacomprehensiveschool.ac.ke")
            maps_link = st.text_input("Google Maps Embed URL / Location Link", "https://maps.google.com")
            
            if st.form_submit_button("Update Contact Info"):
                try:
                    supabase.table("school_contact").upsert({
                        "id": 1,
                        "address": address,
                        "phone": phone,
                        "email": email,
                        "maps_link": maps_link
                    }).execute()
                    st.success("Contact information updated!")
                except Exception as e:
                    st.error(f"Error updating contact: {e}")
    
    st.markdown("### Contact Details")
    st.markdown("**Physical Address:** Migori County, Kenya")
    st.markdown("**Phone:** +254 700 000 000")
    st.markdown("**Email:** info@keacomprehensiveschool.ac.ke")
    st.markdown("[View Location on Google Maps](https://maps.google.com)")
