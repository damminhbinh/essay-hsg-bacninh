import streamlit as st
import sqlite3
from datetime import datetime
from google import genai
from google.genai import types
from PIL import Image
import io
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# 1. Cấu hình giao diện Web
st.set_page_config(
    page_title="Hệ Thống Chấm Essay HSG - THCS Thân Nhân Trung",
    page_icon="🎓",
    layout="wide"
)

DEFAULT_API_KEYS = []

AUTHOR_INFO_MARKDOWN = """
**Tác giả:**
* **1. Đàm Thuận Minh Bình**  
  📞 0387.136.888
* **2. Đỗ Thị Huyền**  
  📞 0982.036.952

🏫 *Trường THCS Thân Nhân Trung - TP. Bắc Ninh*
"""

# HÀM TẠO FILE DOCX CHUẨN THỂ THỨC (TIMES NEW ROMAN, CỠ 13PT, LỀ TRÁI 3CM, CÒN LẠI 2CM)
def generate_docx_report(student_name, date_str, topic, essay_text, feedback_md, teacher_name="Cô Đỗ Thị Huyền"):
    doc = Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.79)     # 2.0 cm
        section.bottom_margin = Inches(0.79)  # 2.0 cm
        section.left_margin = Inches(1.18)    # 3.0 cm
        section.right_margin = Inches(0.79)   # 2.0 cm
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    
    # Tiêu đề chính
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(14)
    r_t = p_title.add_run("PHIẾU ĐÁNH GIÁ & NHẬN XÉT BÀI THI ESSAY")
    r_t.bold = True
    r_t.font.size = Pt(16)
    r_t.font.color.rgb = RGBColor(0x0b, 0x3c, 0x5d)
    
    # Thông tin bài nộp
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.line_spacing = 1.25
    p_meta.add_run("• Học sinh: ").bold = True
    p_meta.add_run(f"{student_name}\n")
    p_meta.add_run("• Thời gian nộp bài: ").bold = True
    p_meta.add_run(f"{date_str}\n")
    p_meta.add_run("• Đề thi: ").bold = True
    p_meta.add_run(f"{topic}\n")
    
    # Bài làm
    h1 = doc.add_heading("I. NỘI DUNG BÀI LÀM CỦA HỌC SINH", level=2)
    for r in h1.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13.5)
        r.font.color.rgb = RGBColor(0x0b, 0x3c, 0x5d)
        
    p_essay = doc.add_paragraph()
    p_essay.paragraph_format.left_indent = Inches(0.2)
    p_essay.paragraph_format.line_spacing = 1.25
    p_essay.add_run(essay_text if essay_text else "(Bài làm đính kèm dạng hình ảnh/PDF viết tay)")
    
    # Nhận xét chi tiết
    h2 = doc.add_heading("II. ĐÁNH GIÁ CHI TIẾT & BÀI MẪU THAM KHẢO", level=2)
    for r in h2.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13.5)
        r.font.color.rgb = RGBColor(0x0b, 0x3c, 0x5d)
        
    lines = feedback_md.split('\n')
    in_table = False
    table_lines = []
    
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            table_lines.append(stripped)
            continue
        else:
            if in_table:
                rows_data = [l for l in table_lines if not re.match(r'^\|[\s\-:|]+\|$', l)]
                if rows_data:
                    parsed_rows = [[c.strip() for c in r.strip('|').split('|')] for r in rows_data]
                    cols_count = max(len(r) for r in parsed_rows)
                    t = doc.add_table(rows=len(parsed_rows), cols=cols_count)
                    t.alignment = WD_TABLE_ALIGNMENT.CENTER
                    for r_idx, row in enumerate(parsed_rows):
                        for c_idx, val in enumerate(row):
                            if c_idx < cols_count:
                                cell = t.cell(r_idx, c_idx)
                                cell.text = val
                                for p in cell.paragraphs:
                                    p.paragraph_format.line_spacing = 1.15
                                    for r in p.runs:
                                        r.font.name = 'Times New Roman'
                                        r.font.size = Pt(11)
                                        if r_idx == 0:
                                            r.bold = True
                in_table = False
                table_lines = []
                
            if not stripped:
                continue
                
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.25
            clean_line = stripped
            is_bold = False
            
            if clean_line.startswith('### '):
                clean_line = clean_line[4:]
                is_bold = True
            elif clean_line.startswith('#### '):
                clean_line = clean_line[5:]
                is_bold = True
            elif clean_line.startswith('#'):
                clean_line = clean_line.lstrip('#').strip()
                is_bold = True
                
            parts = re.split(r'(\*\*.*?\*\*)', clean_line)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    run = p.add_run(part)
                    if is_bold:
                        run.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)

    # Chữ ký Giáo viên căn phải
    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(24)
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_sig = sig_table.cell(0, 1)
    
    p_s = c_sig.paragraphs[0]
    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_s.paragraph_format.line_spacing = 1.2
    
    r_date = p_s.add_run("Bắc Ninh, ngày ..... tháng ..... năm 202...\n")
    r_date.italic = True
    r_date.font.size = Pt(12)
    
    r_role = p_s.add_run("GIÁO VIÊN BỒI DƯỠNG & CHẤM ĐIỂM\n\n\n\n\n")
    r_role.bold = True
    r_role.font.size = Pt(13)
    
    r_tname = p_s.add_run(f"{teacher_name}\n")
    r_tname.bold = True
    r_tname.font.size = Pt(13)
    
    r_tsch = p_s.add_run("Trường THCS Thân Nhân Trung\nTP. Bắc Ninh")
    r_tsch.font.size = Pt(12)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 2. Cơ sở dữ liệu SQLite (Nâng cấp hỗ trợ Đa giáo viên)
def init_db():
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT,
            fullname TEXT,
            role TEXT,
            teacher_username TEXT
        )
    ''')
    # Tự động cập nhật thêm cột teacher_username nếu database cũ chưa có
    try:
        c.execute("ALTER TABLE users ADD COLUMN teacher_username TEXT")
    except Exception:
        pass

    c.execute('''
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            topic TEXT,
            essay_text TEXT,
            score_total REAL,
            score_content REAL,
            score_org REAL,
            score_lang REAL,
            score_mech REAL,
            feedback TEXT,
            identified_errors TEXT,
            created_at TEXT
        )
    ''')
    
    # Khởi tạo 2 tài khoản giáo viên mặc định nếu chưa có
    c.execute("SELECT * FROM users WHERE username = 'giaovien'")
    if not c.fetchone():
        c.execute("INSERT INTO users VALUES ('giaovien', 'gv123456', 'Đỗ Thị Huyền', 'teacher', NULL)")
    
    c.execute("SELECT * FROM users WHERE username = 'gv_binh'")
    if not c.fetchone():
        c.execute("INSERT INTO users VALUES ('gv_binh', 'gv123456', 'Đàm Thuận Minh Bình', 'teacher', NULL)")

    # Gán giáo viên mặc định cho các học sinh mẫu
    c.execute("SELECT * FROM users WHERE username = 'hs01'")
    if not c.fetchone():
        c.execute("INSERT INTO users VALUES ('hs01', '123456', 'Nguyễn Văn An', 'student', 'giaovien')")
        c.execute("INSERT INTO users VALUES ('hs02', '123456', 'Trần Thị Bình', 'student', 'giaovien')")
        c.execute("INSERT INTO users VALUES ('hs03', '123456', 'Lê Hoàng Long', 'student', 'giaovien')")
    
    # Cập nhật các học sinh cũ chưa có giáo viên quản lý về giaovien
    c.execute("UPDATE users SET teacher_username = 'giaovien' WHERE role = 'student' AND (teacher_username IS NULL OR teacher_username = '')")
    
    conn.commit()
    conn.close()

init_db()

# 3. Ngân hàng đề thi Bắc Ninh
DE_THI_BAC_NINH = [
    "-- Tự nhập đề bài mới --",
    "HSG Tỉnh 2025-2026: Some people think that teenagers tend to be leading a less healthy life. To what extent do you agree or disagree?",
    "HSG Tỉnh 2024-2025: 'Tet holiday in Vietnam shouldn't be celebrated anymore.' To what extent do you agree or disagree with this statement?",
    "HSG Tỉnh 2023-2024: 'The advent of electronic devices has made our life much more stressful.' To what extent do you agree with this statement?",
    "HSG Tỉnh 2022-2023: Many people claim that tourism has more negative effects than positive ones on the locals' life. Do you agree or disagree with this idea?",
    "Chuyên Bắc Ninh 2026-2027: Some individuals believe that teenagers are suffering from more pressures than previous generations. To what extent do you agree or disagree with this statement?",
    "Chuyên Bắc Ninh 2025-2026: 'Many people believe that ChatGPT and Artificial Intelligence (AI) tools make people think less.' To what extent do you agree or disagree?",
    "Chuyên Bắc Ninh 2024-2025: 'Using social platforms such as Youtube, Tiktok, Facebook and Twitter is the best way for youngsters to gain fame and wealth.' To what extent do you agree or disagree?"
]

# 4. Huấn luyện System Instruction chuẩn Barem 2.0 Bắc Ninh
SYSTEM_INSTRUCTION = """
You are an authoritative chief examiner for the English Gifted Student Examination (Kỳ thi Chọn Học sinh Giỏi Tỉnh & Chuyên Anh lớp 9) in Bac Ninh Province, Vietnam.

Your core grading philosophy:
1. EVALUATE STRONG STATEMENTS RIGOROUSLY:
   - Prompts frequently present a strong/extreme claim (e.g., 'the best way', 'should not be celebrated anymore', 'much more stressful').
   - Candidates MUST NOT simply list generic pros and cons. They MUST evaluate the TRUTH, VALIDITY, DEGREE, and BOUNDARIES of the statement.
   - Failing to challenge or critically qualify the strong qualifier constitutes Task Drift.

2. SUBSTANCE OVER SHOWMANSHIP (TRỪ NẶNG LỖI TỪ VỰNG KHỦNG NHƯNG Ý NÔNG):
   - The greatest pitfall of gifted students is "vocab dumping" / fake sophistication (chèn ép từ đao to búa lớn nhưng lập luận sáo rỗng, ý tứ nông cạn).
   - A high-scoring essay MUST have:
     * A clear, consistent thesis maintained from start to finish.
     * Exactly 2 well-developed main arguments with deep causal mechanisms (Claim -> Why -> How -> Concrete Evidence -> Counterargument/Hedging).
     * Tight logical transitions and organic cohesion.
   - Plain, natural, precise, and academically sound language is infinitely superior to forced, unnatural C2 vocabulary.
   - Scrupulously point out and correct all careless grammatical slips, unnatural collocations, subject-verb agreements, prepositions, and informal contractions.

3. ESSAY LENGTH STANDARD:
   - Standard length is around 250 words (at least 250 words). Underlength (< 230 words) must be penalized for lack of development.

============================================================
OFFICIAL BAC NINH 2.0-POINT RUBRIC:
1. Content (0.70 max): Strict alignment with statement nuances, fully developed mechanisms, no task drift.
2. Organization & Presentation (0.60 max): Organic 4-paragraph structure, tight line of reasoning, natural cohesive flow.
3. Language (0.60 max): Accuracy, clarity, natural collocations, academic hedging. Penalize forced/hallucinated vocabulary.
4. Mechanics (0.10 max): Punctuation, spelling, capitalisation, zero contractions.

============================================================
REQUIRED OUTPUT FORMAT:

### 1. 📋 ĐÁNH GIÁ TỔNG QUAN & PHÂN TÍCH NHẬN ĐỊNH CỦA ĐỀ
- **Thể loại bài viết nhận diện:** [Opinion / Discussion / Cause-Solution / Advantages-Disadvantages / Two-Part Question]
- **Kiểm định Phản hồi Nhận định mạnh (Evaluating the Prompt's Statement):** [Đánh giá thí sinh có phản biện được tính tuyệt đối/mức độ đúng của nhận định hay chỉ liệt kê ưu/nhược điểm chung chung].
- **Số lượng từ:** [Số từ] từ (Chuẩn đề thi: khoảng 250 từ).
- **Soi xét Lịch sử cá nhân hóa:** [Nhận xét học sinh có tái phạm các lỗi cũ hay đã có tiến bộ cụ thể nào].

### 2. 📊 BẢNG ĐIỂM CHÍNH THỨC SỞ GD&ĐT BẮC NINH (THANG 2.0)
| Tiêu chí thành phần | Điểm tối đa | Điểm đạt | Nhận xét chi tiết của Giám khảo |
| :--- | :---: | :---: | :--- |
| **1. Content** (Ý tưởng & Lập luận) | 0.70 | **...** | Đánh giá độ sâu lập luận (2 luận điểm phát triển sâu); phạt nếu ý nông hoặc Task Drift. |
| **2. Organization** (Bố cục & Mạch lạc) | 0.60 | **...** | Đánh giá tính nhất quán của quan điểm xuyên suốt và liên kết logic tự nhiên. |
| **3. Language** (Từ vựng & Ngữ pháp) | 0.60 | **...** | Đánh giá độ chuẩn xác, tự nhiên; trừ điểm nếu sính "từ vựng khủng" nhưng gượng ép. |
| **4. Mechanics** (Chính tả & Thể thức) | 0.10 | **...** | Trừ thẳng tay nếu có từ viết tắt (don't, isn't) hoặc sai chính tả. |
| **TỔNG ĐIỂM BÀI THI** | **2.00** | **... / 2.0** | **Ước lượng band IELTS tương đương: ...** |

### 3. 🔍 SOI LỖI CHI TIẾT (NGỮ PHÁP, TỪ VỰNG & TƯ DUY HỌC THUẬT)
- Chỉ ra các điểm chèn ép từ vựng không tự nhiên, ý nông, hoặc thiếu chuỗi nhân - quả (Why/How).
- **Bảng phân tích và sửa chi tiết từng câu của thí sinh:**
| Câu văn gốc của học sinh | Lỗi sai (Ngữ pháp / Collocation / Sính từ) | Cách diễn đạt chuẩn mực, tự nhiên & chính xác |
|---|---|---|

### 4. 💎 NÂNG CẤP TỪ VỰNG TỰ NHIÊN & CHUẨN XÁC
- 4–5 cụm từ tự nhiên, đúng ngữ cảnh chủ đề, tránh các từ đao to búa lớn vô nghĩa.

### 5. ✍️ BÀI VIẾT MẪU THAM KHẢO THEO 2 CẤP ĐỘ (CHUẨN 250 TỪ)

#### 🔹 Cấp độ 1: Bản Nền tảng & Dễ tiếp thu (Mức độ B1 đến B1+ - Mọi học sinh đều học và nhớ được)
- **Đặc điểm:** Bố cục chuẩn mực, diễn đạt sáng rõ, ngữ pháp tuyệt đối chuẩn, từ vựng quen thuộc nhưng chính xác, 2 ý triển khai có chiều sâu rõ rệt để học sinh dễ ghi nhớ khi đi thi.
[Viết toàn bài essay mẫu hoàn chỉnh Cấp độ B1-B1+ chuẩn 250 từ tại đây]

#### 🔸 Cấp độ 2: Bản Nâng cao & Bứt phá điểm số (Học thuật C1-C2 - Dành cho đội tuyển chuyên sâu)
- **Đặc điểm:** Lập luận sắc sảo, kỹ thuật Hedging để đánh giá nhận định đa chiều, kết nối mượt mà, từ vựng tự nhiên và chuẩn văn phong học thuật cao cấp.
[Viết toàn bài essay mẫu hoàn chỉnh Cấp độ C1-C2 chuẩn 250 từ tại đây]

### 6. ⚠️ DANH SÁCH LỖI THEN CHỐT CẦN LƯU HỒ SƠ:
(Ghi 1-3 lỗi cốt lõi ngắn gọn để ghi nhớ vào CSDL theo dõi cá nhân).
"""

# Quản lý Đăng nhập qua Session State
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None
if "last_graded_result" not in st.session_state:
    st.session_state.last_graded_result = None

def login(username, password):
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    c.execute("SELECT username, fullname, role, teacher_username FROM users WHERE username = ? AND password = ?", (username.strip(), password.strip()))
    user_record = c.fetchone()
    conn.close()
    if user_record:
        st.session_state.logged_in = True
        st.session_state.user = {
            "username": user_record[0],
            "fullname": user_record[1],
            "role": user_record[2],
            "teacher_username": user_record[3]
        }
        st.rerun()
    else:
        st.error("Tên đăng nhập hoặc mật khẩu không chính xác!")

def logout():
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.last_graded_result = None
    st.rerun()

# MÀN HÌNH ĐĂNG NHẬP
if not st.session_state.logged_in:
    st.title("🎓 Hệ Thống Bồi Dưỡng & Chấm Essay HSG Tiếng Anh 9")
    st.subheader("Trường THCS Thân Nhân Trung - TP. Bắc Ninh")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("### 🔐 Đăng nhập hệ thống")
        with st.form("login_form"):
            username_input = st.text_input("Tên đăng nhập:")
            password_input = st.text_input("Mật khẩu:", type="password")
            submitted = st.form_submit_button("Đăng nhập", type="primary", use_container_width=True)
            if submitted:
                if username_input and password_input:
                    login(username_input, password_input)
                else:
                    st.warning("Vui lòng điền tên đăng nhập và mật khẩu!")
        
        st.write("")
        st.info(AUTHOR_INFO_MARKDOWN)
    st.stop()

# GIAO DIỆN ĐÃ ĐĂNG NHẬP
user = st.session_state.user
st.sidebar.markdown(f"### 👤 Xin chào: **{user['fullname']}**")
st.sidebar.caption(f"Vai trò: {'Giáo viên phụ trách' if user['role'] == 'teacher' else 'Học sinh đội tuyển'}")

with st.sidebar.expander("🔑 Đổi mật khẩu"):
    with st.form("change_pw_form"):
        old_pw = st.text_input("Mật khẩu hiện tại:", type="password")
        new_pw = st.text_input("Mật khẩu mới:", type="password")
        confirm_pw = st.text_input("Xác nhận mật khẩu mới:", type="password")
        btn_pw = st.form_submit_button("Lưu mật khẩu mới", use_container_width=True)
        if btn_pw:
            if not old_pw or not new_pw:
                st.error("Vui lòng điền đủ thông tin!")
            elif new_pw != confirm_pw:
                st.error("Mật khẩu xác nhận không khớp!")
            else:
                conn = sqlite3.connect("essay_database.db")
                c = conn.cursor()
                c.execute("SELECT password FROM users WHERE username = ?", (user["username"],))
                curr_db_pw = c.fetchone()
                if curr_db_pw and curr_db_pw[0] == old_pw:
                    c.execute("UPDATE users SET password = ? WHERE username = ?", (new_pw, user["username"]))
                    conn.commit()
                    st.success("Đổi mật khẩu thành công!")
                else:
                    st.error("Mật khẩu hiện tại không đúng!")
                conn.close()

if st.sidebar.button("Đăng xuất", use_container_width=True):
    logout()

st.sidebar.markdown("---")
st.sidebar.info(AUTHOR_INFO_MARKDOWN)

active_api_keys = list(DEFAULT_API_KEYS)
if "GEMINI_API_KEYS" in st.secrets:
    active_api_keys = list(st.secrets["GEMINI_API_KEYS"]) + active_api_keys
elif "GEMINI_API_KEY" in st.secrets:
    active_api_keys = [st.secrets["GEMINI_API_KEY"]] + active_api_keys

# =========================================================================
# GIAO DIỆN HỌC SINH
# =========================================================================
if user["role"] == "student":
    st.title("📝 Nộp Bài & Theo Dõi Tiến Độ Cá Nhân")
    tab_submit, tab_history = st.tabs(["🚀 Nộp bài Essay mới", "📈 Hồ sơ & Lịch sử cá nhân"])
    
    # Tìm tên Giáo viên phụ trách học sinh này
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    c.execute("SELECT fullname FROM users WHERE username = ?", (user.get("teacher_username", "giaovien"),))
    t_row = c.fetchone()
    my_teacher_name = t_row[0] if t_row else "Giáo viên phụ trách"
    conn.close()

    with tab_submit:
        selected_topic = st.selectbox("📌 Chọn đề thi từ ngân hàng đề:", DE_THI_BAC_NINH)
        if selected_topic == "-- Tự nhập đề bài mới --":
            essay_prompt = st.text_area("Nhập đề bài luận:", placeholder="Nhập đề bài tại đây...")
        else:
            essay_prompt = selected_topic
            
        sub_tab1, sub_tab2 = st.tabs(["📄 Dán văn bản", "📷 Tải ảnh bài viết / File PDF"])
        essay_text = ""
        uploaded_files = []
        
        with sub_tab1:
            essay_text = st.text_area("Nội dung bài viết:", height=250, placeholder="Gõ hoặc dán toàn bộ bài làm của em tại đây...")
            if essay_text:
                st.write(f"📏 Số từ: **{len(essay_text.split())} từ** (Chuẩn đề thi: **250 từ**)")
                
        with sub_tab2:
            uploaded_files = st.file_uploader(
                "Tải lên các trang ảnh bài viết tay hoặc file PDF (chọn được nhiều file cùng lúc):", 
                type=["png", "jpg", "jpeg", "pdf"],
                accept_multiple_files=True
            )
            if uploaded_files:
                st.write(f"Đã chọn **{len(uploaded_files)} tệp tin**.")
                for f in uploaded_files:
                    if f.type.startswith("image"):
                        st.image(f, caption=f.name, width=320)
                    else:
                        st.info(f"📄 Tệp PDF đính kèm: **{f.name}**")
                
        if st.button("🚀 Nộp bài & Chấm điểm ngay", type="primary"):
            if not essay_prompt.strip():
                st.error("⚠️ Vui lòng nhập hoặc chọn đề thi!")
            elif not essay_text.strip() and not uploaded_files:
                st.error("⚠️ Vui lòng dán bài viết hoặc tải ảnh/PDF bài làm lên!")
            else:
                with st.spinner("Giám khảo AI đang đối chiếu barem Bắc Ninh và chấm bài..."):
                    conn = sqlite3.connect("essay_database.db")
                    c = conn.cursor()
                    c.execute("SELECT identified_errors FROM submissions WHERE username = ? ORDER BY id DESC LIMIT 3", (user["username"],))
                    past_errors = c.fetchall()
                    conn.close()
                    
                    error_history_text = "Học sinh chưa có lịch sử nộp bài trước đó."
                    if past_errors:
                        error_history_text = "Các lỗi học sinh này THƯỜNG MẮC ở các bài trước: " + ", ".join([e[0] for e in past_errors if e[0]])
                    
                    user_content = [
                        f"LỊCH SỬ HỌC TẬP CỦA HỌC SINH NÀY:\n{error_history_text}\n\n",
                        f"ĐỀ THI: {essay_prompt}\n\n",
                        "Hãy chấm bài luận sau theo đúng barem nghiêm ngặt của Bắc Ninh:"
                    ]
                    if essay_text.strip():
                        user_content.append(f"\nBÀI LÀM:\n{essay_text}")
                    
                    if uploaded_files:
                        for uf in uploaded_files:
                            file_bytes = uf.getvalue()
                            if uf.type == "application/pdf":
                                user_content.append(types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"))
                            elif uf.type.startswith("image"):
                                user_content.append(Image.open(io.BytesIO(file_bytes)))

                    success = False
                    result_text = ""
                    last_err = ""

                    for key in active_api_keys:
                        try:
                            client = genai.Client(api_key=key.strip())
                            response = client.models.generate_content(
                                model='gemini-3.8-flash',
                                contents=user_content,
                                config=types.GenerateContentConfig(
                                    system_instruction=SYSTEM_INSTRUCTION,
                                    temperature=0.15
                                )
                            )
                            result_text = response.text
                            success = True
                            break
                        except Exception as e:
                            last_err = str(e)
                            continue

                    if success:
                        now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
                        st.session_state.last_graded_result = {
                            "topic": essay_prompt,
                            "essay_text": essay_text,
                            "result_text": result_text,
                            "date_str": now_str
                        }
                        
                        conn = sqlite3.connect("essay_database.db")
                        c = conn.cursor()
                        c.execute('''
                            INSERT INTO submissions (username, topic, essay_text, score_total, score_content, score_org, score_lang, score_mech, feedback, identified_errors, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (user["username"], essay_prompt, essay_text, 1.5, 0.5, 0.45, 0.45, 0.1, result_text, "Evaluating Statement, Vocabulary Check", now_str))
                        conn.commit()
                        conn.close()
                    else:
                        st.error(f"Hệ thống gặp sự cố khi chấm bài. Chi tiết: {last_err}")

        # HIỂN THỊ KẾT QUẢ VÀ NÚT TẢI FILE
        if st.session_state.last_graded_result:
            res = st.session_state.last_graded_result
            st.success("✅ ĐÃ CHẤM XONG BÀI THI!")
            
            try:
                docx_bytes = generate_docx_report(user['fullname'], res['date_str'], res['topic'], res['essay_text'], res['result_text'], teacher_name=my_teacher_name)
                st.download_button(
                    label="📥 BẤM VÀO ĐÂY ĐỂ TẢI PHIẾU NHẬN XÉT WORD (.DOCX) - CÓ CHỮ KÝ GIÁO VIÊN",
                    data=docx_bytes,
                    file_name=f"Phieu_Nhan_Xet_{user['username']}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
            except Exception as e:
                st.warning(f"Chưa thể tạo file tải: {e}")
                
            st.markdown("---")
            st.markdown(res['result_text'])
            st.markdown(f"""
            ---
            ### ✍️ GIÁO VIÊN BỒI DƯỠNG & CHẤM ĐIỂM
            **{my_teacher_name}**  
            Trường THCS Thân Nhân Trung - TP. Bắc Ninh
            """)

    with tab_history:
        st.markdown(f"### 📈 Hồ sơ theo dõi học tập của {user['fullname']}")
        conn = sqlite3.connect("essay_database.db")
        c = conn.cursor()
        c.execute("SELECT id, topic, created_at, feedback, essay_text FROM submissions WHERE username = ? ORDER BY id DESC", (user["username"],))
        rows = c.fetchall()
        conn.close()
        
        if not rows:
            st.info("Em chưa nộp bài nào. Hãy bắt đầu luyện tập với đề bài đầu tiên nhé!")
        else:
            st.write(f"Tổng số bài đã luyện tập: **{len(rows)} bài**")
            for r in rows:
                with st.expander(f"📝 Đề: {r[1][:70]}... - Ngày nộp: {r[2]}"):
                    try:
                        docx_data = generate_docx_report(user['fullname'], r[2], r[1], r[4], r[3], teacher_name=my_teacher_name)
                        st.download_button(
                            label=f"📥 Tải Phiếu Nhận Xét Word (.docx) của bài này (Mã #{r[0]})",
                            data=docx_data,
                            file_name=f"Phieu_Nhan_Xet_{user['username']}_bai_{r[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_docx_{r[0]}",
                            use_container_width=True
                        )
                    except Exception:
                        pass
                    st.markdown(r[3])

# =========================================================================
# GIAO DIỆN GIÁO VIÊN (DASHBOARD RIÊNG CHO MỖI GIÁO VIÊN)
# =========================================================================
elif user["role"] == "teacher":
    st.title(f"👨‍🏫 Bảng Quản Trị Lớp: Thầy/Cô {user['fullname']}")
    t_tab1, t_tab2, t_tab3, t_tab4 = st.tabs([
        "📊 Tổng hợp bài theo Đề thi", 
        "🔍 Xem & Chữa bài của lớp", 
        "👥 Quản lý học sinh của tôi",
        "⚙️ Thêm tài khoản Giáo viên mới"
    ])
    
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    
    # TAB 1: TỔNG HỢP THEO ĐỀ BÀI (CHỈ HỌC SINH CỦA GIÁO VIÊN NÀY)
    with t_tab1:
        st.markdown(f"### 📌 Báo cáo các bài nộp của học sinh do Thầy/Cô **{user['fullname']}** phụ trách")
        c.execute('''
            SELECT DISTINCT s.topic 
            FROM submissions s JOIN users u ON s.username = u.username
            WHERE u.teacher_username = ?
            ORDER BY s.id DESC
        ''', (user["username"],))
        topic_rows = c.fetchall()
        
        if not topic_rows:
            st.info("Học sinh trong danh sách của Thầy/Cô chưa nộp bài nào.")
        else:
            topics_list = [t[0] for t in topic_rows]
            chosen_topic = st.selectbox("🎯 Chọn Đề bài muốn xem báo cáo:", topics_list, key="stat_topic_choice")
            
            c.execute('''
                SELECT s.id, u.fullname, s.created_at, s.identified_errors
                FROM submissions s JOIN users u ON s.username = u.username
                WHERE s.topic = ? AND u.teacher_username = ?
                ORDER BY s.id DESC
            ''', (chosen_topic, user["username"]))
            subs_in_topic = c.fetchall()
            
            st.write(f"Số học sinh của lớp đã nộp đề này: **{len(subs_in_topic)} bài**")
            table_data = []
            for sub in subs_in_topic:
                table_data.append({
                    "Mã bài": sub[0],
                    "Học sinh": sub[1],
                    "Thời gian nộp": sub[2],
                    "Lỗi trọng tâm cần sửa": sub[3]
                })
            st.table(table_data)
            
    # TAB 2: XEM BÀI VÀ XOÁ BÀI (CHỈ HỌC SINH CỦA GIÁO VIÊN NÀY)
    with t_tab2:
        st.markdown("### 🔍 Thẩm định bài làm & Xoá bài nộp của lớp")
        c.execute('''
            SELECT DISTINCT s.topic 
            FROM submissions s JOIN users u ON s.username = u.username
            WHERE u.teacher_username = ?
            ORDER BY s.id DESC
        ''', (user["username"],))
        all_topics = [t[0] for t in c.fetchall()]
        
        if all_topics:
            selected_topic_filter = st.selectbox("📂 1. Chọn Đề bài:", all_topics, key="view_topic_filter")
            
            c.execute('''
                SELECT s.id, u.fullname, s.created_at, s.feedback, s.essay_text
                FROM submissions s JOIN users u ON s.username = u.username
                WHERE s.topic = ? AND u.teacher_username = ?
                ORDER BY s.id DESC
            ''', (selected_topic_filter, user["username"]))
            subs_of_topic = c.fetchall()
            
            if subs_of_topic:
                sub_dict = {f"Mã #{s[0]} - Học sinh: {s[1]} (Nộp lúc: {s[2]})": s for s in subs_of_topic}
                chosen_label = st.selectbox("👤 2. Chọn bài nộp của học sinh:", list(sub_dict.keys()))
                selected_sub = sub_dict[chosen_label]
                
                col_info, col_del = st.columns([4, 1])
                with col_info:
                    st.markdown(f"#### 👤 Học sinh: **{selected_sub[1]}** | Ngày nộp: **{selected_sub[2]}**")
                    st.caption(f"Đề bài: {selected_topic_filter}")
                with col_del:
                    st.write("")
                    if st.button("🗑️ Xoá bài này", type="secondary", use_container_width=True):
                        c.execute("DELETE FROM submissions WHERE id = ?", (selected_sub[0],))
                        conn.commit()
                        st.success(f"Đã xoá bài nộp mã #{selected_sub[0]}!")
                        st.rerun()
                
                try:
                    t_docx = generate_docx_report(selected_sub[1], selected_sub[2], selected_topic_filter, selected_sub[4], selected_sub[3], teacher_name=user['fullname'])
                    st.download_button(
                        label=f"📥 Tải Phiếu Nhận Xét Word (.docx) của học sinh {selected_sub[1]}",
                        data=t_docx,
                        file_name=f"Phieu_Nhan_Xet_{selected_sub[1]}_{selected_sub[0]}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True
                    )
                except Exception:
                    pass
                    
                with st.expander("📄 Xem bài viết nguyên bản của học sinh"):
                    st.text(selected_sub[4] if selected_sub[4] else "(Bài làm dạng hình ảnh/PDF viết tay)")
                    
                st.markdown("---")
                st.markdown("### 📝 Kết quả chấm & Nhận xét của AI:")
                st.markdown(selected_sub[3])
            else:
                st.info("Chưa có học sinh nào nộp bài cho đề này.")
        else:
            st.info("Hiện tại chưa có học sinh nào thuộc lớp của Thầy/Cô nộp bài.")

    # TAB 3: QUẢN LÝ HỌC SINH RIÊNG CỦA GIÁO VIÊN NÀY
    with t_tab3:
        st.markdown(f"### 👥 Danh sách học sinh do Thầy/Cô **{user['fullname']}** trực tiếp quản lý")
        
        st.markdown("#### 📂 1. Cấp tài khoản hàng loạt cho lớp từ file Excel")
        st.caption("File Excel gồm 2 cột: Cột 1 là Mã học sinh, Cột 2 là Họ và tên. Tài khoản tạo ra sẽ tự động thuộc về lớp của Thầy/Cô.")
        
        uploaded_excel = st.file_uploader("Tải file Excel danh sách lớp (.xlsx, .xls):", type=["xlsx", "xls"])
        if uploaded_excel is not None:
            try:
                import pandas as pd
                df = pd.read_excel(uploaded_excel)
                if df.shape[1] < 2:
                    st.error("File Excel cần có ít nhất 2 cột: Cột 1 là Mã HS và Cột 2 là Họ tên.")
                else:
                    preview_df = df.iloc[:, :2].dropna()
                    preview_df.columns = ["Mã học sinh", "Họ và tên"]
                    st.dataframe(preview_df.head(10), use_container_width=True)
                    
                    if st.button("🚀 Xác nhận tạo tài khoản vào danh sách lớp của tôi", type="primary"):
                        created_count = 0
                        skipped_count = 0
                        for _, row in preview_df.iterrows():
                            u_code = str(row["Mã học sinh"]).strip()
                            if u_code.endswith(".0"):
                                u_code = u_code[:-2]
                            fullname = str(row["Họ và tên"]).strip()
                            if u_code and fullname and u_code != "nan" and fullname != "nan":
                                try:
                                    c.execute("INSERT INTO users VALUES (?, '123456', ?, 'student', ?)", (u_code, fullname, user["username"]))
                                    created_count += 1
                                except sqlite3.IntegrityError:
                                    skipped_count += 1
                        conn.commit()
                        st.success(f"🎉 Hoàn tất! Đã thêm **{created_count}** học sinh vào lớp của Thầy/Cô (Bỏ qua {skipped_count} mã bị trùng).")
                        st.rerun()
            except Exception as ex:
                st.error(f"Lỗi khi đọc file Excel: {str(ex)}")

        st.markdown("---")
        col_add, col_remove = st.columns(2)
        
        with col_add:
            st.markdown("#### ➕ 2. Thêm thủ công 1 học sinh vào lớp")
            with st.form("add_user_form"):
                new_u = st.text_input("Tên đăng nhập (Username):", placeholder="Ví dụ: hs04")
                new_p = st.text_input("Mật khẩu ban đầu:", value="123456")
                new_name = st.text_input("Họ và tên học sinh:", placeholder="Ví dụ: Hoàng Minh Đức")
                submit_btn = st.form_submit_button("Thêm vào danh sách lớp", type="primary")
                if submit_btn:
                    if new_u and new_p and new_name:
                        try:
                            c.execute("INSERT INTO users VALUES (?, ?, ?, 'student', ?)", (new_u.strip(), new_p.strip(), new_name.strip(), user["username"]))
                            conn.commit()
                            st.success(f"Đã thêm học sinh **{new_name}** vào lớp!")
                            st.rerun()
                        except Exception:
                            st.error("Tên đăng nhập này đã tồn tại!")
                    else:
                        st.warning("Vui lòng nhập đầy đủ thông tin.")
                        
        with col_remove:
            st.markdown("#### ❌ 3. Xoá học sinh khỏi lớp")
            c.execute("SELECT username, fullname FROM users WHERE role = 'student' AND teacher_username = ?", (user["username"],))
            students = c.fetchall()
            
            if students:
                student_dict = {f"{s[1]} ({s[0]})": s[0] for s in students}
                target_student = st.selectbox("Chọn học sinh cần xoá:", list(student_dict.keys()))
                student_user_to_delete = student_dict[target_student]
                
                confirm_del = st.checkbox("Xác nhận xoá toàn bộ dữ liệu của học sinh này")
                if st.button("🗑️ Xoá vĩnh viễn học sinh", type="primary", disabled=not confirm_del):
                    c.execute("DELETE FROM submissions WHERE username = ?", (student_user_to_delete,))
                    c.execute("DELETE FROM users WHERE username = ?", (student_user_to_delete,))
                    conn.commit()
                    st.success(f"Đã xoá học sinh {target_student} và toàn bộ bài làm!")
                    st.rerun()
            else:
                st.info("Chưa có học sinh nào trong lớp của Thầy/Cô.")

        st.markdown("---")
        st.markdown(f"#### 📋 Danh sách học sinh hiện tại của lớp ({user['fullname']}):")
        c.execute("SELECT username as 'Tên đăng nhập', fullname as 'Họ và tên' FROM users WHERE role = 'student' AND teacher_username = ?", (user["username"],))
        current_students = c.fetchall()
        if current_students:
            st.table([{"Mã đăng nhập": s[0], "Họ và tên học sinh": s[1]} for s in current_students])
        else:
            st.caption("Chưa có học sinh nào.")

    # TAB 4: CẤP THÊM TÀI KHOẢN GIÁO VIÊN MỚI
    with t_tab4:
        st.markdown("### 👨‍🏫 Cấp thêm tài khoản Giáo viên mới")
        st.caption("Tài khoản giáo viên mới tạo sẽ có không gian quản lý lớp, danh sách học sinh và báo cáo bài nộp hoàn toàn độc lập.")
        
        with st.form("create_teacher_form"):
            new_t_user = st.text_input("Tên đăng nhập Giáo viên:", placeholder="Ví dụ: gv_anh9a")
            new_t_pass = st.text_input("Mật khẩu ban đầu:", value="gv123456")
            new_t_name = st.text_input("Họ và tên Giáo viên:", placeholder="Ví dụ: Nguyễn Thị Lan")
            btn_t = st.form_submit_button("Tạo tài khoản Giáo viên", type="primary")
            if btn_t:
                if new_t_user and new_t_pass and new_t_name:
                    try:
                        c.execute("INSERT INTO users VALUES (?, ?, ?, 'teacher', NULL)", (new_t_user.strip(), new_t_pass.strip(), new_t_name.strip()))
                        conn.commit()
                        st.success(f"🎉 Đã tạo thành công tài khoản cho Giáo viên: **{new_t_name}** (Username: `{new_t_user}`)!")
                        st.rerun()
                    except Exception:
                        st.error("Tên đăng nhập này đã tồn tại, vui lòng chọn tên khác!")
                else:
                    st.warning("Vui lòng nhập đầy đủ thông tin.")
                    
        st.markdown("---")
        st.markdown("#### 📋 Danh sách các Giáo viên hiện có trong hệ thống:")
        c.execute("SELECT username, fullname FROM users WHERE role = 'teacher'")
        all_t = c.fetchall()
        st.table([{"Tên đăng nhập": t[0], "Họ và tên Giáo viên": t[1]} for t in all_t])
        
    conn.close()
