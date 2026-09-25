import streamlit as st
import sqlite3
import pandas as pd
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

# Khối thông tin Tác giả hiển thị dùng chung
AUTHOR_INFO_MARKDOWN = """
**Tác giả:**
* **1. Đàm Thuận Minh Bình**  
  📞 0387.136.888
* **2. Đỗ Thị Huyền**  
  📞 0982.036.952

🏫 *Trường THCS Thân Nhân Trung - TP. Bắc Ninh*
"""

# HÀM TẠO FILE DOCX CHUẨN THỂ THỨC (TIMES NEW ROMAN, CỠ 13PT, LỀ TRÁI 3CM, CÒN LẠI 2CM)
def generate_docx_report(student_name, date_str, topic, essay_text, feedback_md):
    doc = Document()
    
    # Thiết lập lề trang chuẩn hành chính: Trái 3.0cm, Trên/Dưới/Phải 2.0cm
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
      
    # Tiêu đề
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(12)
    r_t = p_title.add_run("PHIẾU ĐÁNH GIÁ & NHẬN XÉT BÀI THI ESSAY")
    r_t.bold = True
    r_t.font.size = Pt(15)
    r_t.font.color.rgb = RGBColor(0x0b, 0x3c, 0x5d)
    
    # Thông tin bài nộp
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.line_spacing = 1.25
    p_meta.add_run(f"• Học sinh: ").bold = True
    p_meta.add_run(f"{student_name}\n")
    p_meta.add_run(f"• Thời gian nộp bài: ").bold = True
    p_meta.add_run(f"{date_str}\n")
    p_meta.add_run(f"• Đề thi: ").bold = True
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
    p_essay.add_run(essay_text if essay_text else "(Bài nộp qua hình ảnh viết tay)")
    
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
    
    r_tname = p_s.add_run("Cô Đỗ Thị Huyền\n")
    r_tname.bold = True
    r_tname.font.size = Pt(13)
    
    r_tsch = p_s.add_run("Trường THCS Thân Nhân Trung\n📞 0982.036.952")
    r_tsch.font.size = Pt(12)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 2. Cơ sở dữ liệu SQLite
def init_db():
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT,
            fullname TEXT,
            role TEXT
        )
    ''')
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
    c.execute("SELECT * FROM users WHERE username = 'giaovien'")
    if not c.fetchone():
        c.execute("INSERT INTO users VALUES ('giaovien', 'gv123456', 'Đỗ Thị Huyền', 'teacher')")
        c.execute("INSERT INTO users VALUES ('hs01', '123456', 'Nguyễn Văn An', 'student')")
        c.execute("INSERT INTO users VALUES ('hs02', '123456', 'Trần Thị Bình', 'student')")
        c.execute("INSERT INTO users VALUES ('hs03', '123456', 'Lê Hoàng Long', 'student')")
    conn.commit()
    conn.close()

init_db()

# 3. Ngân hàng đề thi
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
You are an extremely strict, uncompromising, and highly authoritative chief examiner for the English Gifted Student Examination (Kỳ thi Chọn Học sinh Giỏi Tỉnh & Chuyên Anh lớp 9) in Bac Ninh Province, Vietnam.

Your absolute priority is to enforce iron discipline in grading. Gifted students must face real examination rigor: do NOT award inflated, sympathetic, or default median scores.

============================================================
I. CRITICAL SCORING CEILINGS & DISQUALIFYING ERRORS:

1. TASK DRIFT / OFF-TOPIC / TANGENTIAL RESPONSE (LỆCH TRỌNG TÂM CÂU HỎI):
   - Definition: Failing to address the exact prompt qualifiers (e.g., Prompt asks about "THE BEST WAY", but the student writes about general pros/cons of social media; Prompt asks about "STRESSFUL", but the student only describes electronic devices).
   - HARD CEILING PENALTY:
     * Completely Off-topic: Content = 0.00 to 0.10 / 0.70.
     * Task Drift / Tangential (Lệch trọng tâm / Lạc đề một phần): Content MUST BE CAPPED at 0.15 to 0.25 / 0.70. NO EXCEPTIONS. Do not give 0.35+ or 0.40+ if the student fails to answer the core debate.
     * Associated Penalty: If Content is penalised for Task Drift, Organization MUST ALSO be capped at 0.25 / 0.60 because the overall line of reasoning is invalid.

2. SEVEN RED FLAGS TO PENALIZE HEAVILY:
   - Idea Dumping: Listing 3-5 ideas without mechanism -> Deduct Content down to 0.30 - 0.35 max.
   - Examples Without Analysis: Using examples to replace logical explanation -> Deduct Content.
   - Repetition: Circling back to the same argument -> Deduct Organization.
   - Memorised / Generic Templates: Pre-learned robotic shells -> Deduct Organization & Language.
   - Overclaiming: Unhedged claims ("always", "completely", "never") -> Deduct Language & Content.
   - Fake Sophistication & Collocation Hallucination (e.g., "in a blaze of people", "social media mechanism offer") -> Penalize Language strictly down to 0.20 - 0.30 / 0.60.
   - Mechanics: Academic writing forbidding informal contractions ("don't", "isn't", "can't"). Deduct Mechanics immediately down to 0.00 - 0.04 / 0.10 if contractions or basic misspellings exist.

============================================================
II. ESSAY TYPE DECODING RULES:
A. DISCUSSIVE ESSAY: Examine + Evaluate + Arrive at a reasoned judgement. Body: CLAIM -> WHY -> HOW -> EXAMPLE -> EVALUATION.
B. ARGUMENTATIVE ESSAY: Claim + Support + Defend. Counterargument: CONCEDE -> QUALIFY -> REBUT.

============================================================
III. OFFICIAL BAC NINH 2.0-POINT RUBRIC (STRICT SCORING):
1. Content (0.70 pt max)
2. Organization & Presentation (0.60 pt max)
3. Language (0.60 pt max)
4. Mechanics (0.10 pt max)

============================================================
IV. REQUIRED OUTPUT FORMAT:

### 1. 📋 ĐÁNH GIÁ TỔNG QUAN & DẠNG BÀI
- **Thể loại bài viết:** [Discussive Essay / Argumentative Essay]
- **Kiểm định Trọng tâm đề thi (Task Response Audit):** [Trúng đề / Lệch trọng tâm / Lạc đề hoàn toàn - Nêu rõ lý do đối chiếu với từ khóa cốt lõi của đề].
- **Số lượng từ:** [Số từ] từ (Chuẩn đề: 200–250 từ).
- **Soi xét Lịch sử cá nhân hóa:** [Nhận xét học sinh có tái phạm các lỗi đã mắc ở các bài trước hay đã có cải thiện cụ thể nào].

### 2. 📊 BẢNG ĐIỂM CHÍNH THỨC SỞ GD&ĐT BẮC NINH (THANG 2.0)
| Tiêu chí thành phần | Điểm tối đa | Điểm đạt | Nhận xét chi tiết của Giám khảo |
| :--- | :---: | :---: | :--- |
| **1. Content** (Ý tưởng & Lập luận) | 0.70 | **...** | Đánh giá tính trúng đề; phạt trần điểm nghiêm ngặt nếu Task drift/Idea dumping. |
| **2. Organization** (Bố cục & Mạch lạc) | 0.60 | **...** | Đánh giá cấu trúc 4 đoạn chuẩn; tính liên kết logic, trừ điểm nếu lập luận gãy khúc. |
| **3. Language** (Từ vựng & Ngữ pháp) | 0.60 | **...** | Bắt lỗi collocation tự chế, fake sophistication, cấu trúc câu gượng ép. |
| **4. Mechanics** (Chính tả & Thể thức) | 0.10 | **...** | Trừ thẳng tay nếu có từ viết tắt (don't, isn't) hoặc sai chính tả. |
| **TỔNG ĐIỂM BÀI THI** | **2.00** | **... / 2.0** | **Ước lượng band IELTS tương đương: ...** |

### 3. 🔍 SOI LỖI LẬP LUẬN THEO CHUYÊN ĐỀ TẬP HUẤN
- **Bảng phân tích câu văn chi tiết:**
| Câu văn gốc của học sinh | Lỗi sai (Tư duy / Ngữ pháp / Collocation) | Đề xuất sửa chữa nâng cao |
|---|---|---|

### 4. 💎 NÂNG CẤP TỪ VỰNG & NGỮ PHÁP THEN CHỐT
- 4–5 cụm collocations hữu ích sửa chữa đúng trọng tâm của đề.

### 5. ✍️ BÀI VIẾT MẪU THAM KHẢO THEO 2 CẤP ĐỘ (200–230 TỪ)

#### 🔹 Cấp độ 1: Bản Nền tảng & Dễ tiếp thu (Mức độ B1 đến B1+ - Mọi học sinh đều học và nhớ được)
- **Đặc điểm:** Dùng câu văn ngắn gọn, từ ngữ quen thuộc, ngữ pháp đơn giản nhưng tuyệt đối chuẩn xác (không sai thì, không sai chia động từ), liên kết tự nhiên (Firstly, In addition, However, As a result), lập luận bám sát đề thi để học sinh dễ hiểu, dễ ghi nhớ vào phòng thi.
[Viết toàn bài essay mẫu hoàn chỉnh Cấp độ B1-B1+ tại đây]

#### 🔸 Cấp độ 2: Bản Nâng cao & Bứt phá điểm số (Học thuật C1-C2 - Dành cho đội tuyển chuyên sâu)
- **Đặc điểm:** Văn phong trang trọng, sử dụng các cụm collocation đắt giá, kỹ thuật Hedging, liên kết câu phức hợp và phân tích cơ chế đa chiều để đạt điểm tuyệt đối về tiêu chí Language và Content.
[Viết toàn bài essay mẫu hoàn chỉnh Cấp độ C1-C2 tại đây]

### 6. ⚠️ DANH SÁCH LỖI THEN CHỐT CẦN LƯU HỒ SƠ:
(Ghi 1-3 lỗi cốt lõi ngắn gọn để ghi vào CSDL theo dõi cá nhân).
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
    c.execute("SELECT username, fullname, role FROM users WHERE username = ? AND password = ?", (username.strip(), password.strip()))
    user_record = c.fetchone()
    conn.close()
    if user_record:
        st.session_state.logged_in = True
        st.session_state.user = {"username": user_record[0], "fullname": user_record[1], "role": user_record[2]}
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
st.sidebar.caption(f"Vai trò: {'Giáo viên quản trị' if user['role'] == 'teacher' else 'Học sinh đội tuyển'}")

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
    
    with tab_submit:
        selected_topic = st.selectbox("📌 Chọn đề thi từ ngân hàng đề:", DE_THI_BAC_NINH)
        if selected_topic == "-- Tự nhập đề bài mới --":
            essay_prompt = st.text_area("Nhập đề bài luận:", placeholder="Nhập đề bài tại đây...")
        else:
            essay_prompt = selected_topic
            
        sub_tab1, sub_tab2 = st.tabs(["📄 Dán văn bản", "📷 Tải ảnh bài viết tay"])
        essay_text = ""
        uploaded_image = None
        
        with sub_tab1:
            essay_text = st.text_area("Nội dung bài viết:", height=250, placeholder="Gõ hoặc dán toàn bộ bài làm của em tại đây...")
            if essay_text:
                st.write(f"📏 Số từ: **{len(essay_text.split())} từ** (Độ dài chuẩn: 200–250 từ)")
                
        with sub_tab2:
            uploaded_file = st.file_uploader("Tải lên ảnh bài viết tay:", type=["png", "jpg", "jpeg"])
            if uploaded_file:
                uploaded_image = Image.open(uploaded_file)
                st.image(uploaded_image, caption="Bài làm viết tay", use_container_width=True)
                
        if st.button("🚀 Nộp bài & Chấm điểm ngay", type="primary"):
            if not essay_prompt.strip():
                st.error("⚠️ Vui lòng nhập hoặc chọn đề thi!")
            elif not essay_text.strip() and not uploaded_image:
                st.error("⚠️ Vui lòng dán nội dung bài hoặc tải ảnh lên!")
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
                    if uploaded_image:
                        user_content.append(uploaded_image)

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
                        ''', (user["username"], essay_prompt, essay_text, 1.5, 0.5, 0.45, 0.45, 0.1, result_text, "Task Drift, Overclaiming", now_str))
                        conn.commit()
                        conn.close()
                    else:
                        st.error(f"Hệ thống gặp sự cố khi chấm bài. Chi tiết: {last_err}")

        # KHU VỰC HIỂN THỊ KẾT QUẢ VÀ NÚT TẢI FILE CỐ ĐỊNH
        if st.session_state.last_graded_result:
            res = st.session_state.last_graded_result
            st.success("✅ ĐÃ CHẤM XONG BÀI THI!")
            
            # Nút tải file Word đặt nổi bật ở đầu
            try:
                docx_bytes = generate_docx_report(user['fullname'], res['date_str'], res['topic'], res['essay_text'], res['result_text'])
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
            st.markdown("""
            ---
            ### ✍️ GIÁO VIÊN BỒI DƯỠNG & CHẤM ĐIỂM
            **Cô Đỗ Thị Huyền**  
            Trường THCS Thân Nhân Trung - TP. Bắc Ninh  
            📞 Số điện thoại: 0982.036.952
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
                        docx_data = generate_docx_report(user['fullname'], r[2], r[1], r[4], r[3])
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
# GIAO DIỆN GIÁO VIÊN (DASHBOARD QUẢN TRỊ)
# =========================================================================
elif user["role"] == "teacher":
    st.title("👨‍🏫 Bảng Điều Khiển Quản Trị Giáo Viên")
    t_tab1, t_tab2, t_tab3 = st.tabs(["📊 Tổng hợp kết quả cả lớp", "🔍 Xem bài & Xoá bài nộp", "👥 Quản lý & Xoá học sinh"])
    
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    
    with t_tab1:
        st.markdown("### 📌 Báo cáo tổng thể đội tuyển HSG")
        c.execute('''
            SELECT s.id, u.fullname, s.topic, s.created_at, s.identified_errors
            FROM submissions s JOIN users u ON s.username = u.username
            ORDER BY s.id DESC
        ''')
        submissions = c.fetchall()
        
        if not submissions:
            st.info("Hiện tại chưa có bài nộp nào trong hệ thống.")
        else:
            st.write(f"Tổng số lượt nộp bài toàn đội tuyển: **{len(submissions)} lượt**")
            table_data = []
            for sub in submissions:
                table_data.append({
                    "Mã bài": sub[0],
                    "Học sinh": sub[1],
                    "Đề bài": sub[2][:50] + "...",
                    "Thời gian": sub[3],
                    "Lỗi trọng tâm cần sửa": sub[4]
                })
            st.table(table_data)
            
            st.markdown("---")
            st.markdown("### 💡 Gợi ý Chữa bài chung trên lớp (AI Teacher Assistant):")
            st.warning("""
            **Các nhược điểm học sinh hay mắc nhiều nhất:**
            1. **Lệch trọng tâm (Task Drift):** Bỏ quên từ khóa so sánh nhất hoặc từ khóa điều kiện của đề.
            2. **Thiếu Mechanism:** Mới nêu Claim đã vội đưa ví dụ, chưa giải thích chuỗi nguyên nhân - hệ quả (Why/How).
            3. **Overclaiming:** Khẳng định tuyệt đối, thiếu ngôn ngữ học thuật chừng mực (Hedging).
            """)
            
    with t_tab2:
        st.markdown("### 🔍 Thẩm định bài làm & Xoá bài nộp")
        c.execute("SELECT s.id, u.fullname, s.topic, s.created_at, s.feedback, s.essay_text FROM submissions s JOIN users u ON s.username = u.username ORDER BY s.id DESC")
        all_subs = c.fetchall()
        
        if all_subs:
            sub_options = {f"[{sub[0]}] {sub[1]} - {sub[2][:40]}... ({sub[3]})": sub for sub in all_subs}
            chosen = st.selectbox("Chọn bài nộp cần xem hoặc xoá:", list(sub_options.keys()))
            selected_sub = sub_options[chosen]
            
            col_info, col_del = st.columns([4, 1])
            with col_info:
                st.markdown(f"#### 👤 Học sinh: **{selected_sub[1]}** | Ngày nộp: **{selected_sub[3]}**")
                st.info(f"**Đề bài:** {selected_sub[2]}")
            with col_del:
                st.write("")
                st.write("")
                if st.button("🗑️ Xoá bài này", type="secondary", use_container_width=True):
                    c.execute("DELETE FROM submissions WHERE id = ?", (selected_sub[0],))
                    conn.commit()
                    st.success(f"Đã xoá thành công bài nộp mã #{selected_sub[0]}!")
                    st.rerun()
            
            # Nút tải file Word dành riêng cho Giáo viên
            try:
                t_docx = generate_docx_report(selected_sub[1], selected_sub[3], selected_sub[2], selected_sub[5], selected_sub[4])
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
                st.text(selected_sub[5])
                
            st.markdown("---")
            st.markdown("### 📝 Kết quả chấm & Nhận xét của AI:")
            st.markdown(selected_sub[4])
        else:
            st.info("Không có bài nộp nào để hiển thị.")

    # TAB 3: QUẢN LÝ & CẤP TÀI KHOẢN HỌC SINH
    with t_tab3:
        st.markdown("### 📂 Cấp tài khoản hàng loạt từ file Excel")
        st.caption("File Excel chỉ cần có 2 cột bất kỳ (Cột 1 là Mã học sinh/Username, Cột 2 là Họ và tên học sinh). Mật khẩu mặc định tạo ra sẽ là: 123456")
        
        uploaded_excel = st.file_uploader("Chọn file Excel danh sách học sinh (.xlsx, .xls):", type=["xlsx", "xls"])
        
        if uploaded_excel is not None:
            try:
                import pandas as pd
                df = pd.read_excel(uploaded_excel)
                if df.shape[1] < 2:
                    st.error("File Excel cần có ít nhất 2 cột: Cột 1 là Mã HS và Cột 2 là Họ tên.")
                else:
                    st.write("👀 **Xem trước dữ liệu trích xuất từ file Excel:**")
                    preview_df = df.iloc[:, :2].dropna()
                    preview_df.columns = ["Mã học sinh", "Họ và tên"]
                    st.dataframe(preview_df.head(10), use_container_width=True)
                    
                    if st.button("🚀 Xác nhận tạo tài khoản cho toàn bộ danh sách trên", type="primary"):
                        created_count = 0
                        skipped_count = 0
                        
                        for _, row in preview_df.iterrows():
                            u_code = str(row["Mã học sinh"]).strip()
                            # Loại bỏ số thập phân nếu mã bị định dạng dạng số thực trong Excel (ví dụ 101.0 -> 101)
                            if u_code.endswith(".0"):
                                u_code = u_code[:-2]
                            fullname = str(row["Họ và tên"]).strip()
                            
                            if u_code and fullname and u_code != "nan" and fullname != "nan":
                                try:
                                    c.execute("INSERT INTO users VALUES (?, '123456', ?, 'student')", (u_code, fullname))
                                    created_count += 1
                                except sqlite3.IntegrityError:
                                    # Trùng username đã có trong database
                                    skipped_count += 1
                                    
                        conn.commit()
                        st.success(f"🎉 Hoàn tất! Đã thêm thành công **{created_count}** tài khoản mới. (Bỏ qua {skipped_count} tài khoản bị trùng lặp mã).")
                        st.rerun()
            except Exception as ex:
                st.error(f"Lỗi khi đọc file Excel: {str(ex)}")

        st.markdown("---")
        
        col_add, col_remove = st.columns(2)
        with col_add:
            st.markdown("### ➕ Thêm thủ công từng học sinh")
            with st.form("add_user_form"):
                new_u = st.text_input("Tên đăng nhập (Username):", placeholder="Ví dụ: hs04")
                new_p = st.text_input("Mật khẩu ban đầu:", value="123456")
                new_name = st.text_input("Họ và tên học sinh:", placeholder="Ví dụ: Hoàng Minh Đức")
                submit_btn = st.form_submit_button("Thêm học sinh này", type="primary")
                if submit_btn:
                    if new_u and new_p and new_name:
                        try:
                            c.execute("INSERT INTO users VALUES (?, ?, ?, 'student')", (new_u.strip(), new_p.strip(), new_name.strip()))
                            conn.commit()
                            st.success(f"Đã tạo tài khoản cho **{new_name}**!")
                            st.rerun()
                        except Exception:
                            st.error("Tên đăng nhập này đã tồn tại!")
                    else:
                        st.warning("Vui lòng nhập đầy đủ thông tin.")
                        
        with col_remove:
            st.markdown("### ❌ Xoá tài khoản Học sinh")
            c.execute("SELECT username, fullname FROM users WHERE role = 'student'")
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
                    st.success(f"Đã xoá học sinh {target_student} và toàn bộ bài nộp liên quan!")
                    st.rerun()
            else:
                st.info("Chưa có học sinh nào trong danh sách.")

        st.markdown("---")
        st.markdown("#### 📋 Danh sách tài khoản hiện tại:")
        c.execute("SELECT username as 'Tên đăng nhập', fullname as 'Họ và tên', role as 'Vai trò' FROM users")
        current_users = c.fetchall()
        st.table([{"Tên đăng nhập": u[0], "Họ và tên": u[1], "Vai trò": "Giáo viên" if u[2] == "teacher" else "Học sinh"} for u in current_users])
        
    conn.close()
