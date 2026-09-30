import streamlit as st
import psycopg2
from psycopg2 import pool
from datetime import datetime
from zoneinfo import ZoneInfo
import time
from google import genai
from google.genai import types
from PIL import Image
import io
import re
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Hàm lấy thời gian chuẩn múi giờ Việt Nam GMT+7
def get_vn_time_str(fmt="%d/%m/%Y %H:%M"):
    return datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).strftime(fmt)

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

# HÀM BÓC TÁCH ĐIỂM THỰC TẾ VÀ LỖI THỰC TẾ TỪ PHẢN HỒI CỦA GIÁM KHẢO AI
def parse_scores_from_feedback(text):
    s_content = 0.45
    s_org = 0.40
    s_lang = 0.40
    s_mech = 0.08
    s_total = 1.33
    errors_str = "Chưa ghi nhận lỗi nghiêm trọng"
    
    if not text:
        return s_content, s_org, s_lang, s_mech, s_total, errors_str
        
    try:
        m_c = re.search(r'Content.*?(?:0\.70|0\.7)\s*\|\s*[\*_`]*([0-9.]+)', text, re.IGNORECASE)
        if m_c:
            s_content = float(m_c.group(1))
        
        m_o = re.search(r'Organization.*?(?:0\.60|0\.6)\s*\|\s*[\*_`]*([0-9.]+)', text, re.IGNORECASE)
        if m_o:
            s_org = float(m_o.group(1))
        
        m_l = re.search(r'Language.*?(?:0\.60|0\.6)\s*\|\s*[\*_`]*([0-9.]+)', text, re.IGNORECASE)
        if m_l:
            s_lang = float(m_l.group(1))
        
        m_m = re.search(r'Mechanics.*?(?:0\.10|0\.1)\s*\|\s*[\*_`]*([0-9.]+)', text, re.IGNORECASE)
        if m_m:
            s_mech = float(m_m.group(1))
        
        m_t = re.search(r'TỔNG ĐIỂM BÀI THI.*?(?:2\.00|2\.0)\s*\|\s*[\*_`]*([0-9.]+)', text, re.IGNORECASE)
        if m_t:
            s_total = float(m_t.group(1))
        else:
            s_total = round(max(0.0, s_content + s_org + s_lang + s_mech), 2)
            
        m_err = re.search(r'(?:6\.\s*⚠\s*DANH SÁCH LỖI THEN CHỐT CẦN LƯU HỒ SƠ|DANH SÁCH LỖI THEN CHỐT)[:\s*\n]+(.*?)(?:\n###|\Z)', text, re.DOTALL | re.IGNORECASE)
        if m_err:
            raw_err = m_err.group(1).strip()
            cleaned_lines = [re.sub(r'^[\s*\-0-9.)]+', '', line).strip() for line in raw_err.split('\n') if line.strip()]
            if cleaned_lines:
                errors_str = " | ".join(cleaned_lines[:3])
                if len(errors_str) > 120:
                    errors_str = errors_str[:117] + "..."
    except Exception:
        pass
        
    return s_content, s_org, s_lang, s_mech, s_total, errors_str

# HÀM TẠO FILE DOCX CHUẨN THỂ THỨC (TIMES NEW ROMAN, CỠ 13PT, LỀ TRÁI 3CM, CÒN LẠI 2CM)
def generate_docx_report(student_name, date_str, topic, essay_text, feedback_md, teacher_name="Cô Đỗ Thị Huyền"):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)     # 2.0 cm
        section.bottom_margin = Inches(0.79)  # 2.0 cm
        section.left_margin = Inches(1.18)    # 3.0 cm
        section.right_margin = Inches(0.79)   # 2.0 cm
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(14)
    r_t = p_title.add_run("PHIẾU ĐÁNH GIÁ & NHẬN XÉT BÀI THI ESSAY")
    r_t.bold = True
    r_t.font.size = Pt(16)
    r_t.font.color.rgb = RGBColor(0x0b, 0x3c, 0x5d)
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.line_spacing = 1.25
    p_meta.add_run("• Học sinh: ").bold = True
    p_meta.add_run(f"{student_name}\n")
    p_meta.add_run("• Thời gian nộp bài: ").bold = True
    p_meta.add_run(f"{date_str}\n")
    p_meta.add_run("• Đề thi: ").bold = True
    p_meta.add_run(f"{topic}\n")
    
    h1 = doc.add_heading("I. NỘI DUNG BÀI LÀM CỦA HỌC SINH", level=2)
    for r in h1.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13.5)
        r.font.color.rgb = RGBColor(0x0b, 0x3c, 0x5d)
        
    p_essay = doc.add_paragraph()
    p_essay.paragraph_format.left_indent = Inches(0.2)
    p_essay.paragraph_format.line_spacing = 1.25
    p_essay.add_run(essay_text if essay_text else "(Bài làm đính kèm dạng hình ảnh/PDF viết tay)")
    
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

# 2. Quản lý Kết nối Cơ sở Dữ liệu Đám mây Supabase (PostgreSQL)
def get_db_connection():
    db_url = st.secrets.get("DATABASE_URL")
    if not db_url:
        st.error("Chưa cấu hình biến DATABASE_URL trong Streamlit Secrets! Vui lòng cấu hình để tiếp tục.")
        st.stop()
    return psycopg2.connect(db_url)

SUPER_ADMIN_USERS = ["giaovien", "gv_binh"]

# 3. Huấn luyện System Instruction chuẩn Barem 2.0 Bắc Ninh
SYSTEM_INSTRUCTION = """
You are an authoritative chief examiner for the English Gifted Student Examination (Kỳ thi Chọn Học sinh Giỏi Tỉnh & Chuyên Anh lớp 9) in Bac Ninh Province, Vietnam.

Your core grading philosophy:
1. EVALUATE STRONG STATEMENTS RIGOROUSLY:
   - Candidates MUST evaluate the truth, validity, degree, and boundaries of strong/extreme claims, rather than just listing generic pros/cons.

2. SUBSTANCE OVER SHOWMANSHIP (TRỪ NẶNG LỖI TỪ VỰNG KHỦNG NHƯNG Ý NÔNG):
   - A high-scoring essay MUST have:
     * A clear, consistent thesis maintained throughout.
     * Exactly 2 well-developed main arguments with deep causal mechanisms (Claim -> Why -> How -> Concrete Evidence -> Counterargument/Hedging).
     * Tight logical transitions and organic cohesion.
   - Plain, natural, precise, and academically sound language is vastly superior to forced, unnatural vocabulary.

3. PENALTY RULE FOR ESSAY LENGTH (QUY ĐỊNH TRỪ ĐIỂM PHẠT THIẾU TỪ TRÊN TỔNG ĐIỂM):
   - Độ dài tiêu chuẩn: khoảng 250 từ (Standard length: ~250 words).
   - Điểm phạt thiếu từ được TRỪ TRỰC TIẾP VÀO TỔNG ĐIỂM SAU KHI ĐÃ CỘNG 4 TIÊU CHÍ THÀNH PHẦN:
     * Thiếu dưới 10% (225 – 249 từ): Không trừ hoặc trừ 0.10 điểm (ghi rõ lý do nếu trừ).
     * Thiếu 10% – 25% (190 – 224 từ): Trừ thẳng 0.20 – 0.30 điểm tổng.
     * Thiếu trên 25% (< 190 từ): Trừ nặng 0.40 – 0.50 điểm tổng.

4. STRICT FORMATTING RULE FOR TABLES (QUY TẮC HIỂN THỊ BẢNG BIỂU):
   - NEVER put a bullet point (* or -) directly before a Markdown table header.
   - Always leave at least one blank newline before and after every table so that it renders as a clean visual grid on both web and Word documents.

5. BILINGUAL EXPLANATIONS (DỊCH NGHĨA TIẾNG VIỆT ĐẦY ĐỦ):
   - Every suggested correction, upgrade, vocabulary item, and phrase MUST include a clear Vietnamese translation (dịch nghĩa tiếng Việt) so students can readily understand and memorize.

============================================================
OFFICIAL BAC NINH 2.0-POINT RUBRIC:
1. Content (0.70 max): Strict alignment with statement nuances, fully developed mechanisms, no task drift.
2. Organization & Presentation (0.60 max): Organic 4-paragraph structure, tight line of reasoning, natural cohesive flow.
3. Language (0.60 max): Accuracy, clarity, natural collocations, academic hedging. Penalize forced/hallucinated vocabulary.
4. Mechanics (0.10 max): Punctuation, spelling, capitalisation, zero contractions.
* PENALTY DEDUCTION (Trừ điểm phạt độ dài): Áp dụng sau khi cộng 4 tiêu chí trên.

============================================================
REQUIRED OUTPUT FORMAT:

### 1. 📋 ĐÁNH GIÁ TỔNG QUAN & PHÂN TÍCH NHẬN ĐỊNH CỦA ĐỀ
- **Thể loại bài viết nhận diện:** [Opinion / Discussion / Cause-Solution / Advantages-Disadvantages / Two-Part Question]
- **Kiểm định Phản hồi Nhận định mạnh (Evaluating the Prompt's Statement):** [Đánh giá thí sinh có phản biện được tính tuyệt đối/mức độ đúng của nhận định hay chỉ liệt kê ưu/nhược điểm chung chung].
- **Số lượng từ bài làm:** [Số từ] từ (Chuẩn: ~250 từ).
- **Mức phạt độ dài (nếu có):** [Không bị phạt / Trừ ... điểm vì thiếu ... từ theo quy chế].
- **Soi xét Lịch sử cá nhân hóa:** [Nhận xét học sinh có tái phạm các lỗi cũ hay đã có tiến bộ cụ thể nào].

### 2. 📊 BẢNG ĐIỂM CHÍNH THỨC SỞ GD&ĐT BẮC NINH (THANG 2.0)

| Tiêu chí thành phần | Điểm tối đa | Điểm đạt | Nhận xét chi tiết của Giám khảo |
| :--- | :---: | :---: | :--- |
| **1. Content** (Ý tưởng & Lập luận) | 0.70 | **...** | Đánh giá độ sâu lập luận (2 luận điểm phát triển sâu); phạt nếu ý nông hoặc Task Drift. |
| **2. Organization** (Bố cục & Mạch lạc) | 0.60 | **...** | Đánh giá tính nhất quán của quan điểm xuyên suốt và liên kết logic tự nhiên. |
| **3. Language** (Từ vựng & Ngữ pháp) | 0.60 | **...** | Đánh giá độ chuẩn xác, tự nhiên; trừ điểm nếu sính "từ vựng khủng" nhưng gượng ép. |
| **4. Mechanics** (Chính tả & Thể thức) | 0.10 | **...** | Trừ thẳng tay nếu có từ viết tắt (don't, isn't) hoặc sai chính tả. |
| **5. Điểm phạt thiếu từ (Penalty)** | -0.50 | **...** | Trừ trực tiếp: [0.00 / -0.10 / -0.20 đến -0.50 tuỳ số từ thực tế]. |
| **TỔNG ĐIỂM BÀI THI** | **2.00** | **... / 2.0** | **Ước lượng band IELTS tương đương: ...** |

*(Ghi chú: Tổng điểm = (Content + Organization + Language + Mechanics) - Điểm phạt thiếu từ).*

### 3. 🔍 SOI LỖI CHI TIẾT (KÈM DỊCH NGHĨA TIẾNG VIỆT ĐỀ XUẤT)

| Câu văn gốc của học sinh | Lỗi sai (Ngữ pháp / Collocation / Sính từ) | Đề xuất sửa chuẩn xác & Dịch nghĩa tiếng Việt |
|---|---|---|
| (Câu gốc của HS) | (Chỉ rõ lỗi) | **English:** [Câu sửa chuẩn tự nhiên]<br>👉 **Dịch nghĩa:** [Bản dịch tiếng Việt tương ứng] |

### 4. 💎 NÂNG CẤP TỪ VỰNG & DIỄN ĐẠT THEN CHỐT (KÈM PHIÊN ÂM & NGHĨA)
1. **[Cụm từ tiếng Anh]** /phiên âm IPA/: Dịch nghĩa tiếng Việt và cách dùng trong chủ đề này.
2. **[Cụm từ tiếng Anh]** /phiên âm IPA/: Dịch nghĩa tiếng Việt và cách dùng trong chủ đề này.
3. **[Cụm từ tiếng Anh]** /phiên âm IPA/: Dịch nghĩa tiếng Việt và cách dùng trong chủ đề này.
4. **[Cụm từ tiếng Anh]** /phiên âm IPA/: Dịch nghĩa tiếng Việt và cách dùng trong chủ đề này.
5. **[Cụm từ tiếng Anh]** /phiên âm IPA/: Dịch nghĩa tiếng Việt và cách dùng trong chủ đề này.

### 5. ✍️ BÀI VIẾT MẪU THAM KHẢO THEO 2 CẤP ĐỘ (CHUẨN ~250 TỪ)

#### 🔹 Cấp độ 1: Bản Nền tảng & Dễ tiếp thu (Mức độ B1 đến B1+ - Mọi học sinh đều học và nhớ được)
- **Đặc điểm:** Bố cục chuẩn mực, diễn đạt sáng rõ, ngữ pháp tuyệt đối chuẩn, từ vựng quen thuộc nhưng chính xác, 2 ý triển khai có chiều sâu rõ rệt.

[Viết toàn bài essay mẫu hoàn chỉnh Cấp độ B1-B1+ chuẩn khoảng 250 từ tại đây]

**Bảng thống kê 10 từ vựng / cụm từ / mẫu câu hữu dụng của Cấp độ 1:**

| STT | Từ vựng / Cụm từ / Mẫu câu | Phiên âm quốc tế (IPA) | Dịch nghĩa tiếng Việt & Ngữ cảnh áp dụng |
|:---:|---|---|---|
| 1 | ... | /.../ | ... |
| 2 | ... | /.../ | ... |
| 3 | ... | /.../ | ... |
| 4 | ... | /.../ | ... |
| 5 | ... | /.../ | ... |
| 6 | ... | /.../ | ... |
| 7 | ... | /.../ | ... |
| 8 | ... | /.../ | ... |
| 9 | ... | /.../ | ... |
| 10 | ... | /.../ | ... |

---

#### 🔸 Cấp độ 2: Bản Nâng cao & Bứt phá điểm số (Học thuật C1-C2 - Dành cho đội tuyển chuyên sâu)
- **Đặc điểm:** Lập luận sắc sảo, kỹ thuật Hedging để đánh giá nhận định đa chiều, kết nối mượt mà, từ vựng tự nhiên và chuẩn văn phong học thuật cao cấp.

[Viết toàn bài essay mẫu hoàn chỉnh Cấp độ C1-C2 chuẩn khoảng 250 từ tại đây]

**Bảng thống kê 10 từ vựng / collocations / cấu trúc học thuật tinh hoa của Cấp độ 2:**

| STT | Từ vựng / Collocation / Cấu trúc | Phiên âm quốc tế (IPA) | Dịch nghĩa tiếng Việt & Giá trị biểu đạt học thuật |
|:---:|---|---|---|
| 1 | ... | /.../ | ... |
| 2 | ... | /.../ | ... |
| 3 | ... | /.../ | ... |
| 4 | ... | /.../ | ... |
| 5 | ... | /.../ | ... |
| 6 | ... | /.../ | ... |
| 7 | ... | /.../ | ... |
| 8 | ... | /.../ | ... |
| 9 | ... | /.../ | ... |
| 10 | ... | /.../ | ... |

### 6. ⚠️ DANH SÁCH LỖI THEN CHỐT CẦN LƯU HỒ SƠ:
(Ghi 1-3 lỗi cốt lõi ngắn gọn để ghi nhớ vào CSDL theo dõi cá nhân).
"""

# Quản lý Đăng nhập
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None
if "last_graded_result" not in st.session_state:
    st.session_state.last_graded_result = None

def login(username, password):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT username, fullname, role, teacher_username FROM users WHERE username = %s AND password = %s", (username.strip(), password.strip()))
    user_record = c.fetchone()
    c.close()
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
is_super_admin = user["username"] in SUPER_ADMIN_USERS

st.sidebar.markdown(f"### 👤 Xin chào: **{user['fullname']}**")
role_label = "Giáo viên Quản trị Trưởng" if is_super_admin else ("Giáo viên phụ trách" if user['role'] == 'teacher' else "Học sinh đội tuyển")
st.sidebar.caption(f"Vai trò: {role_label}")

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
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("SELECT password FROM users WHERE username = %s", (user["username"],))
                curr_db_pw = c.fetchone()
                if curr_db_pw and curr_db_pw[0] == old_pw:
                    c.execute("UPDATE users SET password = %s WHERE username = %s", (new_pw, user["username"]))
                    conn.commit()
                    st.success("Đổi mật khẩu thành công!")
                else:
                    st.error("Mật khẩu hiện tại không đúng!")
                c.close()
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
    
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT fullname FROM users WHERE username = %s", (user.get("teacher_username", "giaovien"),))
    t_row = c.fetchone()
    my_teacher_name = t_row[0] if t_row else "Giáo viên phụ trách"
    
    c.execute("SELECT topic_code, topic_content FROM topics ORDER BY topic_code ASC")
    db_topics = c.fetchall()
    c.close()
    conn.close()

    topic_dict = {f"[{t[0]}] {t[1][:80]}...": f"[{t[0]}] {t[1]}" for t in db_topics}
    topic_options = list(topic_dict.keys())

    with tab_submit:
        if not topic_options:
            st.warning("⚠️ Hiện tại Thầy/Cô chưa mở mã đề thi nào. Em vui lòng quay lại sau!")
            st.stop()
            
        selected_display = st.selectbox("📌 Chọn Mã đề thi Giáo viên đã giao:", topic_options)
        essay_prompt = topic_dict[selected_display]
        st.info(f"**Nội dung đề bài chi tiết ({selected_display.split(']')[0]}]):**\n\n{essay_prompt}")

        # KIỂM TRA DUY NHẤT 1 LẦN: HỌC SINH ĐÃ NỘP MÃ ĐỀ NÀY CHƯA
        conn_check = get_db_connection()
        c_check = conn_check.cursor()
        c_check.execute("SELECT id, created_at, score_total FROM submissions WHERE username = %s AND topic = %s", (user["username"], essay_prompt))
        submitted_record = c_check.fetchone()
        c_check.close()
        conn_check.close()

        has_submitted = submitted_record is not None
        if has_submitted:
            st.warning(f"⚠️ **Thông báo:** Em đã hoàn thành bài thi cho mã đề này vào lúc **{submitted_record[1]}** (Điểm: **{submitted_record[2]}/2.0**). Theo quy định, mỗi đề chỉ được nộp và chấm **1 lần duy nhất**! Em vui lòng sang tab **'Hồ sơ & Lịch sử cá nhân'** để xem lại bài làm.")
            
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
                
        if st.button("🚀 Nộp bài & Chấm điểm ngay", type="primary", disabled=has_submitted):
            if has_submitted:
                st.error("⚠️ Em đã nộp đề thi này rồi, không thể nộp lại!")
                st.stop()
            elif not essay_prompt.strip():
                st.error("⚠️ Vui lòng chọn đề thi!")
            elif not essay_text.strip() and not uploaded_files:
                st.error("⚠️ Vui lòng dán bài viết hoặc tải ảnh/PDF bài làm lên!")
            else:
                with st.spinner("Giám khảo AI đang đối chiếu barem Bắc Ninh và chấm bài..."):
                    conn = get_db_connection()
                    c = conn.cursor()
                    c.execute("SELECT identified_errors FROM submissions WHERE username = %s ORDER BY id DESC LIMIT 3", (user["username"],))
                    past_errors = c.fetchall()
                    c.close()
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

                    valid_api_keys = [k.strip() for k in active_api_keys if k.strip()]

                    if not valid_api_keys:
                        st.error("⚠️ Không tìm thấy API Key nào trong cấu hình Secrets. Vui lòng kiểm tra lại!")
                        st.stop()

                    CANDIDATE_MODELS = [
                        'gemini-3.8-flash',
                        'gemini-3.5-flash',
                        'gemini-3-flash'
                    ]

                    for key in valid_api_keys:
                        if success:
                            break
                        try:
                            client = genai.Client(api_key=key)
                        except Exception as e:
                            last_err = str(e)
                            continue

                        for target_model in CANDIDATE_MODELS:
                            try:
                                response = client.models.generate_content(
                                    model=target_model,
                                    contents=user_content,
                                    config=types.GenerateContentConfig(
                                        system_instruction=SYSTEM_INSTRUCTION,
                                        temperature=0.15
                                    )
                                )
                                if response and response.text:
                                    result_text = response.text
                                    success = True
                                    break
                            except Exception as e:
                                last_err = str(e)
                                if "503" in str(e) or "UNAVAILABLE" in str(e):
                                    time.sleep(1.5)
                                continue

                    if success:
                        now_str = get_vn_time_str("%d/%m/%Y %H:%M")
                        st.session_state.last_graded_result = {
                            "topic": essay_prompt,
                            "essay_text": essay_text,
                            "result_text": result_text,
                            "date_str": now_str
                        }
                        
                        s_c, s_o, s_l, s_m, s_tot, s_err = parse_scores_from_feedback(result_text)
                        
                        conn = get_db_connection()
                        c = conn.cursor()
                        c.execute('''
                            INSERT INTO submissions (username, topic, essay_text, score_total, score_content, score_org, score_lang, score_mech, feedback, identified_errors, created_at)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        ''', (user["username"], essay_prompt, essay_text, s_tot, s_c, s_o, s_l, s_m, result_text, s_err, now_str))
                        conn.commit()
                        c.close()
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
                    file_name=f"Phieu_Nhan_Xet_{user['username']}_{get_vn_time_str('%Y%m%d_%H%M')}.docx",
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
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT id, topic, created_at, feedback, essay_text FROM submissions WHERE username = %s ORDER BY id DESC", (user["username"],))
        rows = c.fetchall()
        c.close()
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
# GIAO DIỆN GIÁO VIÊN
# =========================================================================
elif user["role"] == "teacher":
    st.title(f"👨‍🏫 Bảng Quản Trị Lớp: Thầy/Cô {user['fullname']}")
    
    tab_list = [
        "📊 Bảng điểm Excel & Tổng hợp", 
        "📤 Chấm hàng loạt bài thi giấy",
        "🔍 Xem & Chữa bài chi tiết", 
        "📌 Tạo & Quản lý Mã đề thi",
        "👥 Quản lý học sinh của tôi"
    ]
    if is_super_admin:
        tab_list.append("⚙️ Cấp tài khoản Giáo viên mới")
        
    tabs = st.tabs(tab_list)
    t_tab1 = tabs[0]
    t_tab_batch = tabs[1]
    t_tab2 = tabs[2]
    t_tab_topics = tabs[3]
    t_tab3 = tabs[4]
    
    conn = get_db_connection()
    c = conn.cursor()
    
    # TAB 1: BẢNG ĐIỂM EXCEL & TỔNG HỢP
    with t_tab1:
        st.markdown(f"### 📊 Báo cáo kết quả & Xuất Bảng điểm Excel của lớp")
        
        if is_super_admin:
            c.execute('SELECT DISTINCT topic FROM submissions ORDER BY topic ASC')
        else:
            c.execute('''
                SELECT DISTINCT s.topic 
                FROM submissions s JOIN users u ON s.username = u.username
                WHERE u.teacher_username = %s
                ORDER BY s.topic ASC
            ''', (user["username"],))
            
        topic_rows = c.fetchall()
        
        if not topic_rows:
            st.info("Hiện tại chưa có học sinh nào nộp bài.")
        else:
            topics_list = ["-- Tất cả các đề bài --"] + [t[0] for t in topic_rows]
            chosen_topic = st.selectbox("🎯 Chọn Đề bài để xem hoặc xuất bảng điểm:", topics_list, key="stat_topic_choice")
            
            if chosen_topic == "-- Tất cả các đề bài --":
                if is_super_admin:
                    c.execute('''
                        SELECT s.id, COALESCE(u.fullname, s.username), s.topic, s.score_content, s.score_org, s.score_lang, s.score_mech, s.score_total, s.identified_errors, s.created_at
                        FROM submissions s LEFT JOIN users u ON s.username = u.username
                        ORDER BY s.id DESC
                    ''')
                else:
                    c.execute('''
                        SELECT s.id, u.fullname, s.topic, s.score_content, s.score_org, s.score_lang, s.score_mech, s.score_total, s.identified_errors, s.created_at
                        FROM submissions s JOIN users u ON s.username = u.username
                        WHERE u.teacher_username = %s
                        ORDER BY s.id DESC
                    ''', (user["username"],))
            else:
                if is_super_admin:
                    c.execute('''
                        SELECT s.id, COALESCE(u.fullname, s.username), s.topic, s.score_content, s.score_org, s.score_lang, s.score_mech, s.score_total, s.identified_errors, s.created_at
                        FROM submissions s LEFT JOIN users u ON s.username = u.username
                        WHERE s.topic = %s
                        ORDER BY s.id DESC
                    ''', (chosen_topic,))
                else:
                    c.execute('''
                        SELECT s.id, u.fullname, s.topic, s.score_content, s.score_org, s.score_lang, s.score_mech, s.score_total, s.identified_errors, s.created_at
                        FROM submissions s JOIN users u ON s.username = u.username
                        WHERE s.topic = %s AND u.teacher_username = %s
                        ORDER BY s.id DESC
                    ''', (chosen_topic, user["username"]))
            subs = c.fetchall()
            
            df_export = pd.DataFrame(subs, columns=[
                "Mã bài", "Họ và tên học sinh", "Đề bài", 
                "Content (0.7)", "Org (0.6)", "Lang (0.6)", "Mech (0.1)", 
                "Tổng điểm (2.0)", "Lỗi trọng tâm cần sửa", "Thời gian nộp"
            ])
            
            st.write(f"Số bài nộp: **{len(subs)} bài**")
            st.dataframe(df_export, use_container_width=True)
            
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name="BangDiem_HSG")
            
            clean_topic_code = "All" if chosen_topic == "-- Tất cả các đề bài --" else (chosen_topic.split("]")[0].replace("[", "") if "]" in chosen_topic else "Topic")
            st.download_button(
                label="📥 Tải Bảng Điểm Excel (.xlsx) để lưu trữ",
                data=buffer_excel.getvalue(),
                file_name=f"Bang_Diem_HSG_{clean_topic_code}_{get_vn_time_str('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    # TAB BATCH: GIÁO VIÊN CHẤM HÀNG LOẠT BÀI THI GIẤY THEO MẪU PHIẾU
    with t_tab_batch:
        st.markdown("### 📤 Chấm hàng loạt bài thi tự luận từ ảnh chụp / bản scan phiếu làm bài")
        st.caption("Chức năng dành cho giáo viên chấm tập trung khi thu phiếu thi giấy về. Mỗi file tải lên tương ứng với bài làm của 1 học sinh (file ảnh JPG/PNG hoặc file PDF).")

        c.execute("SELECT topic_code, topic_content FROM topics ORDER BY topic_code ASC")
        batch_topics = c.fetchall()
        
        if not batch_topics:
            st.warning("⚠️ Hiện chưa có mã đề thi nào. Vui lòng tạo mã đề ở tab 'Tạo & Quản lý Mã đề thi' trước!")
        else:
            batch_topic_dict = {f"[{t[0]}] {t[1][:80]}...": f"[{t[0]}] {t[1]}" for t in batch_topics}
            selected_batch_display = st.selectbox("📌 1. Chọn Mã đề thi của xấp bài cần chấm:", list(batch_topic_dict.keys()), key="batch_topic_sel")
            batch_essay_prompt = batch_topic_dict[selected_batch_display]
            st.info(f"**Đề thi áp dụng:** {batch_essay_prompt}")

            uploaded_batch_files = st.file_uploader(
                "📁 2. Chọn toàn bộ ảnh chụp / file PDF phiếu làm bài của học sinh (chọn nhiều file cùng lúc):",
                type=["png", "jpg", "jpeg", "pdf"],
                accept_multiple_files=True,
                key="batch_files_uploader"
            )

            if uploaded_batch_files:
                st.write(f"Đã chọn **{len(uploaded_batch_files)} bài thi** cần chấm.")

                if st.button("🚀 Bắt đầu Chấm tự động toàn bộ xấp bài này", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    valid_api_keys = [k.strip() for k in active_api_keys if k.strip()]
                    CANDIDATE_MODELS = ['gemini-3.8-flash', 'gemini-3.5-flash', 'gemini-3-flash']

                    c.execute("SELECT username, fullname FROM users WHERE role = 'student' AND teacher_username = %s", (user["username"],))
                    my_students_mapping = {s[1].strip().lower(): s[0] for s in c.fetchall()}

                    success_count = 0
                    total_files = len(uploaded_batch_files)

                    for idx, uploaded_file in enumerate(uploaded_batch_files):
                        status_text.info(f"⏳ Đang xử lý bài {idx + 1}/{total_files}: **{uploaded_file.name}**...")
                        
                        file_bytes = uploaded_file.getvalue()
                        user_content_batch = [
                            "DƯỚI ĐÂY LÀ PHIẾU TRẢ LỜI ESSAY CHÍNH THỨC CỦA TRƯỜNG THCS THÂN NHÂN TRUNG.\n",
                            "YÊU CẦU ĐẶC BIỆT:\n",
                            "1. Hãy bóc tách chính xác phần thông tin ở đầu phiếu: 'Họ và tên', 'Lớp', 'Số báo danh' của học sinh.\n",
                            "2. Ở dòng ĐẦU TIÊN CỦA PHẢN HỒI, ghi đúng cú pháp sau để hệ thống nhận diện:\n",
                            "[THÔNG TIN THÍ SINH: Họ và tên: <Tên học sinh> | Lớp: <Lớp> | SBD: <Số báo danh>]\n\n",
                            f"ĐỀ THI: {batch_essay_prompt}\n\n",
                            "3. Đọc kỹ phần chữ viết tay trong mục 'BÀI LÀM CỦA HỌC SINH' và chấm điểm nghiêm ngặt theo barem Bắc Ninh đã được huấn luyện:"
                        ]

                        if uploaded_file.type == "application/pdf":
                            user_content_batch.append(types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"))
                        else:
                            user_content_batch.append(Image.open(io.BytesIO(file_bytes)))

                        batch_graded_text = ""
                        batch_success = False

                        for key in valid_api_keys:
                            if batch_success:
                                break
                            try:
                                client = genai.Client(api_key=key)
                            except Exception:
                                continue

                            for target_model in CANDIDATE_MODELS:
                                try:
                                    response = client.models.generate_content(
                                        model=target_model,
                                        contents=user_content_batch,
                                        config=types.GenerateContentConfig(
                                            system_instruction=SYSTEM_INSTRUCTION,
                                            temperature=0.15
                                        )
                                    )
                                    if response and response.text:
                                        batch_graded_text = response.text
                                        batch_success = True
                                        break
                                except Exception as e:
                                    if "503" in str(e) or "UNAVAILABLE" in str(e):
                                        time.sleep(1.5)
                                    continue

                        if batch_success:
                            extracted_name = "Học sinh chưa rõ tên"
                            name_match = re.search(r'\[THÔNG TIN THÍ SINH:.*?Họ và tên:\s*([^\|\]\n]+)', batch_graded_text, re.IGNORECASE)
                            if name_match:
                                extracted_name = name_match.group(1).strip()
                            else:
                                extracted_name = uploaded_file.name.rsplit('.', 1)[0]

                            matched_username = my_students_mapping.get(extracted_name.lower())
                            if not matched_username:
                                matched_username = f"offline_{re.sub(r'[^a-zA-Z0-9]', '', extracted_name).lower()[:15]}"
                                try:
                                    c.execute("INSERT INTO users (username, password, fullname, role, teacher_username) VALUES (%s, '123456', %s, 'student', %s) ON CONFLICT (username) DO NOTHING", (matched_username, extracted_name, user["username"]))
                                    conn.commit()
                                except Exception:
                                    conn.rollback()

                            sc, so, sl, sm, stot, s_err = parse_scores_from_feedback(batch_graded_text)
                            now_vn = get_vn_time_str("%d/%m/%Y %H:%M")

                            c.execute('''
                                INSERT INTO submissions (username, topic, essay_text, score_total, score_content, score_org, score_lang, score_mech, feedback, identified_errors, created_at)
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                            ''', (matched_username, batch_essay_prompt, f"(Bài chấm giấy từ file {uploaded_file.name})", stot, sc, so, sl, sm, batch_graded_text, s_err, now_vn))
                            conn.commit()

                            success_count += 1
                            st.write(f"✅ **Đã chấm xong ({idx+1}/{total_files}):** Thí sinh **{extracted_name}** — Điểm: **{stot}/2.0**")

                        progress_bar.progress((idx + 1) / total_files)

                    status_text.success(f"🎉 ĐÃ HOÀN TẤT! Chấm thành công **{success_count}/{total_files} bài thi**.")
                    st.info("Thầy/Cô hãy chuyển sang tab **'Bảng điểm Excel & Tổng hợp'** để xem kết quả toàn diện và tải bảng điểm Excel của đợt thi này!")

    # TAB 2: XEM BÀI VÀ XOÁ BÀI
    with t_tab2:
        st.markdown("### 🔍 Thẩm định bài làm & Xoá bài nộp của lớp")
        if is_super_admin:
            c.execute('SELECT DISTINCT topic FROM submissions ORDER BY topic ASC')
        else:
            c.execute('''
                SELECT DISTINCT s.topic 
                FROM submissions s JOIN users u ON s.username = u.username
                WHERE u.teacher_username = %s
                ORDER BY s.topic ASC
            ''', (user["username"],))
        all_topics = [t[0] for t in c.fetchall()]
        
        if all_topics:
            topic_options = ["-- Tất cả các đề bài --"] + all_topics
            selected_topic_filter = st.selectbox("📂 1. Chọn Đề bài:", topic_options, key="view_topic_filter")
            
            if selected_topic_filter == "-- Tất cả các đề bài --":
                if is_super_admin:
                    c.execute('''
                        SELECT s.id, COALESCE(u.fullname, s.username), s.created_at, s.feedback, s.essay_text, s.topic
                        FROM submissions s LEFT JOIN users u ON s.username = u.username
                        ORDER BY s.id DESC
                    ''')
                else:
                    c.execute('''
                        SELECT s.id, u.fullname, s.created_at, s.feedback, s.essay_text, s.topic
                        FROM submissions s JOIN users u ON s.username = u.username
                        WHERE u.teacher_username = %s
                        ORDER BY s.id DESC
                    ''', (user["username"],))
            else:
                if is_super_admin:
                    c.execute('''
                        SELECT s.id, COALESCE(u.fullname, s.username), s.created_at, s.feedback, s.essay_text, s.topic
                        FROM submissions s LEFT JOIN users u ON s.username = u.username
                        WHERE s.topic = %s
                        ORDER BY s.id DESC
                    ''', (selected_topic_filter,))
                else:
                    c.execute('''
                        SELECT s.id, u.fullname, s.created_at, s.feedback, s.essay_text, s.topic
                        FROM submissions s JOIN users u ON s.username = u.username
                        WHERE s.topic = %s AND u.teacher_username = %s
                        ORDER BY s.id DESC
                    ''', (selected_topic_filter, user["username"]))
                
            subs_of_topic = c.fetchall()
            
            if subs_of_topic:
                if selected_topic_filter == "-- Tất cả các đề bài --":
                    sub_dict = {f"Mã #{s[0]} - Học sinh: {s[1]} - Đề: {s[5][:35]}... (Nộp lúc: {s[2]})": s for s in subs_of_topic}
                else:
                    sub_dict = {f"Mã #{s[0]} - Học sinh: {s[1]} (Nộp lúc: {s[2]})": s for s in subs_of_topic}
                    
                chosen_label = st.selectbox("👤 2. Chọn bài nộp của học sinh:", list(sub_dict.keys()))
                selected_sub = sub_dict[chosen_label]
                
                col_info, col_del = st.columns([4, 1])
                with col_info:
                    st.markdown(f"#### 👤 Học sinh: **{selected_sub[1]}** | Ngày nộp: **{selected_sub[2]}**")
                    st.caption(f"Đề bài: {selected_sub[5]}")
                with col_del:
                    st.write("")
                    if st.button("🗑️ Xoá bài này", type="secondary", use_container_width=True):
                        c.execute("DELETE FROM submissions WHERE id = %s", (selected_sub[0],))
                        conn.commit()
                        st.success(f"Đã xoá bài nộp mã #{selected_sub[0]}!")
                        st.rerun()
                
                try:
                    t_docx = generate_docx_report(selected_sub[1], selected_sub[2], selected_sub[5], selected_sub[4], selected_sub[3], teacher_name=user['fullname'])
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
                st.info("Chưa có học sinh nào nộp bài.")
        else:
            st.info("Hiện tại chưa có bài nộp nào trong hệ thống.")

    # TAB 3: TẠO VÀ QUẢN LÝ MÃ ĐỀ THI
    with t_tab_topics:
        st.markdown("### 📌 Tạo Đề thi mới & Gán Mã đề (VD: HSG01, OTC01)")
        st.caption("Mã đề tạo tại đây sẽ hiển thị trong danh mục lựa chọn đề bài của học sinh để các em nộp bài chính xác theo yêu cầu.")
        
        with st.form("create_topic_form"):
            t_code = st.text_input("Mã đề (viết liền, không dấu - ví dụ: HSG03, OTC02, CHUYEN01):", placeholder="HSG03").strip().upper()
            t_content = st.text_area("Nội dung câu hỏi đề thi (kèm yêu cầu độ dài):", placeholder="Ví dụ: 'Some people think that... Write an essay of around 250 words...'")
            btn_create_topic = st.form_submit_button("Lưu & Ban hành Mã đề này", type="primary")
            if btn_create_topic:
                if t_code and t_content:
                    try:
                        now_t = get_vn_time_str("%Y-%m-%d %H:%M")
                        c.execute("INSERT INTO topics (topic_code, topic_content, created_by, created_at) VALUES (%s, %s, %s, %s)", (t_code, t_content.strip(), user["username"], now_t))
                        conn.commit()
                        st.success(f"🎉 Đã lưu thành công Mã đề: **[{t_code}]**!")
                        st.rerun()
                    except psycopg2.IntegrityError:
                        conn.rollback()
                        st.error(f"Mã đề '{t_code}' đã tồn tại! Vui lòng đặt mã khác.")
                else:
                    st.warning("Vui lòng điền đủ Mã đề và Nội dung đề thi!")
                    
        st.markdown("---")
        st.markdown("#### 📋 Danh sách các Mã đề hiện có trong hệ thống:")
        c.execute("SELECT topic_code, topic_content, created_by, created_at FROM topics ORDER BY topic_code ASC")
        current_topics = c.fetchall()
        if current_topics:
            t_table = []
            for t in current_topics:
                t_table.append({
                    "Mã đề": t[0],
                    "Nội dung đề thi": t[1][:70] + "...",
                    "Người tạo": t[2],
                    "Ngày tạo": t[3]
                })
            st.table(t_table)
            
            del_t_code = st.selectbox("Chọn mã đề muốn xoá khỏi danh mục:", [t[0] for t in current_topics], key="del_topic_sel")
            if st.button("🗑️ Xoá mã đề này", type="secondary"):
                c.execute("DELETE FROM topics WHERE topic_code = %s", (del_t_code,))
                conn.commit()
                st.success(f"Đã xoá mã đề {del_t_code}!")
                st.rerun()
        else:
            st.info("Chưa có đề thi nào trong ngân hàng đề.")

    # TAB 4: QUẢN LÝ HỌC SINH RIÊNG CỦA GIÁO VIÊN
    with t_tab3:
        st.markdown(f"### 👥 Danh sách học sinh do Thầy/Cô **{user['fullname']}** trực tiếp quản lý")
        
        st.markdown("#### 📂 1. Cấp tài khoản hàng loạt cho lớp từ file Excel")
        st.caption("File Excel gồm 2 cột: Cột 1 là Mã học sinh, Cột 2 là Họ và tên. Tài khoản tạo ra sẽ tự động thuộc về lớp của Thầy/Cô.")
        
        uploaded_excel = st.file_uploader("Tải file Excel danh sách lớp (.xlsx, .xls):", type=["xlsx", "xls"])
        if uploaded_excel is not None:
            try:
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
                                    c.execute("INSERT INTO users (username, password, fullname, role, teacher_username) VALUES (%s, '123456', %s, 'student', %s)", (u_code, fullname, user["username"]))
                                    conn.commit()
                                    created_count += 1
                                except psycopg2.IntegrityError:
                                    conn.rollback()
                                    skipped_count += 1
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
                            c.execute("INSERT INTO users (username, password, fullname, role, teacher_username) VALUES (%s, %s, %s, 'student', %s)", (new_u.strip(), new_p.strip(), new_name.strip(), user["username"]))
                            conn.commit()
                            st.success(f"Đã thêm học sinh **{new_name}** vào lớp!")
                            st.rerun()
                        except Exception:
                            conn.rollback()
                            st.error("Tên đăng nhập này đã tồn tại!")
                    else:
                        st.warning("Vui lòng nhập đầy đủ thông tin.")
                        
        with col_remove:
            st.markdown("#### ❌ 3. Xoá học sinh khỏi lớp")
            c.execute("SELECT username, fullname FROM users WHERE role = 'student' AND teacher_username = %s", (user["username"],))
            students = c.fetchall()
            
            if students:
                student_dict = {f"{s[1]} ({s[0]})": s[0] for s in students}
                target_student = st.selectbox("Chọn học sinh cần xoá:", list(student_dict.keys()))
                student_user_to_delete = student_dict[target_student]
                
                confirm_del = st.checkbox("Xác nhận xoá toàn bộ dữ liệu của học sinh này")
                if st.button("🗑️ Xoá vĩnh viễn học sinh", type="primary", disabled=not confirm_del):
                    c.execute("DELETE FROM submissions WHERE username = %s", (student_user_to_delete,))
                    c.execute("DELETE FROM users WHERE username = %s", (student_user_to_delete,))
                    conn.commit()
                    st.success(f"Đã xoá học sinh {target_student} và toàn bộ bài làm!")
                    st.rerun()
            else:
                st.info("Chưa có học sinh nào trong lớp của Thầy/Cô.")

        st.markdown("---")
        st.markdown(f"#### 📋 Danh sách học sinh hiện tại của lớp ({user['fullname']}):")
        c.execute("SELECT username, fullname FROM users WHERE role = 'student' AND teacher_username = %s", (user["username"],))
        current_students = c.fetchall()
        if current_students:
            st.table([{"Mã đăng nhập": s[0], "Họ và tên học sinh": s[1]} for s in current_students])
        else:
            st.caption("Chưa có học sinh nào.")

    # TAB 5: CẤP THÊM TÀI KHOẢN GIÁO VIÊN MỚI (CHỈ SUPER ADMIN)
    if is_super_admin and len(tabs) > 5:
        with tabs[5]:
            st.markdown("### 👑 Khu vực Quản trị Trưởng: Cấp thêm tài khoản Giáo viên")
            st.info("💡 **Lưu ý:** Chỉ tài khoản Quản trị trưởng mới có quyền truy cập tab này.")
            
            with st.form("create_teacher_form"):
                new_t_user = st.text_input("Tên đăng nhập Giáo viên:", placeholder="Ví dụ: gv_lan")
                new_t_pass = st.text_input("Mật khẩu ban đầu:", value="gv123456")
                new_t_name = st.text_input("Họ và tên Giáo viên:", placeholder="Ví dụ: Nguyễn Thị Lan")
                btn_t = st.form_submit_button("Tạo tài khoản Giáo viên", type="primary")
                if btn_t:
                    if new_t_user and new_t_pass and new_t_name:
                        try:
                            c.execute("INSERT INTO users (username, password, fullname, role, teacher_username) VALUES (%s, %s, %s, 'teacher', NULL)", (new_t_user.strip(), new_t_pass.strip(), new_t_name.strip()))
                            conn.commit()
                            st.success(f"🎉 Đã tạo thành công tài khoản cho Giáo viên: **{new_t_name}** (Username: `{new_t_user}`)!")
                            st.rerun()
                        except Exception:
                            conn.rollback()
                            st.error("Tên đăng nhập này đã tồn tại, vui lòng chọn tên khác!")
                    else:
                        st.warning("Vui lòng nhập đầy đủ thông tin.")
                        
            st.markdown("---")
            st.markdown("#### 📋 Danh sách tất cả Giáo viên trong hệ thống:")
            c.execute("SELECT username, fullname FROM users WHERE role = 'teacher'")
            all_t = c.fetchall()
            st.table([{"Tên đăng nhập": t[0], "Họ và tên Giáo viên": t[1], "Quyền hạn": "Quản trị trưởng (Super Admin)" if t[0] in SUPER_ADMIN_USERS else "Giáo viên bộ môn"} for t in all_t])
        
    c.close()
    conn.close()
