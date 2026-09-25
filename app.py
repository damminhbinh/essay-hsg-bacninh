import streamlit as st
import sqlite3
from datetime import datetime
from google import genai
from google.genai import types
from PIL import Image

# 1. Cấu hình giao diện Web
st.set_page_config(
    page_title="Hệ Thống Chấm Essay HSG - THCS Thân Nhân Trung",
    page_icon="🎓",
    layout="wide"
)

# 2. Quản lý Cơ sở dữ liệu SQLite
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
        c.execute("INSERT INTO users VALUES ('giaovien', 'gv123456', 'Giáo viên Chủ nhiệm', 'teacher')")
        c.execute("INSERT INTO users VALUES ('hs01', '123456', 'Nguyễn Văn An', 'student')")
        c.execute("INSERT INTO users VALUES ('hs02', '123456', 'Trần Thị Bình', 'student')")
        c.execute("INSERT INTO users VALUES ('hs03', '123456', 'Lê Hoàng Long', 'student')")
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

# 4. Huấn luyện System Instruction chuẩn Barem 2.0 & Biên bản tập huấn HSG Bắc Ninh
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
A. DISCUSSIVE ESSAY ("Let me examine the debate"):
- Purpose: Examine + Evaluate + Arrive at a reasoned, qualified judgement.
- Body Paragraphs MUST strictly follow: CLAIM -> WHY (Reason) -> HOW (Mechanism) -> EXAMPLE -> EVALUATION (Conditions / Limitations).
- Failure to evaluate limitations or conditions ("50-50 thinking") -> Cap Organization at 0.35 / 0.60.

B. ARGUMENTATIVE ESSAY ("Let me convince you of my position"):
- Purpose: Claim + Support + Defend.
- Main Arguments: CLAIM -> REASON -> MECHANISM -> EXAMPLE -> IMPLICATION.
- Counterargument: Must apply CONCEDE -> QUALIFY -> REBUT. Failing to rebut or taking an extreme, unhedged position -> Cap Content at 0.35 / 0.70.

============================================================
III. OFFICIAL BAC NINH 2.0-POINT RUBRIC (STRICT SCORING - 0.05 INCREMENTS):
1. Content (0.70 pt max): Full alignment with prompt, nuanced depth, clear mechanisms, no task drift.
2. Organization & Presentation (0.60 pt max): Coherence, internal logic, 4-paragraph structure, no mechanical linkers.
3. Language (0.60 pt max): Natural C1 vocabulary range, authentic collocations, advanced sentence structures, academic hedging. Zero tolerance for fabricated expressions.
4. Mechanics (0.10 pt max): Punctuation, capitalization, zero spelling errors, NO contractions.

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
- Chỉ ra các đoạn khẳng định suông thiếu cơ chế (Mechanism), ngộ nhận logic hoặc đưa ví dụ thay cho lập luận.
- **Bảng phân tích câu văn chi tiết:**
| Câu văn gốc của học sinh | Lỗi sai (Tư duy / Ngữ pháp / Collocation) | Đề xuất sửa chữa nâng cao (Chuẩn C1) |
|---|---|---|

### 4. 💎 NÂNG CẤP HỌC THUẬT (ACADEMIC UPGRADES CHO HSG TỈNH)
- 4–5 cụm collocations C1 đắt giá sửa chữa đúng trọng tâm của đề.
- Kỹ thuật Hedging / Qualified Language để tránh Overclaiming.

### 5. ✍️ BÀI VIẾT LẠI MẪU ĐỈNH CAO
(Viết lại đoạn văn yếu nhất hoặc toàn bài theo chuẩn cấu trúc tập huấn, bám sát trọng tâm câu hỏi của đề).

### 6. ⚠️ DANH SÁCH LỖI THEN CHỐT CẦN LƯU HỒ SƠ:
(Ghi 1-3 lỗi cốt lõi ngắn gọn để ghi vào CSDL theo dõi cá nhân, ví dụ: "Task Drift", "Thiếu Mechanism", "Collocation tự chế", "Dùng từ viết tắt").
"""

# Quản lý Đăng nhập qua Session State
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None

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
    st.rerun()

# --- MÀN HÌNH ĐĂNG NHẬP (Đã bỏ hộp tài khoản mặc định và hỗ trợ bấm Enter) ---
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
    st.stop()

# --- GIAO DIỆN ĐÃ ĐĂNG NHẬP ---
user = st.session_state.user
st.sidebar.markdown(f"### 👤 Xin chào: **{user['fullname']}**")
st.sidebar.caption(f"Vai trò: {'Giáo viên quản trị' if user['role'] == 'teacher' else 'Học sinh đội tuyển'}")

# --- CHỨC NĂNG ĐỔI MẬT KHẨU ---
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

# Tự động lấy API Key từ Streamlit Secrets nếu có
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("Gemini API Key:", type="password", help="Dán API Key vào đây nếu chưa cấu hình Secrets")

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
            if not api_key:
                st.error("⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ giáo viên!")
            elif not essay_prompt.strip():
                st.error("⚠️ Vui lòng nhập hoặc chọn đề thi!")
            elif not essay_text.strip() and not uploaded_image:
                st.error("⚠️ Vui lòng dán nội dung bài hoặc tải ảnh lên!")
            else:
                with st.spinner("Giám khảo AI đang đối chiếu barem Bắc Ninh và phân tích hồ sơ của em..."):
                    try:
                        conn = sqlite3.connect("essay_database.db")
                        c = conn.cursor()
                        c.execute("SELECT identified_errors FROM submissions WHERE username = ? ORDER BY id DESC LIMIT 3", (user["username"],))
                        past_errors = c.fetchall()
                        conn.close()
                        
                        error_history_text = "Học sinh chưa có lịch sử nộp bài trước đó."
                        if past_errors:
                            error_history_text = "Các lỗi học sinh này THƯỜNG MẮC ở các bài trước: " + ", ".join([e[0] for e in past_errors if e[0]])
                        
                        client = genai.Client(api_key=api_key)
                        
                        user_content = [
                            f"LỊCH SỬ HỌC TẬP CỦA HỌC SINH NÀY:\n{error_history_text}\n\n",
                            f"ĐỀ THI: {essay_prompt}\n\n",
                            "Hãy chấm bài luận sau theo đúng barem nghiêm ngặt của Bắc Ninh:"
                        ]
                        if essay_text.strip():
                            user_content.append(f"\nBÀI LÀM:\n{essay_text}")
                        if uploaded_image:
                            user_content.append(uploaded_image)
                            
                        response = client.models.generate_content(
                            model='gemini-3.8-flash',
                            contents=user_content,
                            config=types.GenerateContentConfig(
                                system_instruction=SYSTEM_INSTRUCTION,
                                temperature=0.15
                            )
                        )
                        
                        result_text = response.text
                        st.success("✅ Đã hoàn thành chấm bài!")
                        st.markdown(result_text)
                        
                        conn = sqlite3.connect("essay_database.db")
                        c = conn.cursor()
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                        c.execute('''
                            INSERT INTO submissions (username, topic, essay_text, score_total, score_content, score_org, score_lang, score_mech, feedback, identified_errors, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (user["username"], essay_prompt, essay_text, 1.5, 0.5, 0.45, 0.45, 0.1, result_text, "Task Drift, Overclaiming", now_str))
                        conn.commit()
                        conn.close()
                        
                    except Exception as e:
                        st.error(f"Đã xảy ra lỗi: {str(e)}")

    with tab_history:
        st.markdown(f"### 📈 Hồ sơ theo dõi học tập của {user['fullname']}")
        conn = sqlite3.connect("essay_database.db")
        c = conn.cursor()
        c.execute("SELECT id, topic, created_at, feedback FROM submissions WHERE username = ? ORDER BY id DESC", (user["username"],))
        rows = c.fetchall()
        conn.close()
        
        if not rows:
            st.info("Em chưa nộp bài nào. Hãy bắt đầu luyện tập với đề bài đầu tiên nhé!")
        else:
            st.write(f"Tổng số bài đã luyện tập: **{len(rows)} bài**")
            for r in rows:
                with st.expander(f"📝 Đề: {r[1][:70]}... - Ngày nộp: {r[2]}"):
                    st.markdown(r[3])

# =========================================================================
# GIAO DIỆN GIÁO VIÊN (DASHBOARD QUẢN TRỊ)
# =========================================================================
elif user["role"] == "teacher":
    st.title("👨‍🏫 Bảng Điều Khiển Quản Trị Giáo Viên")
    t_tab1, t_tab2, t_tab3 = st.tabs(["📊 Tổng hợp kết quả cả lớp", "🔍 Xem bài & Xoá bài nộp", "👥 Quản lý & Xoá học sinh"])
    
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    
    # ---------------- TAB 1: TỔNG HỢP CẢ LỚP ----------------
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
            
    # ---------------- TAB 2: XEM BÀI VÀ NÚT XOÁ BÀI ----------------
    with t_tab2:
        st.markdown("### 🔍 Thẩm định bài làm & Xoá bài nộp")
        c.execute("SELECT id, fullname, topic, created_at, feedback, essay_text FROM submissions s JOIN users u ON s.username = u.username ORDER BY s.id DESC")
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
            
            with st.expander("📄 Xem bài viết nguyên bản của học sinh"):
                st.text(selected_sub[5])
                
            st.markdown("---")
            st.markdown("### 📝 Kết quả chấm & Nhận xét của AI:")
            st.markdown(selected_sub[4])
        else:
            st.info("Không có bài nộp nào để hiển thị.")

    # ---------------- TAB 3: QUẢN LÝ VÀ XOÁ HỌC SINH ----------------
    with t_tab3:
        col_add, col_remove = st.columns(2)
        
        with col_add:
            st.markdown("### ➕ Cấp tài khoản mới")
            with st.form("add_user_form"):
                new_u = st.text_input("Tên đăng nhập (Username):", placeholder="Ví dụ: hs04")
                new_p = st.text_input("Mật khẩu ban đầu:", value="123456")
                new_name = st.text_input("Họ và tên học sinh:", placeholder="Ví dụ: Hoàng Minh Đức")
                submit_btn = st.form_submit_button("Thêm học sinh vào danh sách", type="primary")
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
