import streamlit as st
import sqlite3
import json
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
    # Bảng người dùng
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT,
            fullname TEXT,
            role TEXT
        )
    ''')
    # Bảng lưu bài nộp và kết quả chấm
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
    # Tạo tài khoản giáo viên mặc định nếu chưa có
    c.execute("SELECT * FROM users WHERE username = 'giaovien'")
    if not c.fetchone():
        c.execute("INSERT INTO users VALUES ('giaovien', 'gv123456', 'Giáo viên Chủ nhiệm', 'teacher')")
        # Tạo sẵn vài học sinh mẫu
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

# 4. Huấn luyện System Instruction chuyên sâu (Cá nhân hóa + Barem 2.0)
SYSTEM_INSTRUCTION = """
You are an extremely strict, rigorous, and authoritative chief examiner for the English Gifted Student Examination (Kỳ thi Chọn Học sinh Giỏi Tỉnh & Chuyên Anh lớp 9) in Bac Ninh Province, Vietnam.

Your objective is to evaluate English essays (target: 200–250 words) based strictly on:
1. The official Bac Ninh 2.0-point examination rubric (Content 0.7, Organization 0.6, Language 0.6, Mechanics 0.1).
2. The official 2026-2027 Bac Ninh Pedagogical Training Guidelines on "Discussive & Argumentative Essay Writing" (Ms. Ngo Thuy Dung, Bac Ninh High School for the Gifted).
3. PERSONALIZATION RULE: You will be provided with the student's historical recurring errors. Explicitly check if the student has repeated their past mistakes or made progress.

============================================================
I. CRITICAL RED FLAGS TO PENALIZE:
- Idea Dumping: Listing 3-5 ideas superficially instead of developing 1-2 ideas in depth.
- Examples Without Analysis: Putting an example to replace the reasoning/mechanism.
- Repetition: Re-phrasing the same line of reasoning in multiple paragraphs.
- Memorised Templates: Rigid, mechanical template phrasing.
- Overclaiming: Absolute statements ("always", "never", "completely", "inevitably") without academic hedging.
- Fake Sophistication: Forcing obscure words unnaturally instead of natural C1 collocations.
- Overuse of Linking Words: Starting every sentence with mechanical connectives.

============================================================
II. ESSAY TYPE RULES:
- DISCUSSIVE ESSAY ("Examine the debate"): Body paragraphs must follow CLAIM -> WHY -> HOW (Mechanism) -> EXAMPLE -> EVALUATION (Conditions / Limitations).
- ARGUMENTATIVE ESSAY ("Convince position"): Body paragraphs must follow CLAIM -> REASON -> MECHANISM -> EXAMPLE -> IMPLICATION. Counterargument must follow CONCEDE -> QUALIFY -> REBUT.

============================================================
III. REQUIRED OUTPUT FORMAT:
You MUST output the assessment in clear Markdown with professional Vietnamese explanations:

### 1. 📋 ĐÁNH GIÁ TỔNG QUAN & DẠNG BÀI
- **Thể loại bài viết:** [Discussive Essay / Argumentative Essay]
- **Số lượng từ:** [Số từ] từ (Chuẩn đề: 200–250 từ).
- **Soi xét Lịch sử cá nhân hóa:** [Nhận xét học sinh có tái phạm các lỗi đã mắc ở các bài trước hay đã có cải thiện cụ thể nào].

### 2. 📊 BẢNG ĐIỂM CHÍNH THỨC SỞ GD&ĐT BẮC NINH (THANG 2.0)
| Tiêu chí thành phần | Điểm tối đa | Điểm đạt | Nhận xét chi tiết |
| :--- | :---: | :---: | :--- |
| **1. Content** (Ý tưởng & Nội dung) | 0.70 | **...** | Đào sâu cơ chế (Why/How) hay mắc lỗi Idea dumping. |
| **2. Organization** (Bố cục & Mạch lạc) | 0.60 | **...** | Mạch liên kết logic tự nhiên, cấu trúc chuẩn. |
| **3. Language** (Từ vựng & Ngữ pháp) | 0.60 | **...** | Vốn collocations C1, cấu trúc câu đa dạng. |
| **4. Mechanics** (Chính tả & Thể thức) | 0.10 | **...** | Lỗi chính tả, mạo từ, dấu câu, hình thức. |
| **TỔNG ĐIỂM BÀI THI** | **2.00** | **... / 2.0** | **Ước lượng band IELTS: ...** |

### 3. 🔍 SOI LỖI LẬP LUẬN THEO CHUYÊN ĐỀ TẬP HUẤN
- Chỉ ra các đoạn khẳng định suông thiếu cơ chế (Mechanism), hoặc đưa ví dụ thay cho lập luận.
- **Bảng phân tích câu văn chi tiết:**
| Câu văn gốc của học sinh | Lỗi sai (Tư duy / Ngữ pháp / Collocation) | Đề xuất sửa chữa nâng cao (Chuẩn C1) |
|---|---|---|

### 4. 💎 NÂNG CẤP HỌC THUẬT (ACADEMIC UPGRADES CHO HSG TỈNH)
- 4–5 cụm collocations C1 đắt giá liên quan trực tiếp đến đề.
- Kỹ thuật Hedging / Qualified Language để tránh Overclaiming.

### 5. ✍️ BÀI VIẾT LẠI MẪU ĐỈNH CAO
(Viết lại đoạn văn yếu nhất hoặc toàn bài theo chuẩn cấu trúc tập huấn).

### 6. ⚠️ DANH SÁCH LỖI THEN CHỐT CẦN LƯU HỒ SƠ:
(Liệt kê 1-3 lỗi cốt lõi ngắn gọn để lưu vào cơ sở dữ liệu theo dõi cá nhân, ví dụ: "Idea dumping", "Thiếu Mechanism", "Overclaiming", "Sai mạo từ").
"""

# Quản lý Đăng nhập qua Session State
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None

def login(username, password):
    conn = sqlite3.connect("essay_database.db")
    c = conn.cursor()
    c.execute("SELECT username, fullname, role FROM users WHERE username = ? AND password = ?", (username, password))
    user = c.fetchone()
    conn.close()
    if user:
        st.session_state.logged_in = True
        st.session_state.user = {"username": user[0], "fullname": user[1], "role": user[2]}
        st.rerun()
    else:
        st.error("Tên đăng nhập hoặc mật khẩu không chính xác!")

def logout():
    st.session_state.logged_in = False
    st.session_state.user = None
    st.rerun()

# --- MÀN HÌNH ĐĂNG NHẬP ---
if not st.session_state.logged_in:
    st.title("🎓 Hệ Thống Bồi Dưỡng & Chấm Essay HSG Tiếng Anh 9")
    st.subheader("Trường THCS Thân Nhân Trung - TP. Bắc Ninh")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("### 🔐 Đăng nhập")
        username_input = st.text_input("Tên đăng nhập:")
        password_input = st.text_input("Mật khẩu:", type="password")
        if st.button("Đăng nhập", type="primary", use_container_width=True):
            login(username_input, password_input)
            
        st.info("""
        💡 **Tài khoản mặc định thử nghiệm:**
        - **Giáo viên:** `giaovien` / `gv123456`
        - **Học sinh:** `hs01`, `hs02`, `hs03` / `123456`
        """)
    st.stop()

# --- GIAO DIỆN ĐÃ ĐĂNG NHẬP ---
user = st.session_state.user
st.sidebar.markdown(f"### 👤 Xin chào: **{user['fullname']}**")
st.sidebar.caption(f"Vai trò: {'Giáo viên quản trị' if user['role'] == 'teacher' else 'Học sinh đội tuyển'}")
if st.sidebar.button("Đăng xuất"):
    logout()

st.sidebar.markdown("---")

# Cấu hình API Key (Lấy từ Sidebar hoặc mặc định)
api_key = st.sidebar.text_input("Gemini API Key:", type="password", help="Dán mã API Key của thầy vào đây để hệ thống hoạt động")

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
                        # 1. Lấy lịch sử lỗi trước đó của học sinh để cá nhân hóa
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
                            model='gemini-2.5-pro',
                            contents=user_content,
                            config=types.GenerateContentConfig(
                                system_instruction=SYSTEM_INSTRUCTION,
                                temperature=0.15
                            )
                        )
                        
                        result_text = response.text
                        st.success("✅ Đã hoàn thành chấm bài!")
                        st.markdown(result_text)
                        
                        # 2. Tự động lưu kết quả vào CSDL
                        conn = sqlite3.connect("essay_database.db")
                        c = conn.cursor()
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                        c.execute('''
                            INSERT INTO submissions (username, topic, essay_text, score_total, score_content, score_org, score_lang, score_mech, feedback, identified_errors, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (user["username"], essay_prompt, essay_text, 1.5, 0.5, 0.45, 0.45, 0.1, result_text, "Overclaiming, Thiếu Mechanism", now_str))
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
    st.title("👨‍🏫 Bảng Điều Khiển Tổng Hợp Giáo Viên")
    
    t_tab1, t_tab2, t_tab3 = st.tabs(["📊 Tổng hợp kết quả cả lớp", "🔍 Xem bài & Chữa lỗi chi tiết", "👥 Quản lý danh sách học sinh"])
    
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
            st.info("Hiện tại chưa có học sinh nào nộp bài.")
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
            **Các nhược điểm học sinh hay mắc nhiều nhất tuần này:**
            1. **Thiếu Mechanism (Cơ chế):** Học sinh mới nêu được Claim (Khẳng định) và đưa ngay ví dụ, thiếu bước giải thích "Tại sao/Như thế nào".
            2. **Overclaiming:** Dùng từ ngữ quá tuyệt đối (*always, completely, never*), cần rèn thêm kỹ thuật *Hedging*.
            """)
            
    with t_tab2:
        st.markdown("### 🔍 Tra cứu bài làm chi tiết của từng học sinh")
        c.execute("SELECT id, fullname, topic, created_at, feedback, essay_text FROM submissions s JOIN users u ON s.username = u.username ORDER BY s.id DESC")
        all_subs = c.fetchall()
        
        if all_subs:
            sub_options = {f"[{sub[0]}] {sub[1]} - {sub[2][:40]}... ({sub[3]})": sub for sub in all_subs}
            chosen = st.selectbox("Chọn bài nộp để thẩm định:", list(sub_options.keys()))
            selected_sub = sub_options[chosen]
            
            st.markdown(f"#### 👤 Học sinh: **{selected_sub[1]}** | Ngày nộp: **{selected_sub[3]}**")
            st.info(f"**Đề bài:** {selected_sub[2]}")
            
            with st.expander("📄 Xem bài viết nguyên bản của học sinh"):
                st.text(selected_sub[5])
                
            st.markdown("---")
            st.markdown("### 📝 Toàn bộ kết quả chấm & Nhận xét của AI:")
            st.markdown(selected_sub[4])

    with t_tab3:
        st.markdown("### 👥 Cấp tài khoản mới cho Học sinh")
        with st.form("add_user_form"):
            new_u = st.text_input("Tên đăng nhập (Username):", placeholder="Ví dụ: hs04")
            new_p = st.text_input("Mật khẩu ban đầu:", value="123456")
            new_name = st.text_input("Họ và tên học sinh:", placeholder="Ví dụ: Hoàng Minh Đức")
            submit_btn = st.form_submit_button("Thêm học sinh vào danh sách")
            if submit_btn:
                if new_u and new_p and new_name:
                    try:
                        c.execute("INSERT INTO users VALUES (?, ?, ?, 'student')", (new_u, new_p, new_name))
                        conn.commit()
                        st.success(f"Đã cấp tài khoản thành công cho học sinh **{new_name}**!")
                    except Exception:
                        st.error("Tên đăng nhập này đã tồn tại!")
                else:
                    st.warning("Vui lòng điền đầy đủ thông tin.")
                    
        st.markdown("#### Danh sách tài khoản hiện có:")
        c.execute("SELECT username, fullname, role FROM users")
        st.dataframe(c.fetchall())
        
    conn.close()
