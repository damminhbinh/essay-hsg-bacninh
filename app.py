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
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
- **1. Đàm Thuận Minh Bình**  
  📞 0387.136.888
- **2. Đỗ Thị Huyền**  
  📞 0982.036.952

🏫 *Trường THCS Thân Nhân Trung - TP. Bắc Ninh*
"""

# HÀM BÓC TÁCH ĐIỂM THỰC TẾ THEO 8 TIÊU CHÍ CHI TIẾT & LỖI TỪ PHẢN HỒI CỦA GIÁM KHẢO AI
def parse_scores_from_feedback(text):
    s_content = 0.0
    s_org = 0.0
    s_lang = 0.0
    s_mech = 0.0
    s_total = 0.0
    errors_str = "Chưa ghi nhận lỗi nghiêm trọng"
    
    if not text:
        return s_content, s_org, s_lang, s_mech, s_total, errors_str
        
    try:
        # Làm sạch các ký tự markdown thừa (*, _, `, [, ]) để regex bắt chính xác 100%
        clean_text = re.sub(r'[\*_`\[\]]', '', text)

        def extract_val(pattern):
            m = re.search(pattern, clean_text, re.IGNORECASE)
            return float(m.group(1)) if m else 0.0

        # 1. Content (0.70 = 1a + 1b)
        s_1a = extract_val(r'1a\.\s*Main Ideas\s*(?:&|and)\s*Details\s*:\s*([0-9.]+)')
        s_1b = extract_val(r'1b\.\s*Communicating Intentions\s*:\s*([0-9.]+)')
        s_content = round(s_1a + s_1b, 3)

        # 2. Organization (0.60 = 2a + 2b)
        s_2a = extract_val(r'2a\.\s*Cohesion\s*(?:&|and)\s*Coherence\s*:\s*([0-9.]+)')
        s_2b = extract_val(r'2b\.\s*Structural Composition\s*:\s*([0-9.]+)')
        s_org = round(s_2a + s_2b, 3)

        # 3. Language (0.60 = 3a + 3b)
        s_3a = extract_val(r'3a\.\s*Vocabulary Range\s*:\s*([0-9.]+)')
        s_3b = extract_val(r'3b\.\s*Grammatical Control\s*:\s*([0-9.]+)')
        s_lang = round(s_3a + s_3b, 3)

        # 4. Punctuation, Spelling & Word Count (0.10 = 4a + 4b)
        s_4a = extract_val(r'4a\.\s*Spelling\s*(?:&|and)\s*Punctuation\s*:\s*([0-9.]+)')
        s_4b = extract_val(r'4b\.\s*Word Count\s*(?:&|and)\s*Presentation\s*:\s*([0-9.]+)')
        s_mech = round(s_4a + s_4b, 3)

        # Tổng điểm (ưu tiên lấy từ TOTAL SCORE, nếu không có thì cộng 4 nhóm)
        m_t = re.search(r'TOTAL SCORE\s*:\s*([0-9.]+)\s*/\s*2(?:\.0+)?', clean_text, re.IGNORECASE)
        if m_t:
            s_total = round(float(m_t.group(1)), 2)
        else:
            s_total = round(s_content + s_org + s_lang + s_mech, 2)

        # Bóc tách lỗi then chốt từ bảng Actionable Corrections
        m_err = re.search(r'4\.\s*\*{0,2}Actionable Corrections\*{0,2}[:\s*\n]+(.*?)(?:\n###|\n5\.|\Z)', text, re.DOTALL | re.IGNORECASE)
        if m_err:
            raw_err = m_err.group(1).strip()
            cleaned_lines = [
                re.sub(r'^[\s*\-0-9.)|]+', '', line).strip()
                for line in raw_err.split('\n')
                if line.strip() and not line.strip().startswith('|---') and not line.strip().lower().startswith('| lỗi gốc')
            ]
            if cleaned_lines:
                errors_str = " | ".join(cleaned_lines[:3])
                if len(errors_str) > 120:
                    errors_str = errors_str[:117] + "..."
    except Exception:
        pass
        
    return s_content, s_org, s_lang, s_mech, s_total, errors_str

# CÁC HÀM HỖ TRỢ ĐỊNH DẠNG BẢNG VÀ CHỮ CHO FILE WORD
def set_table_styling(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="B0C4DE"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="B0C4DE"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="B0C4DE"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="B0C4DE"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def set_cell_properties(cell, bg_hex=None, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    if bg_hex:
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>')
        tcPr.append(shd)
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def add_formatted_runs_to_paragraph(p, text, font_size=11, default_bold=False):
    lines = re.split(r'<br\s*/?>|\n', text)
    first_line = True
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
        if not first_line:
            p.add_run('\n')
        first_line = False
        
        tokens = re.split(r'(\*\*.*?\*\*|\*.*?\*)', cleaned_line)
        for token in tokens:
            if not token:
                continue
            if token.startswith('**') and token.endswith('**') and len(token) >= 4:
                r = p.add_run(token[2:-2])
                r.bold = True
            elif token.startswith('*') and token.endswith('*') and len(token) >= 2:
                r = p.add_run(token[1:-1])
                r.italic = True
            else:
                r = p.add_run(token)
                if default_bold:
                    r.bold = True
            r.font.name = 'Times New Roman'
            r.font.size = Pt(font_size)

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
                    set_table_styling(t)
                    
                    for r_idx, row in enumerate(parsed_rows):
                        is_header = (r_idx == 0)
                        row_bg = "0B3C5D" if is_header else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
                        
                        for c_idx, val in enumerate(row):
                            if c_idx < cols_count:
                                cell = t.cell(r_idx, c_idx)
                                set_cell_properties(cell, bg_hex=row_bg, top=100, bottom=100, left=140, right=140)
                                p = cell.paragraphs[0]
                                p.text = ""
                                p.paragraph_format.line_spacing = 1.15
                                p.paragraph_format.space_before = Pt(0)
                                p.paragraph_format.space_after = Pt(0)
                                
                                if is_header:
                                    r = p.add_run(val)
                                    r.bold = True
                                    r.font.name = 'Times New Roman'
                                    r.font.size = Pt(11)
                                    r.font.color.rgb = RGBColor(255, 255, 255)
                                else:
                                    add_formatted_runs_to_paragraph(p, val, font_size=11)
                                    
                in_table = False
                table_lines = []
                
            if not stripped:
                continue
                
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.25
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(3)
            
            clean_line = stripped
            is_heading_bold = False
            heading_size = 13
            
            if clean_line.startswith('### '):
                clean_line = clean_line[4:]
                is_heading_bold = True
                heading_size = 13.5
            elif clean_line.startswith('#### '):
                clean_line = clean_line[5:]
                is_heading_bold = True
                heading_size = 12.5
            elif clean_line.startswith('#'):
                clean_line = clean_line.lstrip('#').strip()
                is_heading_bold = True
                
            add_formatted_runs_to_paragraph(p, clean_line, font_size=heading_size, default_bold=is_heading_bold)

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

# 3. Huấn luyện System Instruction chuẩn Barem 2.0 Bắc Ninh & Template Argumentative Essay
SYSTEM_INSTRUCTION = """
============================================================
0. QUY TẮC BẮT BUỘC VỀ SỰ PHÙ HỢP CỦA CHỦ ĐỀ (OFF-TOPIC ZERO RULE):
- Trước khi chấm, kiểm tra mức độ liên quan giữa bài viết của học sinh và yêu cầu của đề bài.
- Lưu ý khi chấm ảnh chụp bài viết tay: Phải cố gắng đọc toàn bộ nội dung chữ viết tay của thí sinh; chỉ áp dụng điểm 0 khi nội dung thực sự hoàn toàn không liên quan đến đề bài.
- Nếu bài viết lạc đề hoàn toàn, chép lại đề/văn mẫu không ăn nhập, hoặc trả lời sai chủ đề trọng tâm:
  + Đặt TẤT CẢ 8 tiêu chí thành phần và TOTAL SCORE = 0.0 / 2.0.
  + Nêu rõ lý do: "Bài viết lạc đề / không trả lời đúng yêu cầu của đề bài (Off-topic). Điểm đánh giá: 0."

============================================================
# BAREM CHẤM ĐIỂM CHI TIẾT (RUBRIC)

### 1. Content (0.7 điểm) - 35%
- **a. Main ideas and details (0.35 điểm):**
  - Đưa ra quan điểm rõ ràng (Agree / Disagree / Partial agreement) ngay từ mở bài.
  - Có ít nhất 2 luận điểm chính được phát triển mạch lạc với ví dụ, giải thích hoặc dẫn chứng cụ thể.
  - Phạt nếu lạc đề, luận điểm mơ hồ hoặc chỉ liệt kê mà không đào sâu:
    - Đầy đủ, sâu sắc: 0.30 - 0.35 điểm.
    - Đủ ý nhưng phát triển còn mỏng: 0.20 - 0.25 điểm.
    - Thiếu ý hoặc mất cân đối nghiêm trọng: 0.10 - 0.15 điểm.
    - Lạc đề: 0.0 điểm.
- **b. Communicating intentions (0.35 điểm):**
  - Thể hiện lập trường xuyên suốt từ đầu đến kết luận mà không tự mâu thuẫn.
  - Lập luận có tính thuyết phục cao, phân tích được cả tính 2 mặt nếu chọn thế trung lập:
    - Thuyết phục, nhất quán cao: 0.30 - 0.35 điểm.
    - Ý định rõ nhưng đôi chỗ lập luận lỏng lẻo: 0.20 - 0.25 điểm.
    - Không rõ lập trường hoặc tự mâu thuẫn: 0.05 - 0.15 điểm.

### 2. Organization & Presentation (0.6 điểm) - 30%
- **a. Coherence, cohesion & clarity (0.30 điểm):**
  - Mạch lập luận tự nhiên, chuyển đoạn và chuyển ý mượt mà.
  - Sử dụng từ nối/liên kết câu (cohesive devices) linh hoạt, tránh lạm dụng máy móc (Firstly, Secondly, Last but not least).
    - Tự nhiên, chuyển ý tinh tế: 0.25 - 0.30 điểm.
    - Đạt mức cơ bản, đôi chỗ lạm dụng hoặc gượng gạo: 0.15 - 0.20 điểm.
    - Rời rạc, thiếu liên kết: 0.0 - 0.10 điểm.
- **b. Well-structured composition (0.30 điểm):**
  - Cấu trúc chuẩn 4 hoặc 5 đoạn: Introduction (Hook + Thesis statement), 2-3 Body paragraphs (Topic sentence + Supporting sentences), Conclusion (Restatement + Final thought).
    - Cấu trúc chuẩn mực, phân đoạn chuẩn: 0.25 - 0.30 điểm.
    - Có đủ mở - thân - kết nhưng chia đoạn chưa tối ưu: 0.15 - 0.20 điểm.
    - Không phân đoạn hoặc chỉ có 1 đoạn thân bài cồng kềnh: 0.0 - 0.10 điểm.

### 3. Language (0.6 điểm) - 30% (Tiêu chuẩn HSG Chuyên Anh THCS - B2 to C1)
- **a. Vocabulary & Complex structures (0.30 điểm):**
  - Đòi hỏi từ vựng nâng cao đúng ngữ cảnh (academic collocations, topic-specific vocabulary: cognitive offloading, critical thinking, personalized learning, plagiarism...).
  - Đa dạng cấu trúc: câu phức, mệnh đề quan hệ rút gọn, đảo ngữ, câu điều kiện, bị động nâng cao.
    - Vốn từ phong phú, cấu trúc đa dạng và tự nhiên: 0.25 - 0.30 điểm.
    - Dùng từ đúng nhưng còn đơn giản (A2-B1), ít biến hóa cấu trúc: 0.15 - 0.20 điểm.
    - Cố gượng ép từ vựng (chêm từ đao to búa lớn sai ngữ cảnh): 0.10 - 0.15 điểm.
- **b. Grammatical control (0.30 điểm):**
  - Kiểm soát tốt sự hòa hợp chủ-vị, thì của động từ, mạo từ, giới từ.
    - Không hoặc chỉ có 1-2 lỗi nhỏ không ảnh hưởng nghĩa (slips): 0.25 - 0.30 điểm.
    - Xuất hiện một số lỗi ngữ pháp nhưng người đọc vẫn hiểu: 0.15 - 0.20 điểm.
    - Lỗi ngữ pháp dày đặc gây cản trở đọc hiểu: 0.0 - 0.10 điểm.

### 4. Punctuation, Spelling and Word Count (0.1 điểm) - 5%
- **a. Punctuation & Spelling (0.05 điểm):**
  - Đúng chính tả và chấm câu (dấu phẩy sau trạng ngữ, dấu chấm, viết hoa đầu câu):
    - Không mắc lỗi: 0.05 điểm.
    - Có 2-3 lỗi: 0.025 điểm.
    - Trên 3 lỗi: 0.0 điểm.
- **b. Formatting / Handwriting (0.05 điểm):**
  - Quy ước khi chấm văn bản số/text: Tiêu chí này được quy đổi sang "Trình bày chuẩn mực & Độ dài hợp lệ":
    - Độ dài đạt từ 220 - 280 từ (dung sai ±10% cho mức 250 từ) và trình bày văn bản chuẩn: 0.05 điểm.
    - Quá ngắn (<200 từ) hoặc quá dài (>320 từ): 0.025 điểm.
    - Dưới 150 từ: 0.0 điểm.

============================================================
# QUY TẮC PHẠT VÀ TÍNH ĐIỂM
1. **Làm tròn điểm:** Tổng điểm làm tròn đến 2 chữ số thập phân (bước nhảy 0.05 hoặc 0.025).
2. **Khách quan:** Chỉ chấm những gì có trên bài viết, không tự suy diễn ý định mà học sinh chưa diễn đạt thành lời.
3. **Phân tích lỗi cụ thể:** Trích dẫn chính xác lỗi sai (Word/Sentence quote) kèm theo đề xuất sửa đổi.

============================================================
# QUY CHUẨN GỢI Ý VIẾT LẠI CHO DẠNG "TO WHAT EXTENT DO YOU AGREE OR DISAGREE?" (ARGUMENTATIVE ESSAY)
Khi đề bài thuộc dạng "To what extent do you agree or disagree?", phần gợi ý viết lại BẮT BUỘC phải tuân thủ cấu trúc 4 đoạn và sử dụng linh hoạt bộ khung Template chuẩn sau:
- **Paragraph 1 - Introduction:**
  + Template Standard: "The issue of whether [statement] remains highly controversial. While there are legitimate arguments on both sides, I would argue that [clear position], primarily because [reason 1] and [reason 2]."
  + Hoặc Template Qualified Thesis: "Although [concession], I contend that [position]. This is because [reason 1], while [reason 2] further strengthens the case for this position."
- **Paragraph 2 - Body 1 (Main Argument):**
  + Template: "One of the strongest arguments in favour of [position] is that [claim]. This is because [reason]. In practical terms, [mechanism/development]. For instance, [example]. The implication is that [consequence]."
- **Paragraph 3 - Body 2 (Counterargument & Rebuttal theo mô hình Concede -> Qualify -> Rebut):**
  + Template Standard: "Admittedly, opponents may argue that [counterargument]. This concern is understandable because [reason]. However, it overlooks the fact that [rebuttal]."
  + Hoặc Template Advanced: "It could reasonably be argued that [counterargument]. While this may hold true in [specific circumstances], it becomes less convincing when [different circumstance], particularly because [rebuttal]."
- **Paragraph 4 - Conclusion:**
  + Template: "In conclusion, although [counterargument], the arguments in favour of [position] are more compelling because [reason 1] and [reason 2]. Therefore, [final judgement]."

============================================================
# ĐỊNH DẠNG ĐẦU RA (OUTPUT MARKDOWN)

Trả về kết quả theo đúng cấu trúc sau:
1. **Word Count Check:** [Số từ thực tế]
2. **Score Breakdown:**
   - 1a. Main Ideas & Details: [X]/0.35
   - 1b. Communicating Intentions: [X]/0.35
   - 2a. Cohesion & Coherence: [X]/0.30
   - 2b. Structural Composition: [X]/0.30
   - 3a. Vocabulary Range: [X]/0.30
   - 3b. Grammatical Control: [X]/0.30
   - 4a. Spelling & Punctuation: [X]/0.05
   - 4b. Word Count & Presentation: [X]/0.05
   - **TOTAL SCORE:** [X]/2.0
3. **Detailed Justification:** Nhận xét ưu điểm và nhược điểm cụ thể cho từng tiêu chí.
4. **Actionable Corrections:**

| Lỗi gốc | Loại lỗi | Gợi ý sửa đổi | Giải thích ngắn |
|---|---|---|---|
| ... | ... | ... | ... |

5. **Suggested Rewrite (Gợi ý viết lại chuẩn ~250 từ):**
- Viết lại bài luận hoàn chỉnh cho đề bài trên. Nếu là dạng "To what extent do you agree or disagree?", áp dụng triệt để cấu trúc 4 đoạn và các Template Argumentative Essay ở trên (giữ văn phong học thuật B2-C1, tránh overclaiming).
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
                                        system_instruction=SYSTEM_INSTRUCTION                                        
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
        col_title, col_sync = st.columns([3, 1])
        with col_title:
            st.markdown(f"### 📊 Báo cáo kết quả & Xuất Bảng điểm Excel của lớp")
        with col_sync:
            st.write("")
            if st.button("🔄 Đồng bộ lại điểm từ nhận xét cũ", use_container_width=True):
                c.execute("SELECT id, feedback FROM submissions")
                all_old_rows = c.fetchall()
                updated_cnt = 0
                for r_id, fb_text in all_old_rows:
                    if fb_text:
                        sc, so, sl, sm, stot, serr = parse_scores_from_feedback(fb_text)
                        if stot > 0:
                            c.execute('''
                                UPDATE submissions
                                SET score_content = %s, score_org = %s, score_lang = %s, score_mech = %s, score_total = %s, identified_errors = %s
                                WHERE id = %s
                            ''', (sc, so, sl, sm, stot, serr, r_id))
                            updated_cnt += 1
                conn.commit()
                st.success(f"Đã cập nhật đồng bộ lại điểm cho {updated_cnt} bài!")
                st.rerun()
        
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
                                            system_instruction=SYSTEM_INSTRUCTION                                            
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
        c.execute("SELECT username, fullname, password FROM users WHERE role = 'student' AND teacher_username = %s ORDER BY username ASC", (user["username"],))
        current_students = c.fetchall()
        total_hs = len(current_students)
        
        st.markdown(f"#### 📋 Danh sách học sinh hiện tại của lớp ({user['fullname']}) — Tổng số: **{total_hs} học sinh**")
        
        if current_students:
            table_data = [
                {
                    "STT": idx + 1, 
                    "Mã đăng nhập": s[0], 
                    "Họ và tên học sinh": s[1], 
                    "Mật khẩu hiện tại": s[2]
                } 
                for idx, s in enumerate(current_students)
            ]
            st.dataframe(table_data, use_container_width=True, hide_index=True)
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
