import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import requests
import json
import io
import re
import os

try:
    import openpyxl
except ImportError:
    openpyxl = None

CONFIG_FILE = "config_keys.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

config_data = load_config()

st.set_page_config(
    page_title="Phần mềm Cụ thể hóa Văn bản Hành chính",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    div[data-testid="stVerticalBlock"] > div {
        gap: 0.4rem !important;
    }
    
    .word-page {
        background-color: #ffffff;
        color: #000000;
        width: 100%;
        max-height: 460px;
        overflow-y: auto;
        padding: 25px 35px;
        border: 1px solid #d3d3d3;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15);
        font-family: 'Times New Roman', Times, serif;
        font-size: 13pt;
        line-height: 1.0;
        word-wrap: break-word;
        border-radius: 4px;
    }
    .word-page table {
        width: 100% !important;
        border-collapse: collapse !important;
        border: none !important;
        margin-bottom: 8px;
    }
    .word-page td {
        border: none !important;
        vertical-align: top;
        padding: 0px 4px;
        font-family: 'Times New Roman', Times, serif;
    }
    .word-page p {
        font-family: 'Times New Roman', Times, serif;
        font-size: 13pt;
        line-height: 1.0;
        margin-top: 4pt;
        margin-bottom: 4pt;
        text-align: justify;
        text-indent: 0.98cm;
    }
    .word-page p.kinh-gui {
        text-indent: 0 !important;
        margin-left: 0cm !important;
        text-align: center !important;
        margin-top: 4pt;
        margin-bottom: 4pt;
    }
    .word-page p.noi-nhan {
        text-indent: 0cm !important;
        margin-left: 0.98cm !important;
        text-align: left !important;
        margin-top: 2pt;
        margin-bottom: 2pt;
    }
    .word-page p.title-bold {
        text-indent: 0 !important;
        text-align: center !important;
        font-weight: bold;
        margin-top: 4pt;
        margin-bottom: 4pt;
    }
    .custom-underline {
        text-decoration: underline;
    }
</style>
""", unsafe_allow_html=True)

if "current_draft" not in st.session_state:
    st.session_state.current_draft = ""
if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []

BUILTIN_TEMPLATES_DANG = {
    "📌 Mẫu Công văn chuẩn (HD 05-HD/VPTW)": """KÍNH GỬI: Các chi, đảng bộ trực thuộc.
1. Căn cứ ban hành...
2. Nội dung chỉ đạo, triển khai...
3. Tổ chức thực hiện và báo cáo kết quả...""",

    "📌 Đề cương Kế hoạch chuẩn (HD 05-HD/VPTW)": """KẾ HOẠCH Về việc...
I. MỤC ĐÍCH, YÊU CẦU
1. Mục đích
2. Yêu cầu
II. NỘI DUNG VÀ NHIỆM VỤ CỤ THỂ
1. Nhiệm vụ trọng tâm
2. Giải pháp thực hiện
III. TỔ CHỨC THỰC HIỆN
1. Phân công trách nhiệm
2. Tiến độ và thời gian hoàn thành""",

    "📌 Đề cương Báo cáo chuẩn (HD 05-HD/VPTW)": """BÁO CÁO Tình hình...
I. KẾT QUẢ ĐẠT ĐƯỢC
1. Công tác chỉ đạo, quán triệt
2. Kết quả thực hiện các nhiệm vụ
II. HẠN CHẾ, KHUYẾT ĐIỂM VÀ NGUYÊN NHÂN
1. Hạn chế, tồn tại
2. Nguyên nhân (chủ quan, khách quan)
III. PHƯƠNG HƯỚNG, NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI""",

    "📌 Mẫu Giấy mời chuẩn (HD 05-HD/VPTW)": """GIẤY MỜI Về việc...
Đảng ủy phường Nhơn Trạch trân trọng kính mời: ...
- Thời gian: Vào lúc ... giờ ..., ngày ... tháng ... năm 2026.
- Địa điểm: Hội trường Đảng ủy phường Nhơn Trạch.
- Chủ trì: Đồng chí Bí thư Đảng ủy phường.
- Nội dung: ...""",

    "📌 Mẫu Tờ trình chuẩn (HD 05-HD/VPTW)": """TỜ TRÌNH Về việc...
KÍNH GỬI: Ban Thường vụ / Cơ quan cấp trên.
I. SỰ CẦN THIẾT / CĂN CỨ TRÌNH
II. NỘI DUNG CHÍNH CỦA TỜ TRÌNH
III. ĐỀ XUẤT, KIẾN NGHỊ"""
}

BUILTIN_TEMPLATES_NN = {
    "📌 Mẫu Công văn chuẩn (NĐ 30/2020/NĐ-CP)": """KÍNH GỬI: Các phòng, ban, đơn vị trực thuộc.
1. Căn cứ thực hiện...
2. Nội dung giao nhiệm vụ/thông báo...
3. Yêu cầu báo cáo/thời hạn hoàn thành...""",

    "📌 Đề cương Kế hoạch chuẩn (NĐ 30/2020/NĐ-CP)": """KẾ HOẠCH Về việc...
I. MỤC ĐÍCH, YÊU CẦU
II. NỘI DUNG VÀ CHỈ TIÊU NHIỆM VỤ
III. TỔ CHỨC THỰC HIỆN VÀ KINH PHÍ""",

    "📌 Đề cương Báo cáo chuẩn (NĐ 30/2020/NĐ-CP)": """BÁO CÁO Kết quả thực hiện...
I. TÌNH HÌNH VÀ KẾT QUẢ THỰC HIỆN
II. ĐÁNH GIÁ CHUNG (Ưu điểm, Hạn chế, Nguyên nhân)
III. NHIỆM VỤ GIẢI PHÁP VÀ ĐỀ XUẤT, KIẾN NGHỊ""",

    "📌 Mẫu Giấy mời chuẩn (NĐ 30/2020/NĐ-CP)": """GIẤY MỜI Về việc...
Ủy ban nhân dân phường Nhơn Trạch trân trọng kính mời: ...
- Thời gian: Vào lúc ... giờ ..., ngày ... tháng ... năm 2026.
- Địa điểm: Phòng họp UBND phường.
- Chủ trì: Đồng chí Chủ tịch UBND phường.
- Nội dung: ...""",

    "📌 Mẫu Tờ trình chuẩn (NĐ 30/2020/NĐ-CP)": """TỜ TRÌNH Về việc...
KÍNH GỬI: Ủy ban nhân dân cấp trên / Cơ quan có thẩm quyền.
I. CĂN CỨ PHÁP LÝ VÀ SỰ CẦN THIẾT
II. NỘI DUNG ĐỀ XUẤT
III. DỰ THẢO NGHỊ QUYẾT/QUYẾT ĐỊNH KÈM THEO"""
}

def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return ""
    filename = uploaded_file.name.lower()
    text = ""
    try:
        if filename.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8", errors="ignore")
        elif filename.endswith(".docx"):
            doc = Document(uploaded_file)
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif filename.endswith(".doc"):
            raw = uploaded_file.read().decode("latin-1", errors="ignore")
            text = "".join([c for c in raw if c.isprintable() or c in ["\n", "\r", "\t"]])
            text = re.sub(r"\s+", " ", text)
        elif filename.endswith((".xlsx", ".xls")):
            if openpyxl and filename.endswith(".xlsx"):
                wb = openpyxl.load_workbook(uploaded_file, data_only=True)
                sheets_text = []
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        row_vals = [str(cell) for cell in row if cell is not None]
                        if row_vals:
                            sheets_text.append("\t".join(row_vals))
                text = "\n".join(sheets_text)
            else:
                text = f"[Tệp bảng tính {uploaded_file.name} đã tải lên]"
        elif filename.endswith(".pdf"):
            try:
                import pypdf
                reader = pypdf.PdfReader(uploaded_file)
                for page in reader.pages:
                    extracted = page.extract_text()
                    if extracted:
                        text += extracted + "\n"
            except Exception:
                text = f"[Tệp PDF {uploaded_file.name} đã tải lên]"
    except Exception as e:
        st.error(f"Lỗi đọc file {uploaded_file.name}: {e}")
    return text

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        r'<w:tblBorders %s>'
        r'  <w:top w:val="none"/>'
        r'  <w:left w:val="none"/>'
        r'  <w:bottom w:val="none"/>'
        r'  <w:right w:val="none"/>'
        r'  <w:insideH w:val="none"/>'
        r'  <w:insideV w:val="none"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(tblBorders)

def parse_html_to_docx(html_text):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    
    html_text = html_text.replace('&nbsp;', ' ')
    html_text = re.sub(r' +', ' ', html_text)
    
    full_blocks = list(re.finditer(r'<(table|p)([^>]*)>(.*?)</\1>', html_text, flags=re.DOTALL | re.IGNORECASE))
    
    if not full_blocks:
        lines = html_text.split('\n')
        for line in lines:
            line_clean = re.sub(r'<[^>]+>', '', line).strip()
            if line_clean:
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.first_line_indent = Cm(0.98)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.add_run(line_clean)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    for match in full_blocks:
        tag_type = match.group(1).lower()
        attrs = match.group(2)
        content = match.group(3)
        
        if tag_type == 'table':
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', content, flags=re.DOTALL | re.IGNORECASE)
            if not rows:
                continue
            
            parsed_rows = []
            for r in rows:
                cells = re.findall(r'<td([^>]*)>(.*?)</td>', r, flags=re.DOTALL | re.IGNORECASE)
                parsed_rows.append(cells)
                
            num_rows = len(parsed_rows)
            num_cols = max([len(r) for r in parsed_rows]) if parsed_rows else 2
            
            table = doc.add_table(rows=num_rows, cols=num_cols)
            remove_table_borders(table)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            col_widths = [Cm(8.5), Cm(7.5)]
            
            for r_idx, r_cells in enumerate(parsed_rows):
                for c_idx, (c_attrs, c_content) in enumerate(r_cells):
                    if c_idx < num_cols:
                        cell = table.cell(r_idx, c_idx)
                        if c_idx < len(col_widths):
                            cell.width = col_widths[c_idx]
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                        
                        lines = re.split(r'<br\s*/?>', c_content, flags=re.IGNORECASE)
                        for l_idx, line in enumerate(lines):
                            line = line.strip()
                            if l_idx == 0:
                                p = cell.paragraphs[0]
                            else:
                                p = cell.add_paragraph()
                                
                            p.paragraph_format.line_spacing = 1.0
                            p.paragraph_format.space_after = Pt(0)
                            p.paragraph_format.space_before = Pt(0)
                            
                            # Xử lý thụt lề cả khối Nơi nhận trong bảng
                            if 'noi-nhan' in c_attrs or 'noi-nhan' in attrs:
                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                p.paragraph_format.left_indent = Cm(0.98)
                                p.paragraph_format.first_line_indent = Cm(0)
                            elif 'text-align:center' in c_attrs or 'text-align: center' in c_attrs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            elif 'text-align:right' in c_attrs or 'text-align: right' in c_attrs:
                                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            else:
                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                    
                            tokens = re.split(r'(<b>.*?</b>|<strong>.*?</strong>|<i>.*?</i>|<em>.*?</em>|<span class="custom-underline">.*?</span>|<u>.*?</u>)', line, flags=re.DOTALL | re.IGNORECASE)
                            for token in tokens:
                                if not token:
                                    continue
                                is_bold = bool(re.match(r'<(b|strong)>', token, re.IGNORECASE))
                                is_italic = bool(re.match(r'<(i|em)>', token, re.IGNORECASE))
                                is_underline = bool(re.match(r'<(u|span)', token, re.IGNORECASE))
                                clean_t = re.sub(r'<[^>]+>', '', token)
                                if clean_t:
                                    run = p.add_run(clean_t)
                                    run.font.name = 'Times New Roman'
                                    run.font.size = Pt(12) if r_idx == 0 else Pt(13)
                                    if is_bold:
                                        run.bold = True
                                    if is_italic:
                                        run.italic = True
                                    if is_underline:
                                        run.underline = True
                                        
            gap_p = doc.add_paragraph()
            gap_p.paragraph_format.space_after = Pt(2)
            gap_p.paragraph_format.space_before = Pt(0)

        elif tag_type == 'p':
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            
            if 'kinh-gui' in attrs or 'title-bold' in attrs or 'text-align:center' in attrs or 'text-align: center' in attrs:
                p.paragraph_format.left_indent = Cm(0)
                p.paragraph_format.first_line_indent = Cm(0)
                if 'text-align:left' in attrs or 'text-align: left' in attrs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif 'noi-nhan' in attrs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.left_indent = Cm(0.98)
                p.paragraph_format.first_line_indent = Cm(0)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(0.98)
                
            lines = re.split(r'<br\s*/?>', content, flags=re.IGNORECASE)
            for l_idx, line in enumerate(lines):
                line = line.strip()
                if l_idx > 0:
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(6)
                    p.paragraph_format.space_after = Pt(6)
                    if 'kinh-gui' in attrs or 'title-bold' in attrs or 'text-align:center' in attrs or 'text-align: center' in attrs:
                        p.paragraph_format.left_indent = Cm(0)
                        p.paragraph_format.first_line_indent = Cm(0)
                        if 'text-align:left' in attrs or 'text-align: left' in attrs:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif 'noi-nhan' in attrs:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p.paragraph_format.left_indent = Cm(0.98)
                        p.paragraph_format.first_line_indent = Cm(0)
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p.paragraph_format.first_line_indent = Cm(0.98)
                        
                tokens = re.split(r'(<b>.*?</b>|<strong>.*?</strong>|<i>.*?</i>|<em>.*?</em>|<span class="custom-underline">.*?</span>|<u>.*?</u>)', line, flags=re.DOTALL | re.IGNORECASE)
                for token in tokens:
                    if not token:
                        continue
                    is_bold = bool(re.match(r'<(b|strong)>', token, re.IGNORECASE))
                    is_italic = bool(re.match(r'<(i|em)>', token, re.IGNORECASE))
                    is_underline = bool(re.match(r'<(u|span)', token, re.IGNORECASE))
                    clean_t = re.sub(r'<[^>]+>', '', token)
                    if clean_t:
                        run = p.add_run(clean_t)
                        run.font.name = 'Times New Roman'
                        if is_bold:
                            run.bold = True
                        if is_italic:
                            run.italic = True
                        if is_underline:
                            run.underline = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def clean_html_response(res_text):
    res_text = re.sub(r'<think>.*?</think>', '', res_text, flags=re.DOTALL)
    res_text = re.sub(r'```[a-zA-Z]*', '', res_text)
    res_text = res_text.replace('
