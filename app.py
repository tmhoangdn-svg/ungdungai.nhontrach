import streamlit as st
import pandas as pd
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
from datetime import datetime

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ==============================================================================
# 1. CẤU HÌNH TRANG & ĐĂNG NHẬP (GOOGLE SHEET)
# ==============================================================================
st.set_page_config(
    page_title="Phần mềm Cụ thể hóa Văn bản Hành chính",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1YHUgWJs3ZNH_6MVYI2Kwowsh7r0XVYaCXopvw1aD0FU/export?format=csv&gid=901150668"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzWB6-PRwFkezGzSjS29lrNBVnf03Dy0W1P4S0iDjJ9pIqgD5mDa-qKtc4NTw--IWoPgg/exec"

def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "user_info" not in st.session_state:
        st.session_state.user_info = {}
    if "reg_missing" not in st.session_state:
        st.session_state.reg_missing = []

    if not st.session_state.logged_in:
        _, center_col, _ = st.columns([1, 1.2, 1])
        with center_col:
            st.markdown("""
            <div style="text-align: center; margin-top: 25px; margin-bottom: 20px;">
                <span style="font-size: 42px;">🔐</span>
                <h2 style="margin: 8px 0 0 0; font-size: 24px; font-weight: 800; color: #ffd700;">ĐĂNG NHẬP HỆ THỐNG</h2>
                <p style="color: #a0aec0; font-size: 13px; margin-top: 4px;">Phần mềm Cụ thể hóa Văn bản Hành chính</p>
            </div>
            """, unsafe_allow_html=True)

            tab_login, tab_register, tab_forgot = st.tabs(["Đăng nhập", "Đăng ký tài khoản", "Quên mật khẩu"])

            with tab_login:
                with st.form("login_form"):
                    username = st.text_input("Tên đăng nhập")
                    password = st.text_input("Mật khẩu", type="password")
                    btn_login = st.form_submit_button("Đăng nhập", use_container_width=True)

                if btn_login:
                    if username == "admin" and password == "Adminai":
                        st.success("Xin chào Admin!")
                        st.session_state.logged_in = True
                        st.session_state.user_info = {"username": "admin", "fullname": "Quản trị viên", "email_phone": "N/A"}
                        st.rerun()
                    else:
                        try:
                            df = pd.read_csv(SHEET_CSV_URL, skiprows=2)
                            df.columns = [c.strip() for c in df.columns]
                            user_match = df[(df['username'].astype(str) == username) & (df['password'].astype(str) == password)]
                            
                            if not user_match.empty:
                                user_data = user_match.iloc[0]
                                st.success(f"Đăng nhập thành công! Chào mừng {user_data['fullname']}")
                                st.session_state.logged_in = True
                                
                                contact_val = "Chưa cập nhật"
                                if 'email_phone' in user_data and pd.notna(user_data['email_phone']):
                                    contact_val = str(user_data['email_phone'])
                                elif 'Email/SĐT' in user_data and pd.notna(user_data['Email/SĐT']):
                                    contact_val = str(user_data['Email/SĐT'])

                                st.session_state.user_info = {
                                    "username": username,
                                    "fullname": user_data['fullname'],
                                    "email_phone": contact_val
                                }
                                st.rerun()
                            else:
                                st.error("Tên đăng nhập hoặc mật khẩu không chính xác!")
                        except Exception:
                            st.error("Chưa thể kết nối đến dữ liệu tài khoản.")

            with tab_register:
                lbl_user = "Tên đăng nhập mới" + (" :red[*]" if "user" in st.session_state.reg_missing else "")
                lbl_name = "Họ và tên" + (" :red[*]" if "name" in st.session_state.reg_missing else "")
                lbl_contact = "Email hoặc Số điện thoại" + (" :red[*]" if "contact" in st.session_state.reg_missing else "")
                lbl_pass = "Mật khẩu mới" + (" :red[*]" if "pass" in st.session_state.reg_missing else "")
                lbl_conf = "Xác nhận mật khẩu" + (" :red[*]" if "conf" in st.session_state.reg_missing else "")

                with st.form("register_form"):
                    new_user = st.text_input(lbl_user, key="reg_user")
                    new_name = st.text_input(lbl_name, key="reg_name")
                    new_contact = st.text_input(lbl_contact, key="reg_contact")
                    new_pass = st.text_input(lbl_pass, type="password", key="reg_pass")
                    confirm_pass = st.text_input(lbl_conf, type="password", key="reg_conf")
                    btn_register = st.form_submit_button("Đăng ký", use_container_width=True)

                if btn_register:
                    missing = []
                    if not new_user.strip(): missing.append("user")
                    if not new_name.strip(): missing.append("name")
                    if not new_contact.strip(): missing.append("contact")
                    if not new_pass.strip(): missing.append("pass")
                    if not confirm_pass.strip(): missing.append("conf")

                    st.session_state.reg_missing = missing

                    if missing:
                        st.warning("Vui lòng điền đầy đủ các thông tin có dấu (*) đỏ!")
                        st.rerun()
                    elif new_pass != confirm_pass:
                        st.error("Mật khẩu xác nhận không khớp!")
                    else:
                        st.session_state.reg_missing = []
                        payload = {
                            "action": "register",
                            "username": new_user,
                            "password": new_pass,
                            "fullname": new_name,
                            "email_phone": new_contact,
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        try:
                            res = requests.post(WEB_APP_URL, json=payload)
                            if res.status_code == 200:
                                st.success("Đăng ký thành công! Vui lòng quay lại tab Đăng nhập.")
                            else:
                                st.error("Lỗi khi tạo tài khoản.")
                        except Exception:
                            st.error("Không thể kết nối máy chủ đăng ký.")

            with tab_forgot:
                with st.form("forgot_form"):
                    fg_user = st.text_input("Nhập Tên đăng nhập của bạn", key="fg_u")
                    fg_contact = st.text_input("Nhập Email hoặc Số điện thoại đã đăng ký", key="fg_c")
                    fg_new_pass = st.text_input("Mật khẩu mới", type="password", key="fg_p1")
                    fg_confirm_pass = st.text_input("Xác nhận mật khẩu mới", type="password", key="fg_p2")
                    btn_forgot = st.form_submit_button("Đặt lại mật khẩu", use_container_width=True)

                if btn_forgot:
                    if not fg_user or not fg_contact or not fg_new_pass:
                        st.warning("Vui lòng điền đầy đủ thông tin!")
                    elif fg_new_pass != fg_confirm_pass:
                        st.error("Mật khẩu xác nhận không khớp!")
                    else:
                        try:
                            df = pd.read_csv(SHEET_CSV_URL, skiprows=2)
                            df.columns = [c.strip() for c in df.columns]
                            col_email = 'email_phone' if 'email_phone' in df.columns else 'Email/SĐT'
                            match = df[(df['username'].astype(str) == fg_user) & (df[col_email].astype(str) == fg_contact)]
                            
                            if not match.empty:
                                payload = {
                                    "action": "update_password",
                                    "username": fg_user,
                                    "new_password": fg_new_pass
                                }
                                res = requests.post(WEB_APP_URL, json=payload)
                                if res.status_code == 200:
                                    st.success("Đổi mật khẩu thành công! Vui lòng quay lại tab Đăng nhập.")
                                else:
                                    st.error("Lỗi khi cập nhật mật khẩu.")
                            else:
                                st.error("Tên đăng nhập và Email/SĐT không khớp!")
                        except Exception:
                            st.error("Không thể xác minh thông tin.")

        return False
    return True

if not check_login():
    st.stop()

# ==============================================================================
# 2. CẤU HÌNH CSS & LƯU TRỮ KEY
# ==============================================================================
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

st.markdown("""
<style>
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    div[data-testid="stVerticalBlock"] > div {
        gap: 0.4rem !important;
    }
    
    .app-header {
        background: linear-gradient(135deg, #7b0000 0%, #a81010 50%, #c41e1e 100%);
        border: 1px solid #e0a800;
        border-radius: 10px;
        padding: 16px 22px;
        margin-bottom: 18px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.4);
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .app-header-title {
        color: #ffffff;
        font-size: 21px;
        font-weight: 800;
        letter-spacing: 0.5px;
        text-shadow: 1px 1px 3px rgba(0,0,0,0.6);
        margin: 0;
    }
    .app-header-sub {
        color: #ffd700;
        font-size: 12px;
        margin-top: 3px;
        font-weight: 500;
        letter-spacing: 0.3px;
    }
    .section-badge {
        background: linear-gradient(90deg, #d4af37 0%, #f3e5ab 100%);
        color: #4a2c00;
        font-size: 10.5px;
        font-weight: bold;
        padding: 2px 7px;
        border-radius: 4px;
        text-transform: uppercase;
        display: inline-block;
        margin-bottom: 4px;
    }
    .word-page {
        background-color: #ffffff;
        color: #000000;
        width: 100%;
        max-height: 480px;
        overflow-y: auto;
        padding: 25px 35px;
        border: 1px solid #d3d3d3;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
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
        font-size: 11pt !important;
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
                                    run.font.size = Pt(11) if ('noi-nhan' in c_attrs or 'noi-nhan' in attrs) else (Pt(12) if r_idx == 0 else Pt(13))
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
                        run.font.size = Pt(11) if 'noi-nhan' in attrs else Pt(13)
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
    tick = chr(96)
    res_text = re.sub(rf'{tick}{{1,3}}[a-zA-Z]*', '', res_text)
    res_text = res_text.replace(tick, '')
    res_text = re.sub(r'', '', res_text, flags=re.DOTALL)
    res_text = re.sub(r'^(Dưới đây là|Đây là|Gửi bạn|Sau đây là).*\n', '', res_text, flags=re.IGNORECASE)
    res_text = res_text.replace('&nbsp;', ' ')
    return res_text.strip()

def call_gemini_api(api_key, model_name, prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    response = requests.post(url, headers=headers, json=payload, timeout=60)
    if response.status_code == 200:
        res_data = response.json()
        return res_data['candidates'][0]['content']['parts'][0]['text']
    else:
        raise Exception(f"Lỗi API Gemini ({response.status_code}): {response.text}")

# ==============================================================================
# 3. THANH SIDEBAR
# ==============================================================================
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding-top: 0px; padding-bottom: 4px;">
        <span style="font-size: 28px;">🏛️</span>
        <h3 style="color: #ffd700; margin: 2px 0 0 0; font-size: 17px; font-weight: bold;">HỆ THỐNG ĐIỀU HÀNH</h3>
        <p style="color: #a0aec0; font-size: 11px; margin-top: 2px;">Chuẩn Thể Thức Đảng & Nhà Nước</p>
    </div>
    """, unsafe_allow_html=True)
    st.write("---")
    
    khoi_van_ban = st.radio(
        "Chọn Khối văn bản:",
        ["Khối Đảng", "Khối Nhà nước"],
        index=0
    )
    
    if khoi_van_ban == "Khối Đảng":
        the_thuc_note = "Hướng dẫn 05-HD/VPTW của Văn phòng Trung ương Đảng"
        builtin_dict = BUILTIN_TEMPLATES_DANG
        default_agency = "ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH"
    else:
        the_thuc_note = "Nghị định 30/2020/NĐ-CP của Chính phủ"
        builtin_dict = BUILTIN_TEMPLATES_NN
        default_agency = "UBND PHƯƠNG NHƠN TRẠCH"
    
    st.info(f"📌 **Áp dụng:** {the_thuc_note}")
    st.write("---")

    st.subheader("⚙️ Cấu hình AI")
    ai_engine = st.selectbox("AI xử lý chính", ["Google Gemini"])
    
    gemini_api_key = config_data.get("gemini_key", "")
    gemini_api_key = st.text_input("Gemini API key", value=gemini_api_key, type="password")
    gemini_model = st.selectbox("Model", ["gemini-3.6-flash", "gemini-1.5-pro", "gemini-1.5-flash"])

    if st.button("💾 Lưu API Key vĩnh viễn", use_container_width=True):
        new_config = {
            "gemini_key": gemini_api_key
        }
        if save_config(new_config):
            st.success("Đã lưu API Key!")
        else:
            st.error("Lỗi khi lưu cấu hình!")

    st.write("---")
    user_info = st.session_state.get("user_info", {})
    with st.popover(f"👤 Tài khoản ({user_info.get('fullname', 'User')})"):
        st.markdown(f"**Họ tên:** {user_info.get('fullname')}")
        st.markdown(f"**Username:** {user_info.get('username')}")
        st.markdown(f"**Liên hệ:** {user_info.get('email_phone')}")
        st.write("---")
        if st.button("🚪 Đăng xuất", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.user_info = {}
            st.rerun()

# ==============================================================================
# 4. GIAO DIỆN CHÍNH
# ==============================================================================
st.markdown("""
<div class="app-header">
    <div>
        <div class="app-header-title">🏛️ PHẦN MỀM CỤ THỂ HÓA VĂN BẢN HÀNH CHÍNH</div>
        <div class="app-header-sub">HỆ THỐNG HỖ TRỢ BIÊN SOẠN & XỬ LÝ VĂN KIỆN ĐẢNG - CHÍNH QUYỀN TỰ ĐỘNG BẰNG AI</div>
    </div>
    <div style="font-size: 30px; font-weight: bold; color: #ffd700; text-shadow: 1px 1px 3px rgba(0,0,0,0.8);">VN</div>
</div>
""", unsafe_allow_html=True)

# BƯỚC 1 & 2
top_col1, top_col2 = st.columns(2)

with top_col1:
    st.markdown('<span class="section-badge">BƯỚC 1</span> <b>File nguồn & Loại văn bản</b>', unsafe_allow_html=True)
    uploaded_source = st.file_uploader(
        "Tải file nguồn/Đề cương (.docx, .pdf, .png, .jpg...):",
        type=["doc", "docx", "xls", "xlsx", "pdf", "txt", "png", "jpg", "jpeg"],
        key="source_file"
    )
    source_text = read_uploaded_file(uploaded_source)

    doc_type = st.selectbox(
        "Chọn Loại văn bản đầu ra:", 
        ["Kế hoạch", "Công văn", "Báo cáo", "Giấy mời", "Tờ trình", "Hướng dẫn", "Quy chế", "Quyết định", "Thông báo", "Khác"]
    )

with top_col2:
    st.markdown('<span class="section-badge">BƯỚC 2</span> <b>Yêu cầu & Cơ quan ban hành</b>', unsafe_allow_html=True)
    req_detail = st.text_area(
        "Anh muốn cụ thể hóa như thế nào?:",
        placeholder="Soạn thảo văn bản theo yêu cầu chỉ đạo...",
        height=68
    )

    co_quan_ban_hanh = st.text_input("Cơ quan ban hành dự thảo:", value=default_agency)

# BƯỚC 3
st.markdown('<br>', unsafe_allow_html=True)
col3_1, col3_2 = st.columns(2)

with col3_1:
    st.markdown('<span class="section-badge">BƯỚC 3</span> <b>File mẫu riêng & Mẫu gợi ý chuẩn</b>', unsafe_allow_html=True)
    uploaded_sample = st.file_uploader(
        "Tải file mẫu riêng (Chỉ lấy thể thức/khung mẫu):",
        type=["doc", "docx"],
        key="sample_file"
    )
    sample_text = read_uploaded_file(uploaded_sample)

with col3_2:
    st.markdown('<span style="color: #e53e3e; font-size: 13px;">📌</span> <b>Mẫu gợi ý / Đề cương chuẩn:</b>', unsafe_allow_html=True)
    selected_builtin = st.selectbox(
        "Mẫu gợi ý / Đề cương chuẩn:",
        ["(Không chọn mẫu gợi ý)"] + list(builtin_dict.keys()),
        label_visibility="collapsed"
    )

    if selected_builtin != "(Không chọn mẫu gợi ý)" and not sample_text:
        sample_text = builtin_dict[selected_builtin]
        with st.expander("👁️ Xem trước Đề cương/Mẫu gợi ý đã chọn:"):
            st.code(sample_text, language="text")

st.markdown("<br>", unsafe_allow_html=True)
if st.button("⚡ PHÂN TÍCH & CỤ THỂ HÓA VĂN BẢN", type="primary", use_container_width=True):
    if not source_text.strip():
        st.warning("Vui lòng tải lên Văn bản nguồn!")
    else:
        with st.spinner("Gemini đang xử lý..."):
            processed_source = source_text
            processed_sample = sample_text

            if khoi_van_ban == "Khối Đảng":
                header_left = f"ĐẢNG BỘ THÀNH PHỐ ĐỒNG NAI<br><b>{co_quan_ban_hanh}</b><br>*<br>Số: ...-{'CV/ĐU' if doc_type == 'Công văn' else ('KH/ĐU' if doc_type == 'Kế hoạch' else ('BC/ĐU' if doc_type == 'Báo cáo' else ('GM/ĐU' if doc_type == 'Giấy mời' else ('TTr/ĐU' if doc_type == 'Tờ trình' else 'HD/ĐU'))))}<br><i>{'V/v ...' if doc_type in ['Công văn', 'Tờ trình'] else ''}</i>"
                header_right = '<b><u>ĐẢNG CỘNG SẢN VIỆT NAM</u></b><br><i>Nhơn Trạch, ngày ... tháng ... năm 2026</i>'
                signer_position = "<b>T/M BAN THƯỜNG VỤ</b><br><b>BÍ THƯ</b>"
            else:
                header_left = f"ỦY BÀN NHÂN DÂN<br><b>{co_quan_ban_hanh}</b><br>────────<br>Số: .../UBND<br><i>{'V/v ...' if doc_type in ['Công văn', 'Tờ trình'] else ''}</i>"
                header_right = '<b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM<br><u>Độc lập - Tự do - Hạnh phúc</u></b><br>───────────<br><i>Nhơn Trạch, ngày ... tháng ... năm 2026</i>'
                signer_position = "<b>TM. ỦY BAN NHÂN DÂN</b><br><b>CHỦ TỊCH</b>"

            if doc_type == "Công văn":
                layout_structure = """
                2. Phần Kính gửi:
                   <p class="kinh-gui" style="text-align:center; text-indent:0;"><b>Kính gửi:</b> Các đơn vị/chi bộ trực thuộc (hoặc cơ quan liên quan).</p>
                3. Phần Thân công văn:
                   <p class="body-p">Căn cứ...</p>
                   <p class="body-p">Nội dung chỉ đạo...</p>
                """
            elif doc_type == "Giấy mời":
                layout_structure = f"""
                2. Phần Tên loại văn bản:
                   <p class="title-bold" style="text-align:center;"><b>GIẤY MỜI</b><br><b>Về việc ...</b></p>
                3. Phần Thân giấy mời:
                   <p class="body-p"><b>{co_quan_ban_hanh} trân trọng kính mời:</b> ...</p>
                   <p class="body-p"><b>- Thời gian:</b> Vào lúc ... giờ ..., ngày ... tháng ... năm 2026.</p>
                   <p class="body-p"><b>- Địa điểm:</b> Hội trường/Phòng họp {co_quan_ban_hanh}.</p>
                   <p class="body-p"><b>- Chủ trì:</b> Lãnh đạo đơn vị.</p>
                   <p class="body-p"><b>- Nội dung:</b> ...</p>
                """
            elif doc_type == "Tờ trình":
                layout_structure = """
                2. Phần Tên loại văn bản & Kính gửi:
                   <p class="title-bold" style="text-align:center;"><b>TỜ TRÌNH</b><br><b>Về việc ...</b></p>
                   <p class="kinh-gui" style="text-align:center; text-indent:0;"><b>Kính gửi:</b> Cơ quan/Cấp có thẩm quyền cấp trên.</p>
                3. Phần Thân tờ trình:
                   <p class="body-p">Căn cứ...</p>
                   <p class="body-p"><b>I. SỰ CẦN THIẾT / CĂN CỨ TRÌNH</b></p>
                   <p class="body-p">...</p>
                   <p class="body-p"><b>II. NỘI DUNG TRÌNH</b></p>
                   <p class="body-p">...</p>
                   <p class="body-p"><b>III. ĐỀ XUẤT, KIẾN NGHỊ</b></p>
                   <p class="body-p">...</p>
                """
            else:
                layout_structure = f"""
                2. Phần Tên loại văn bản & Trích yếu (BẮT BUỘC CĂN GIỮA, TUYỆT ĐỐI KHÔNG CÓ PHẦN KÍNH GỬI):
                   <p class="title-bold" style="text-align:center;"><b>{doc_type.upper()}</b><br><b>Về việc ...</b></p>
                3. Phần Thân văn bản (BẮT BUỘC CHIA MỤC LA MÃ I., II., III. NẾU CÓ ĐỀ CƯƠNG THÌ BÁM SÁT 100% ĐỀ CƯƠNG TẢI LÊN HOẶC MẪU GỢI Ý):
                   <p class="body-p"><b>I. MỤC ĐÍCH, YÊU CẦU / ĐÁNH GIÁ TÌNH HÌNH</b></p>
                   <p class="body-p">...</p>
                   <p class="body-p"><b>II. NỘI DUNG, PHƯƠNG HƯỚNG / KẾT QUẢ ĐẠT ĐƯỢC</b></p>
                   <p class="body-p">...</p>
                   <p class="body-p"><b>III. TỔ CHỨC THỰC HIỆN / ĐỀ XUẤT KIẾN NGHỊ</b></p>
                   <p class="body-p">...</p>
                """

            prompt = f"""
            Bạn là chuyên viên cao cấp về soạn thảo văn bản hành chính Việt Nam.
            NHIỆM VỤ: Soạn thảo văn bản {doc_type} dưới dạng MÃ HTML CHUẨN ĐÉT TRANG WORD A4.
            BẮT BUỘC TRẢ VỀ TRỰC TIẾP MÃ HTML. KHÔNG SUY NGHĨ LOẰNG NGOẰNG.

            QUY TẮC THỂ THỨC BẮT BUỘC THEO LOẠI VĂN BẢN ({doc_type}):
            - Khối văn bản hiện tại: {khoi_van_ban} ({the_thuc_note})
            - Loại văn bản hiện tại: {doc_type}
            - TUYỆT ĐỐI NẾU KHÔNG PHẢI LÀ CÔNG VĂN VÀ TỜ TRÌNH THÌ KHÔNG ĐƯỢC CÓ MỤC "KÍNH GỬI".
            - PHẦN NƠI NHẬN BẮT BUỘC DÙNG THẺ <p class="noi-nhan"> CHO CẢ TIÊU ĐỀ NƠI NHẬN VÀ CÁC DÒNG LIỆT KÊ.
            - TUYỆT ĐỐI KHÔNG DÙNG KÝ TỰ &nbsp; ĐỂ TẠO KHOẢNG TRẮNG.

            CẤU TRÚC HTML MẪU:
            1. Bảng Tiêu ngữ & Tên cơ quan (Dùng <table> 2 cột KHÔNG VIỀN):
               <table style="width:100%;">
                 <tr>
                   <td style="text-align:center; width:52%;">{header_left}</td>
                   <td style="text-align:center; width:48%;">{header_right}</td>
                 </tr>
               </table>

            {layout_structure}

            4. Bảng Nơi nhận & Chữ ký (Dùng <table> 2 cột KHÔNG VIỀN ở cuối):
               <table style="width:100%; margin-top:25px;">
                 <tr>
                   <td style="width:50%;" class="noi-nhan">
                     <p class="noi-nhan"><b><u><i>Nơi nhận:</i></u></b></p>
                     <p class="noi-nhan">- Như trên;</p>
                     <p class="noi-nhan">- Lưu VP.</p>
                   </td>
                   <td style="text-align:center; width:50%;">
                     {signer_position}<br><br><br><br>
                     <b>Họ và Tên</b>
                   </td>
                 </tr>
               </table>

            QUY TẮC KỸ THUẬT:
            - KHÔNG DÙNG MARKDOWN (*, **, #). CHỈ DÙNG THẺ HTML: <table>, <tr>, <td>, <p>, <b>, <i>, <u>, <br>.
            - KHÔNG LỜI CHÀO/THOẠI. CHỈ TRẢ VỀ DUY NHẤT MÃ HTML.

            THÔNG TIN CHI TIẾT:
            - Khối văn bản: {khoi_van_ban} ({the_thuc_note})
            - Cơ quan ban hành: {co_quan_ban_hanh}
            - Yêu cầu cụ thể hóa: {req_detail}

            NỘI DUNG VĂN BẢN NGUỒN / ĐỀ CƯƠNG (CẤP TRÊN):
            ---
            {processed_source}
            ---

            VĂN BẢN MẪU / ĐỀ CƯƠNG THAM KHẢO (NẾU CÓ):
            ---
            {processed_sample if processed_sample else "Áp dụng cấu trúc chuẩn hành chính " + the_thuc_note}
            ---
            """
            try:
                if not gemini_api_key:
                    st.error("Vui lòng nhập Gemini API Key ở cột bên trái!")
                else:
                    res = call_gemini_api(gemini_api_key, gemini_model, prompt)
                    st.session_state.current_draft = clean_html_response(res)
                    st.session_state.chat_messages = []
            except Exception as e:
                st.error(f"Lỗi khi xử lý AI: {e}")

if st.session_state.current_draft:
    st.divider()
    bottom_col1, bottom_col2 = st.columns([3, 2])
    
    with bottom_col1:
        st.markdown("##### 📄 Bản dự thảo trang Word (A4)")
        st.markdown(f'<div class="word-page">{st.session_state.current_draft}</div>', unsafe_allow_html=True)
        
        docx_bytes = parse_html_to_docx(st.session_state.current_draft)
        file_download_name = f"Du_thao_{doc_type}_{co_quan_ban_hanh}.docx".replace(" ", "_")
        st.download_button(
            label="📥 TẢI VỀ FILE WORD (.DOCX)",
            data=docx_bytes,
            file_name=file_download_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

    with bottom_col2:
        st.markdown("##### 💬 Chat AI sửa đổi (Gemini)")
        st.caption("Nhập yêu cầu (VD: 'Sửa căn cứ 1', 'Bỏ mục II') để AI cập nhật trực tiếp lên trang Word bên trái.")

        chat_container = st.container(height=360)
        with chat_container:
            for msg in st.session_state.chat_messages:
                with st.chat_message(msg["role"]):
                    st.write(msg["content"])

        if user_edit_req := st.chat_input("Nhập yêu cầu chỉnh sửa văn bản..."):
            st.session_state.chat_messages.append({"role": "user", "content": user_edit_req})
            
            with chat_container:
                with st.chat_message("user"):
                    st.write(user_edit_req)

                with st.chat_message("assistant"):
                    with st.spinner("AI đang cập nhật lại văn bản..."):
                        edit_prompt = f"""
                        Bạn là biên tập viên văn bản hành chính.
                        DƯỚI ĐÂY LÀ MÃ HTML BẢN DỰ THẢO VĂN BẢN HIỆN TẠI:
                        ---
                        {st.session_state.current_draft}
                        ---

                        YÊU CẦU CHỈNH SỬA TỪ NGUỜI DÙNG:
                        "{user_edit_req}"

                        QUY TẮC CỐ ĐỊNH:
                        1. Hãy sửa đổi trực tiếp vào MÃ HTML BẢN DỰ THẢO HIỆN TẠI theo đúng yêu cầu trên.
                        2. Giữ nguyên toàn bộ cấu trúc các thẻ HTML (<table>, <tr>, <td>, <p class="kinh-gui">, <p class="noi-nhan">, <p class="body-p">, <b>, <i>, <u>).
                        3. Tuyệt đối KHÔNG dùng ký tự Markdown (*, **, #).
                        4. Tuyệt đối KHÔNG trả lời bằng lời chào/thoại. CHỈ TRẢ VỀ TOÀN BỘ MÃ HTML SAU KHI SỬA.
                        5. TUYỆT ĐỐI KHÔNG DÙNG &nbsp; TẠO KHOẢNG TRẮNG.
                        """
                        try:
                            updated_res = call_gemini_api(gemini_api_key, gemini_model, edit_prompt)
                            cleaned_res = clean_html_response(updated_res)
                            st.session_state.current_draft = cleaned_res
                            st.session_state.chat_messages.append({"role": "assistant", "content": "✅ Đã cập nhật văn bản lên trang Word!"})
                            st.rerun()
                        except Exception as e:
                            st.error(f"Lỗi khi sửa văn bản: {e}")
