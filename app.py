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
# 1. CẤU HÌNH TRANG & ĐĂNG NHẬP
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
# 2. CẤU HÌNH CSS & LƯU KEY
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
    section[data-testid="stSidebar"] > div:first-child {
        padding-top: 0.5rem !important;
        padding-bottom: 1rem !important;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 0.2rem !important;
    }
    .block-container {
        padding-top: 0.6rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    div[data-testid="stVerticalBlock"] > div {
        gap: 0.35rem !important;
    }
    .app-header {
        background: linear-gradient(135deg, #7b0000 0%, #a81010 50%, #c41e1e 100%);
        border: 1px solid #e0a800;
        border-radius: 10px;
        padding: 16px 22px;
        margin-bottom: 16px;
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
        margin-left: 1.0cm !important;
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
    "📌 Mẫu Công văn chuẩn (HD 05-HD/VPTW)": "KÍNH GỬI: Các chi, đảng bộ trực thuộc.\n1. Căn cứ ban hành...\n2. Nội dung chỉ đạo, triển khai...\n3. Tổ chức thực hiện và báo cáo kết quả...",
    "📌 Đề cương Kế hoạch chuẩn (HD 05-HD/VPTW)": "KẾ HOẠCH Về việc...\nI. MỤC ĐÍCH, YÊU CẦU\n1. Mục đích\n2. Yêu cầu\nII. NỘI DUNG VÀ NHIỆM VỤ CỤ THỂ\n1. Nhiệm vụ trọng tâm\n2. Giải pháp thực hiện\nIII. TỔ CHỨC THỰC HIỆN\n1. Phân công trách nhiệm\n2. Tiến độ và thời gian hoàn thành",
    "📌 Đề cương Báo cáo chuẩn (HD 05-HD/VPTW)": "BÁO CÁO Tình hình...\nI. KẾT QUẢ ĐẠT ĐƯỢC\n1. Công tác chỉ đạo, quán triệt\n2. Kết quả thực hiện các nhiệm vụ\nII. HẠN CHẾ, KHUYẾT ĐIỂM VÀ NGUYÊN NHÂN\n1. Hạn chế, tồn tại\n2. Nguyên nhân (chủ quan, khách quan)\nIII. PHƯƠNG HƯỚNG, NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI",
    "📌 Mẫu Giấy mời chuẩn (HD 05-HD/VPTW)": "GIẤY MỜI Về việc...\nĐảng ủy phường Nhơn Trạch trân trọng kính mời: ...\n- Thời gian: Vào lúc ... giờ ..., ngày ... tháng ... năm 2026.\n- Địa điểm: Hội trường Đảng ủy phường Nhơn Trạch.\n- Chủ trì: Đồng chí Bí thư Đảng ủy phường.\n- Nội dung: ...",
    "📌 Mẫu Tờ trình chuẩn (HD 05-HD/VPTW)": "TỜ TRÌNH Về việc...\nKÍNH GỬI: Ban Thường vụ / Cơ quan cấp trên.\nI. SỰ CẦN THIẾT / CĂN CỨ TRÌNH\nII. NỘI DUNG CHÍNH CỦA TỜ TRÌNH\nIII. ĐỀ XUẤT, KIẾN NGHỊ"
}

BUILTIN_TEMPLATES_NN = {
    "📌 Mẫu Công văn chuẩn (NĐ 30/2020/NĐ-CP)": "KÍNH GỬI: Các phòng, ban, đơn vị trực thuộc.\n1. Căn cứ thực hiện...\n2. Nội dung giao nhiệm vụ/thông báo...\n3. Yêu cầu báo cáo/thời hạn hoàn thành...",
    "📌 Đề cương Kế hoạch chuẩn (NĐ 30/2020/NĐ-CP)": "KẾ HOẠCH Về việc...\nI. MỤC ĐÍCH, YÊU CẦU\nII. NỘI DUNG VÀ CHỈ TIÊU NHIỆM VỤ\nIII. TỔ CHỨC THỰC HIỆN VÀ KINH PHÍ",
    "📌 Đề cương Báo cáo chuẩn (NĐ 30/2020/NĐ-CP)": "BÁO CÁO Kết quả thực hiện...\nI. TÌNH HÌNH VÀ KẾT QUẢ THỰC HIỆN\nII. ĐÁNH GIÁ CHUNG (Ưu điểm, Hạn chế, Nguyên nhân)\nIII. NHIỆM VỤ GIẢI PHÁP VÀ ĐỀ XUẤT, KIẾN NGHỊ",
    "📌 Mẫu Giấy mời chuẩn (NĐ 30/2020/NĐ-CP)": "GIẤY MỜI Về việc...\nỦy ban nhân dân phường Nhơn Trạch trân trọng kính mời: ...\n- Thời gian: Vào lúc ... giờ ..., ngày ... tháng ... năm 2026.\n- Địa điểm: Phòng họp UBND phường.\n- Chủ trì: Đồng chí Chủ tịch UBND phường.\n- Nội dung: ...",
    "📌 Mẫu Tờ trình chuẩn (NĐ 30/2020/NĐ-CP)": "TỜ TRÌNH Về việc...\nKÍNH GỬI: Ủy ban nhân dân cấp trên / Cơ quan có thẩm quyền.\nI. CĂN CỨ PHÁP LÝ VÀ SỰ CẦN THIẾT\nII. NỘI DUNG ĐỀ XUẤT\nIII. DỰ THẢO NGHỊ QUYẾT/QUYẾT ĐỊNH KÈM THEO"
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
                                p.paragraph_format.left_indent = Cm(1.0)
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
                p.paragraph_format.left_indent = Cm(1.0)
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
                        p.paragraph_format.left_indent = Cm(1.0)
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
    res_text = re.sub(r'```[a-zA-Z]*', '', res_text)
    res_text = res_text.replace('
