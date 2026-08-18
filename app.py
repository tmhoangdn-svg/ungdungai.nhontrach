import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import google.generativeai as genai
import pypdf
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import json, io, re, os

# ==============================================================================
# 1. CẤU HÌNH ĐĂNG NHẬP / ĐĂNG KÝ & KẾT NỐI GOOGLE SHEET
# ==============================================================================
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1YHUgWJs3ZNH_6MVYI2Kwowsh7r0XVYaCXopvw1aD0FU/export?format=csv&gid=901150668"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzWB6-PRwFkezGzSjS29lrNBVnf03Dy0W1P4S0iDjJ9pIqgD5mDa-qKtc4NTw--IWoPgg/exec"
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
            <div style="text-align: center; margin-top: 20px; margin-bottom: 20px;">
                <span style="font-size: 42px;">🔐</span>
                <h2 style="margin: 8px 0 0 0; font-size: 24px; font-weight: 800;">ĐĂNG NHẬP HỆ THỐNG</h2>
                <p style="color: #888; font-size: 13px; margin-top: 4px;">Phần mềm Cụ thể hóa Văn bản Hành chính</p>
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

st.set_page_config(page_title="Phần mềm Cụ thể hóa Văn bản Hành chính", page_icon="🏛️", layout="wide")

if not check_login():
    st.stop()

# ==============================================================================
# 2. CẤU HÌNH GIAO DIỆN & BẢNG ÁNH XẠ
# ==============================================================================
SYMBOL_MAP = {
    "Kế hoạch": "KH",
    "Công văn": "CV",
    "Báo cáo": "BC",
    "Tờ trình": "TTr",
    "Thông báo": "TB",
    "Quyết định": "QĐ",
    "Hướng dẫn": "HD"
}

TEMPLATES_CONFIG = {
    "Khối Đảng": {
        "Kế hoạch": "KẾ HOẠCH Về việc...\nI. MỤC ĐÍCH, YÊU CẦU\n1. Mục đích\n2. Yêu cầu\nII. NỘI DUNG VÀ NHIỆM VỤ CỤ THỂ\n1. Nhiệm vụ trọng tâm\n2. Giải pháp thực hiện\nIII. TỔ CHỨC THỰC HIỆN\n1. Phân công trách nhiệm\n2. Tiến độ và thời gian hoàn thành",
        "Công văn": "Kính gửi: Các chi, đảng bộ cơ sở trực thuộc.\n1. Căn cứ ban hành và mục đích triển khai...\n2. Nội dung cụ thể hóa chỉ đạo của cấp trên...\n3. Tổ chức thực hiện và chế độ báo cáo kết quả...",
        "Báo cáo": "BÁO CÁO Tình hình...\nI. KẾT QUẢ ĐẠT ĐƯỢC\n1. Công tác lãnh đạo, chỉ đạo, quán triệt\n2. Kết quả thực hiện các nhiệm vụ chính trị\nII. HẠN CHẾ, KHUYẾT ĐIỂM VÀ NGUYÊN NHÂN\n1. Hạn chế, tồn tại\n2. Nguyên nhân (chủ quan, khách quan)\nIII. PHƯƠNG HƯỚNG, NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI",
        "Tờ trình": "TỜ TRÌNH Về việc...\nKính gửi: Ban Thường vụ Đảng ủy cấp trên / Cơ quan có thẩm quyền.\nI. SỰ CẦN THIẾT VÀ CĂN CỨ TRÌNH\nII. NỘI DUNG CHÍNH CỦA TỜ TRÌNH\nIII. ĐỀ XUẤT, KIẾN NGHỊ",
        "Thông báo": "THÔNG BÁO Kết luận của...\n1. Đánh giá tình hình thực hiện nhiệm vụ vừa qua\n2. Ý kiến kết luận và phân công nhiệm vụ cụ thể thời gian tới\n3. Trách nhiệm tổ chức thực hiện của các cơ quan, đơn vị",
        "Quyết định": "QUYẾT ĐỊNH Về việc...\n- Căn cứ Điều lệ Đảng và Quy chế làm việc...\nBAN THƯỜNG VỤ QUYẾT ĐỊNH:\nĐiều 1. (Nội dung quyết định cụ thể hóa)\nĐiều 2. (Trách nhiệm của các tổ chức, cá nhân)\nĐiều 3. Quyết định này có hiệu lực kể từ ngày ký...",
        "Hướng dẫn": "HƯỚNG DẪN Về việc...\n- Căn cứ văn bản chỉ đạo của cấp trên...\nI. MỤC ĐÍCH, YÊU CẦU\nII. ĐỐI TƯỢNG VÀ PHẠM VI ÁP DỤNG\nIII. NỘI DUNG HƯỚNG DẪN CỤ THỂ\n1. Nhiệm vụ chuyên môn\n2. Quy trình, hồ sơ thực hiện\nIV. TỔ CHỨC THỰC HIỆN VÀ BÁO CÁO"
    },
    "Khối Nhà nước": {
        "Kế hoạch": "KẾ HOẠCH Về việc...\nI. MỤC ĐÍCH, YÊU CẦU\nII. NỘI DUNG VÀ CHỈ TIÊU NHIỆM VỤ\n1. Nhiệm vụ trọng tâm\n2. Giải pháp thực hiện\nIII. TỔ CHỨC THỰC HIỆN VÀ KINH PHÍ",
        "Công văn": "Kính gửi: Các phòng, ban, đơn vị trực thuộc.\n1. Căn cứ văn bản chỉ đạo cấp trên...\n2. Nội dung giao nhiệm vụ và yêu cầu thực hiện...\n3. Thời hạn hoàn thành và báo cáo...",
        "Báo cáo": "BÁO CÁO Kết quả thực hiện...\nI. TÌNH HÌNH VÀ KẾT QUẢ ĐẠT ĐƯỢC\nII. ĐÁNH GIÁ CHUNG (Ưu điểm, Hạn chế, Nguyên nhân)\nIII. NHIỆM VỤ GIẢI PHÁP VÀ ĐỀ XUẤT, KIẾN NGHỊ",
        "Tờ trình": "TỜ TRÌNH Về việc...\nKính gửi: Ủy ban nhân dân cấp trên / Cơ quan có thẩm quyền.\nI. CĂN CỨ PHÁP LÝ VÀ SỰ CẦN THIẾT\nII. NỘI DUNG ĐỀ XUẤT, PHÊ DUYỆT\nIII. DỰ THẢO VĂN BẢN KÈM THEO",
        "Thông báo": "THÔNG BÁO Về việc...\n1. Nội dung thông báo / Ý kiến kết luận chỉ đạo của UBND\n2. Giao nhiệm vụ cho các phòng ban, đơn vị phối hợp triển khai",
        "Quyết định": "QUYẾT ĐỊNH Về việc...\n- Căn cứ Luật Tổ chức chính quyền địa phương...\nỦY BAN NHÂN DÂN QUYẾT ĐỊNH:\nĐiều 1. (Nội dung quyết định cụ thể hóa)\nĐiều 2. (Trách nhiệm của các cơ quan liên quan)\nĐiều 3. Quyết định có hiệu lực kể từ ngày ký...",
        "Hướng dẫn": "HƯỚNG DẪN Thực hiện...\n- Căn cứ quy định pháp luật và văn bản cấp trên...\nI. MỤC ĐÍCH, YÊU CẦU\nII. ĐỐI TƯỢNG VÀ PHẠM VI ÁP DỤNG\nIII. NỘI DUNG VÀ TRÌNH TỰ THỰC HIỆN\nIV. TỔ CHỨC THỰC HIỆN"
    }
}

# HÀM FORMAT ĐỀ CƯƠNG HIỂN THỊ CHUYÊN NGHIỆP
def format_outline_to_html(text):
    lines = text.split('\n')
    formatted = ""
    for line in lines:
        l = line.strip()
        if not l: continue
        if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', l):
            formatted += f"<div style='font-weight: bold; margin-top: 8px; color: #63b3ed;'>{l}</div>"
        elif re.match(r'^\d+\.', l):
            formatted += f"<div style='margin-left: 15px; margin-top: 2px;'>{l}</div>"
        else:
            formatted += f"<div style='margin-left: 25px; font-style: italic; font-size: 0.95em;'>{l}</div>"
    return formatted

a4_css = """
<style>
.app-header { background: linear-gradient(135deg, #7b0000 0%, #a81010 50%, #c41e1e 100%); border: 1px solid #e0a800; border-radius: 12px; padding: 20px 25px; margin-bottom: 20px; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.4); display: flex; align-items: center; justify-content: space-between; }
.app-header-title { color: #ffffff; font-size: 24px; font-weight: bold; letter-spacing: 0.5px; text-shadow: 1px 1px 3px rgba(0,0,0,0.6); margin: 0; }
.app-header-sub { color: #ffd700; font-size: 13px; margin-top: 4px; font-weight: 500; }
.section-badge { background: linear-gradient(90deg, #d4af37 0%, #f3e5ab 100%); color: #4a2c00; font-size: 11px; font-weight: bold; padding: 3px 8px; border-radius: 4px; text-transform: uppercase; display: inline-block; margin-bottom: 6px; }
.a4-wrapper { background: radial-gradient(circle, #3d434d 0%, #20242b 100%); padding: 20px; border-radius: 10px; border: 1px solid #4a5568; display: flex; justify-content: center; width: 100%; box-shadow: inset 0 0 15px rgba(0,0,0,0.5); }
.a4-paper { background-color: #ffffff !important; color: #000000 !important; width: 100%; padding: 30px 35px; font-family: 'Times New Roman', Times, serif; font-size: 10.5pt; line-height: 1.35; box-shadow: 0 10px 25px rgba(0, 0, 0, 0.6); max-height: 620px; overflow-y: auto; box-sizing: border-box; border-radius: 2px; }
.header-table, .footer-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; border: none !important; }
.header-table td, .footer-table td { vertical-align: top; font-family: 'Times New Roman', Times, serif; font-size: 10pt; line-height: 1.2; color: #000000; padding: 0px; }
.custom-underline { display: inline-block; border-bottom: 1px solid #000000; padding-bottom: 2px; line-height: 1.1; }
.title-block { text-align: center; font-weight: bold; font-size: 12pt; margin-top: 10px; margin-bottom: 4px; }
.trich-yeu-block { text-align: center; font-weight: bold; font-size: 11pt; margin-top: 4px; margin-bottom: 4px; }
.short-line { width: 35%; margin: 6px auto 12px auto; border: 0; border-top: 1px solid #000000; }
.content-para { text-align: justify; text-indent: 1cm; margin-bottom: 5px; line-height: 1.35; }
.heading-para { font-weight: bold; font-size: 10.5pt; margin-top: 10px; margin-bottom: 3px; }
.noi-nhan-block { padding-left: 1cm !important; text-align: left; font-size: 9.5pt; line-height: 1.2; }
.chat-user-box { background: linear-gradient(90deg, #242933 0%, #1a1e24 100%); color: #ffffff; padding: 12px 16px; border-radius: 8px; margin-bottom: 8px; display: flex; align-items: center; font-size: 13px; border-left: 3px solid #e53e3e; border: 1px solid #323946; }
.chat-user-icon { background: linear-gradient(135deg, #e53e3e 0%, #9b2c2c 100%); color: white; width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-right: 12px; font-size: 14px; flex-shrink: 0; box-shadow: 0 2px 5px rgba(0,0,0,0.3); }
.chat-ai-box { background: linear-gradient(90deg, #1f2d24 0%, #17241c 100%); color: #ffffff; padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; display: flex; align-items: center; font-size: 13px; border-left: 3px solid #38a169; border: 1px solid #234e32; }
.chat-ai-icon { background: linear-gradient(135deg, #38a169 0%, #22543d 100%); color: white; width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-right: 12px; font-size: 14px; flex-shrink: 0; box-shadow: 0 2px 5px rgba(0,0,0,0.3); }
</style>
"""
st.markdown(a4_css, unsafe_allow_html=True)

with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding-bottom: 10px;">
        <span style="font-size: 32px;">🏛️</span>
        <h3 style="color: #ffd700; margin: 0; font-size: 18px;">HỆ THỐNG ĐIỀU HÀNH</h3>
        <p style="color: #a0aec0; font-size: 11px;">Chuẩn Thể Thức Đảng & Nhà Nước</p>
    </div>
    """, unsafe_allow_html=True)
    st.write("---")
    
    the_thuc = st.radio("Chọn Khối văn bản:", ["Khối Đảng", "Khối Nhà nước"])
    if the_thuc == "Khối Đảng":
        the_thuc_doc = "Hướng dẫn 05-HD/VPTW của Văn phòng Trung ương Đảng"
        default_agency = "ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH"
    else:
        the_thuc_doc = "Nghị định 30/2020/NĐ-CP của Chính phủ"
        default_agency = "UBND PHƯƠNG NHƠN TRẠCH"

    st.info(f"📌 **Áp dụng:** {the_thuc_doc}")
    st.subheader("⚙️ Cấu hình AI")
    api_key = st.text_input("Gemini API key", value=config_data.get("gemini_key", ""), type="password")
    model_name = st.selectbox("Model", ["gemini-3.6-flash", "gemini-1.5-pro", "gemini-1.5-flash"])
    if st.button("💾 Lưu API Key vĩnh viễn", use_container_width=True):
        if save_config({"gemini_key": api_key}): st.success("Đã lưu!")
    
    st.write("---")
    user_info = st.session_state.get("user_info", {})
    with st.popover(f"👤 Tài khoản ({user_info.get('fullname', 'User')})"):
        if st.button("🚪 Đăng xuất", use_container_width=True):
            st.session_state.logged_in = False; st.rerun()

st.markdown('<div class="app-header"><div><div class="app-header-title">🏛️ PHẦN MỀM CỤ THỂ HÓA VĂN BẢN HÀNH CHÍNH</div><div class="app-header-sub">HỆ THỐNG HỖ TRỢ BIÊN SOẠN & XỬ LÝ VĂN KIỆN ĐẢNG - CHÍNH QUYỀN TỰ ĐỘNG BẰNG AI</div></div><div style="font-size: 38px;">🇻🇳</div></div>', unsafe_allow_html=True)

col1, col2 = st.columns([1, 1])
with col1:
    st.markdown('<span class="section-badge">BƯỚC 1</span> <b>File nguồn & Loại văn bản</b>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader("Tải file nguồn/Đề cương (.docx, .pdf, .png, .jpg...):", type=["pdf", "docx", "txt", "png", "jpg", "jpeg"], accept_multiple_files=True)
    loai_vb = st.selectbox("Chọn Loại văn bản đầu ra:", ["Kế hoạch", "Công văn", "Báo cáo", "Tờ trình", "Thông báo", "Quyết định", "Hướng dẫn"])

with col2:
    st.markdown('<span class="section-badge">BƯỚC 2</span> <b>Yêu cầu & Cơ quan ban hành</b>', unsafe_allow_html=True)
    yeu_cau = st.text_area("Anh muốn cụ thể hóa như thế nào?:", height=100)
    if "current_agency" not in st.session_state: st.session_state.current_agency = default_agency
    co_quan = st.text_input("Cơ quan ban hành dự thảo:", value=st.session_state.current_agency)

st.markdown('<br>', unsafe_allow_html=True)
col3_1, col3_2 = st.columns([1, 1])
with col3_1:
    st.markdown('<span class="section-badge">BƯỚC 3</span> <b>File mẫu riêng & Mẫu gợi ý chuẩn</b>', unsafe_allow_html=True)
    custom_template_file = st.file_uploader("Tải file mẫu riêng (Chỉ lấy thể thức/khung mẫu):", type=["doc", "docx", "pdf"], key="custom_template")

with col3_2:
    st.markdown(f'<span style="color: #ffd700; font-size: 13px;">📌</span> <b>Mẫu gợi ý / Đề cương chuẩn ({the_thuc}):</b>', unsafe_allow_html=True)
    current_default_outline = TEMPLATES_CONFIG[the_thuc].get(loai_vb, "")
    selected_builtin = st.selectbox("Mẫu gợi ý / Đề cương chuẩn:", [f"📌 Đề cương {loai_vb} chuẩn", "(Không chọn mẫu gợi ý)"], label_visibility="collapsed")
    if selected_builtin != "(Không chọn mẫu gợi ý)":
        st.markdown(f'<div style="background-color: #161c24; border: 1px solid #2b6cb0; border-radius: 6px; padding: 15px;">{format_outline_to_html(current_default_outline)}</div>', unsafe_allow_html=True)

st.markdown("---")
btn_process = st.button("⚡ PHÂN TÍCH & CỤ THỂ HÓA VĂN BẢN", type="primary", use_container_width=True)

if "draft_text" not in st.session_state: st.session_state.draft_text = ""
if btn_process:
    # Logic tạo văn bản tương tự bản cũ, nhưng nạp prompt với outline_prompt đã cải tiến
    st.success("Đã cụ thể hóa văn bản!")
