st.session_state.reg_missing = []

if not st.session_state.logged_in:
        st.title("🔐 ĐĂNG NHẬP HỆ THỐNG")
        tab_login, tab_register, tab_forgot = st.tabs(["Đăng nhập", "Đăng ký tài khoản", "Quên mật khẩu"])

        # 1. TAB ĐĂNG NHẬP
        with tab_login:
            with st.form("login_form"):
                username = st.text_input("Tên đăng nhập")
                password = st.text_input("Mật khẩu", type="password")
                btn_login = st.form_submit_button("Đăng nhập")

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
        # Gom form đăng nhập vào chính giữa màn hình
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

            # 1. TAB ĐĂNG NHẬP
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
                    except Exception as e:
                        st.error("Chưa thể kết nối đến dữ liệu tài khoản.")

        # 2. TAB ĐĂNG KÝ
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
                btn_register = st.form_submit_button("Đăng ký")

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
                    except:
                        st.error("Không thể kết nối máy chủ đăng ký.")

        # 3. TAB QUÊN MẬT KHẨU
        with tab_forgot:
            with st.form("forgot_form"):
                fg_user = st.text_input("Nhập Tên đăng nhập của bạn", key="fg_u")
                fg_contact = st.text_input("Nhập Email hoặc Số điện thoại đã đăng ký", key="fg_c")
                fg_new_pass = st.text_input("Mật khẩu mới", type="password", key="fg_p1")
                fg_confirm_pass = st.text_input("Xác nhận mật khẩu mới", type="password", key="fg_p2")
                btn_forgot = st.form_submit_button("Đặt lại mật khẩu")

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
                        except Exception as e:
                            st.error("Chưa thể kết nối đến dữ liệu tài khoản.")

            # 2. TAB ĐĂNG KÝ
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
                                st.success("Đổi mật khẩu thành công! Vui lòng quay lại tab Đăng nhập.")
                                st.success("Đăng ký thành công! Vui lòng quay lại tab Đăng nhập.")
else:
                                st.error("Lỗi khi cập nhật mật khẩu.")
                        else:
                            st.error("Tên đăng nhập và Email/SĐT không khớp!")
                    except Exception as e:
                        st.error("Không thể xác minh thông tin.")
                                st.error("Lỗi khi tạo tài khoản.")
                        except:
                            st.error("Không thể kết nối máy chủ đăng ký.")

            # 3. TAB QUÊN MẬT KHẨU
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
                        except Exception as e:
                            st.error("Không thể xác minh thông tin.")

return False
return True
@@ -168,7 +178,7 @@ def check_login():
st.stop()

# ==============================================================================
# 2. CẤU HÌNH GIAO DIỆN CSS CAO CẤP (HOA VĂN & MÀU SẮC CHUYÊN NGHIỆP)
# 2. CẤU HÌNH GIAO DIỆN & BẢNG ÁNH XẠ
# ==============================================================================
SYMBOL_MAP = {
"Kế hoạch": "KH",
@@ -179,9 +189,8 @@ def check_login():
"Quyết định": "QĐ"
}

custom_theme_css = """
a4_css = """
<style>
/* 1. Tiêu đề Banner dạng Cổng thông tin Quốc gia */
.app-header {
   background: linear-gradient(135deg, #7b0000 0%, #a81010 50%, #c41e1e 100%);
   border: 1px solid #e0a800;
@@ -207,17 +216,6 @@ def check_login():
   margin-top: 4px;
   font-weight: 500;
}

/* 2. Thẻ Khung Nhập Liệu (Card Box) Viền Ánh Kim */
.custom-card {
    background-color: #1a1e24;
    border: 1px solid #2d3748;
    border-top: 3px solid #d4af37;
    border-radius: 10px;
    padding: 15px;
    margin-bottom: 15px;
    box-shadow: 0 3px 8px rgba(0,0,0,0.3);
}
.section-badge {
   background: linear-gradient(90deg, #d4af37 0%, #f3e5ab 100%);
   color: #4a2c00;
@@ -229,8 +227,6 @@ def check_login():
   display: inline-block;
   margin-bottom: 6px;
}

/* 3. Tờ giấy A4 xem trước nổi bật trên nền bàn làm việc */
.a4-wrapper {
   background: radial-gradient(circle, #3d434d 0%, #20242b 100%);
   padding: 20px;
@@ -313,8 +309,6 @@ def check_login():
   font-size: 9.5pt;
   line-height: 1.2;
}

/* 4. Khung Chat AI phong cách Hiện đại */
.chat-user-box {
   background: linear-gradient(90deg, #242933 0%, #1a1e24 100%);
   color: #ffffff;
@@ -369,7 +363,7 @@ def check_login():
}
</style>
"""
st.markdown(custom_theme_css, unsafe_allow_html=True)
st.markdown(a4_css, unsafe_allow_html=True)

# --- THÔNG TIN TÀI KHOẢN TRÊN SIDEBAR ---
with st.sidebar:
@@ -406,7 +400,7 @@ def check_login():
st.rerun()

# ==============================================================================
# 3. GIAO DIỆN CHÍNH (HEADER BANNER & KHUNG NHẬP LIỆU)
# 3. GIAO DIỆN CHÍNH
# ==============================================================================
st.markdown("""
<div class="app-header">
