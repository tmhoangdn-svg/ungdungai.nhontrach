import streamlit as st
import google.generativeai as genai
import pypdf, docx, io, re, os, json

# Cấu hình lưu API Key
CONFIG_FILE = "config_keys.json"
def save_config(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(data, f)
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {}
config_data = load_config()

st.set_page_config(page_title="Phần mềm Cụ thể hóa Văn bản", layout="wide")

# GIAO DIỆN
st.markdown("""<style>
.app-header { background: #7b0000; color: white; padding: 20px; border-radius: 10px; text-align: center; }
.a4-paper { background: white; color: black; padding: 40px; font-family: 'Times New Roman'; min-height: 700px; border: 1px solid #ccc; }
</style>""", unsafe_allow_html=True)

st.markdown('<div class="app-header"><h1>🏛️ HỆ THỐNG ĐIỀU HÀNH VĂN BẢN TỰ ĐỘNG</h1></div>', unsafe_allow_html=True)

with st.sidebar:
    st.subheader("🔑 Cấu hình hệ thống")
    api_key = st.text_input("Nhập Gemini API Key của bạn:", value=config_data.get("gemini_key", ""), type="password")
    if st.button("Lưu Key"):
        save_config({"gemini_key": api_key})
        st.success("Đã lưu!")

# HÀM AI THÔNG MINH (TỰ CHỌN MODEL)
def get_ai_response(prompt, files=None):
    if not api_key: raise Exception("Vui lòng nhập API Key!")
    genai.configure(api_key=api_key)
    # Tự động thử nghiệm model Pro trước, nếu lỗi thì tự chuyển sang Flash
    try:
        model = genai.GenerativeModel("gemini-1.5-pro")
        return model.generate_content(prompt).text
    except:
        model = genai.GenerativeModel("gemini-1.5-flash")
        return model.generate_content(prompt).text

col1, col2 = st.columns(2)
with col1:
    uploaded = st.file_uploader("Tải file nguồn:", accept_multiple_files=True)
    loai_vb = st.selectbox("Loại văn bản:", ["Kế hoạch", "Công văn", "Báo cáo", "Tờ trình", "Thông báo", "Quyết định", "Hướng dẫn"])
with col2:
    yeu_cau = st.text_area("Yêu cầu:", height=100)
    co_quan = st.text_input("Cơ quan ban hành:", value="ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH")

if st.button("⚡ PHÂN TÍCH & DỰ THẢO VĂN BẢN", type="primary", use_container_width=True):
    with st.spinner("AI đang xử lý..."):
        try:
            # Thu thập nội dung file
            content_text = ""
            for f in uploaded:
                if f.name.endswith('.pdf'):
                    reader = pypdf.PdfReader(f)
                    content_text += "".join([p.extract_text() for p in reader.pages])
            
            prompt = f"Soạn {loai_vb} cho {co_quan}. Nội dung nguồn: {content_text}. Yêu cầu: {yeu_cau}. Không dùng Markdown."
            st.session_state.draft = get_ai_response(prompt)
            st.rerun()
        except Exception as e:
            st.error(f"Lỗi: {e}. Vui lòng kiểm tra lại API Key!")

if "draft" in st.session_state:
    st.markdown("### 📄 Bản dự thảo")
    st.markdown(f'<div class="a4-paper">{st.session_state.draft.replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)
    
    st.markdown("### 💬 Trợ lý chỉnh sửa")
    edit = st.text_input("Yêu cầu chỉnh sửa:")
    if st.button("Chỉnh sửa"):
        with st.spinner("Đang chỉnh sửa..."):
            st.session_state.draft = get_ai_response(f"Văn bản cũ: {st.session_state.draft}. Yêu cầu sửa: {edit}. Trả về văn bản mới.")
            st.rerun()
