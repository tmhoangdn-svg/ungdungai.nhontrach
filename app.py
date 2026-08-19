import streamlit as st
import pandas as pd
import google.generativeai as genai
import pypdf
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io, re, os, json

# ==============================================================================
# CẤU HÌNH & HÀM HỖ TRỢ
# ==============================================================================
CONFIG_FILE = "config_keys.json"

def save_config(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(data, f)
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {}

config_data = load_config()

st.set_page_config(page_title="Phần mềm Cụ thể hóa Văn bản Hành chính", page_icon="🏛️", layout="wide")

# CSS Giao diện
st.markdown("""
<style>
.app-header { background: #7b0000; color: white; padding: 20px; border-radius: 10px; margin-bottom: 20px; text-align: center; }
.a4-wrapper { background: #2d3748; padding: 20px; border-radius: 10px; }
.a4-paper { background: white; color: black; padding: 40px; font-family: 'Times New Roman'; font-size: 11pt; line-height: 1.5; min-height: 800px; }
</style>
""", unsafe_allow_html=True)

# GIAO DIỆN CHÍNH
st.markdown('<div class="app-header"><h1>🏛️ PHẦN MỀM CỤ THỂ HÓA VĂN BẢN HÀNH CHÍNH</h1></div>', unsafe_allow_html=True)

with st.sidebar:
    st.subheader("⚙️ Cấu hình Gemini Pro")
    api_key = st.text_input("Gemini API Key", value=config_data.get("gemini_key", ""), type="password")
    model_name = st.selectbox("Model", ["gemini-1.5-pro", "gemini-1.5-flash"])
    if st.button("💾 Lưu API Key"):
        save_config({"gemini_key": api_key})
        st.success("Đã lưu!")

col1, col2 = st.columns(2)
with col1:
    uploaded_files = st.file_uploader("Tải file nguồn (PDF, DOCX):", accept_multiple_files=True)
    loai_vb = st.selectbox("Loại văn bản:", ["Kế hoạch", "Công văn", "Báo cáo", "Tờ trình", "Thông báo", "Quyết định", "Hướng dẫn"])
with col2:
    yeu_cau = st.text_area("Yêu cầu cụ thể:", height=100)
    co_quan = st.text_input("Cơ quan ban hành:", value="ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH")

st.markdown("---")
if st.button("⚡ PHÂN TÍCH & DỰ THẢO VĂN BẢN", type="primary", use_container_width=True):
    if not api_key: st.error("Chưa nhập API Key!"); st.stop()
    
    with st.spinner("Đang soạn thảo văn bản..."):
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)
        
        # Xử lý văn bản
        prompt = f"Soạn thảo văn bản {loai_vb} cho {co_quan}. Yêu cầu: {yeu_cau}. Không dùng Markdown."
        response = model.generate_content(prompt)
        st.session_state.draft_text = response.text
        st.rerun()

# HIỂN THỊ KẾT QUẢ
if "draft_text" in st.session_state and st.session_state.draft_text:
    res_col1, res_col2 = st.columns([1, 1])
    with res_col1:
        st.markdown("### 📄 BẢN DỰ THẢO")
        st.markdown(f'<div class="a4-wrapper"><div class="a4-paper">{st.session_state.draft_text.replace(chr(10), "<br>")}</div></div>', unsafe_allow_html=True)
    
    with res_col2:
        st.markdown("### 💬 TRỢ LÝ AI CHỈNH SỬA")
        edit_req = st.text_area("Yêu cầu chỉnh sửa:")
        if st.button("Thực hiện chỉnh sửa"):
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            res = model.generate_content(f"Văn bản cũ: {st.session_state.draft_text}. Yêu cầu sửa: {edit_req}. Trả về văn bản mới.")
            st.session_state.draft_text = res.text
            st.rerun()
