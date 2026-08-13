import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import json
import os

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Phần mềm Cụ thể hóa Văn bản Hành chính",
    page_icon="📝",
    layout="wide"
)

# --- CỘT BÊN TRÁI: CẤU HÌNH HỆ THỐNG ---
with st.sidebar:
    st.title("⚙️ Cấu hình Thể thức & AI")
    
    st.subheader("1. Thể thức Văn bản")
    the_thuc = st.radio(
        "Chọn Khối văn bản:",
        ["Khối Đảng", "Khối Nhà nước"],
        help="Khối Đảng: Hướng dẫn 05-HD/VPTW | Khối Nhà nước: Nghị định 30/2020/NĐ-CP"
    )
    
    st.subheader("2. Cấu hình AI")
    ai_provider = st.selectbox("AI xử lý chính", ["Google Gemini"])
    api_key = st.text_input("Gemini API key", type="password", help="Dán mã API lấy từ aistudio.google.com")
    model_name = st.selectbox("Model", ["gemini-3.6-flash"])

# --- GIỮA: GIAO DIỆN CHÍNH ---
st.title("📝 Phần mềm Cụ thể hóa Văn bản Hành chính")

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("1. File nguồn & Phụ lục (Đa phương thức)")
    uploaded_files = st.file_uploader(
        "Tải lên một hoặc nhiều file (.pdf, .docx, .png, .jpg...):",
        type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        help="Hệ thống hỗ trợ đọc trực tiếp nhiều file PDF ký số, file scan, file Word cùng lúc như NotebookLM."
    )
    
    loai_vb = st.selectbox(
        "Chọn Loại văn bản đầu ra:",
        ["Công văn", "Kế hoạch", "Báo cáo", "Tờ trình", "Thông báo", "Quyết định"]
    )

with col_right:
    st.subheader("2. Yêu cầu & Cơ quan ban hành")
    co_quan = st.text_input(
        "Cơ quan ban hành dự thảo:",
        value="ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH" if the_thuc == "Khối Đảng" else "UBND PHƯƠNG NHƠN TRẠCH"
    )
    yeu_cau = st.text_area(
        "Anh/Chị muốn cụ thể hóa hoặc tổng hợp như thế nào?:",
        height=120,
        placeholder="Ví dụ: Căn cứ Kế hoạch 156 và Phụ lục kèm theo, hãy soạn Kế hoạch tuyên truyền cấp phường tập trung vào lộ trình Giai đoạn 1..."
    )

st.markdown("---")
btn_process = st.button("⚡ PHÂN TÍCH TỔNG HỢP & CỤ THỂ HÓA VĂN BẢN", type="primary", use_container_width=True)

# Session State lưu trữ văn bản dự thảo
if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""

# --- XỬ LÝ KHI BẤM NÚT ---
if btn_process:
    if not api_key:
        st.error("Vui lòng nhập Gemini API Key ở cột bên trái!")
    elif not uploaded_files and not yeu_cau:
        st.warning("Vui lòng tải lên ít nhất 1 file nguồn hoặc nhập yêu cầu cụ thể hóa!")
    else:
        with st.spinner("Đang tải dữ liệu đa file và tổng hợp văn bản chuẩn thể thức..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name)
                
                # Chuẩn bị Prompt chỉ đạo
                prompt = f"""
                Bạn là chuyên gia soạn thảo văn bản hành chính công tác tại Việt Nam.
                Hãy căn cứ vào toàn bộ các file tài liệu nguồn được đính kèm và yêu cầu dưới đây để soạn thảo 01 dự thảo văn bản hoàn chỉnh.
                
                QUY CHUẨN THỂ THỨC:
                - Thể thức chọn: {the_thuc} ({'Hướng dẫn 05-HD/VPTW của Văn phòng TW Đảng' if the_thuc == 'Khối Đảng' else 'Nghị định 30/2020/NĐ-CP của Chính phủ'})
                - Cơ quan ban hành: {co_quan}
                - Loại văn bản cần lập: {loai_vb}
                - Yêu cầu xử lý cụ thể: {yeu_cau}
                
                YÊU CẦU NỘI DUNG & TRÌNH BÀY:
                1. Trích xuất chính xác 100% các căn cứ pháp lý, số hiệu, ngày ban hành từ các file nguồn.
                2. Bố cục đầy đủ: Quốc hiệu, Tiêu ngữ, Tên cơ quan, Số/Ký hiệu, Địa danh ngày tháng, Tên loại & Trích yếu, Nội dung (các mục I, II, III...), Nơi nhận, Chức vụ người ký.
                3. Lời văn chuẩn mực hành chính, rõ ràng, thiết thực.
                """
                
                # Đóng gói danh sách nội dung gửi Gemini (gồm Prompt + tất cả File dưới dạng Bytes/Parts)
                content_parts = [prompt]
                for uf in uploaded_files:
                    bytes_data = uf.read()
                    content_parts.append({
                        "mime_type": uf.type,
                        "data": bytes_data
                    })
                
                # Gọi API Gemini Multimodal
                response = model.generate_content(content_parts)
                st.session_state.draft_text = response.text
                st.success("Đã cụ thể hóa văn bản thành công!")
            except Exception as e:
                st.error(f"Lỗi xử lý: {str(e)}")

# --- HÌNH ẢNH HIỂN THỊ KẾT QUẢ & XUẤT FILE ---
if st.session_state.draft_text:
    res_col1, res_col2 = st.columns([1, 1])
    
    with res_col1:
        st.subheader("📄 Bản dự thảo trang Word (A4)")
        st.text_area("Nội dung dự thảo:", value=st.session_state.draft_text, height=450)
        
        # Hàm xuất file Word (.docx) chuẩn thể thức
        def generate_docx(text):
            doc = docx.Document()
            # Căn lề chuẩn A4
            for section in doc.sections:
                section.top_margin = Inches(0.79)
                section.bottom_margin = Inches(0.79)
                section.left_margin = Inches(1.18)
                section.right_margin = Inches(0.79)
            
            p = doc.add_paragraph(text)
            p.style.font.name = 'Times New Roman'
            p.style.font.size = Pt(13)
            
            bio = io.BytesIO()
            doc.save(bio)
            return bio.getvalue()

        docx_bytes = generate_docx(st.session_state.draft_text)
        st.download_button(
            label="📥 TẢI VỀ FILE WORD (.DOCX)",
            data=docx_bytes,
            file_name="Du_Thao_Van_Ban_Hanh_Chinh.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
        
    with res_col2:
        st.subheader("💬 Chat AI sửa đổi (Google Gemini)")
        edit_instruction = st.text_area("Nhập yêu cầu chỉnh sửa văn bản...", placeholder="VD: 'Bổ sung thêm mốc thời gian hoàn thành 30/9 vào Mục II', 'Sửa lại căn cứ số 1'...")
        if st.button("Chỉnh sửa dự thảo"):
            if edit_instruction and api_key:
                with st.spinner("Đang cập nhật bản dự thảo..."):
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(model_name)
                    edit_prompt = f"Bản dự thảo hiện tại:\n{st.session_state.draft_text}\n\nYêu cầu sửa đổi:\n{edit_instruction}\n\nHãy cập nhật lại toàn bộ bản dự thảo văn bản."
                    res_edit = model.generate_content(edit_prompt)
                    st.session_state.draft_text = res_edit.text
                    st.experimental_rerun()
