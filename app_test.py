import streamlit as st
import google.generativeai as genai
import pypdf
import docx
import io
import re

st.set_page_config(page_title="Phần mềm Cụ thể hóa Văn bản Hành chính", page_icon="📝", layout="wide")

# --- CỘT BÊN TRÁI: CẤU HÌNH ---
with st.sidebar:
    st.title("⚙️ Cấu hình Thể thức & AI")
    the_thuc = st.radio("Chọn Khối văn bản:", ["Khối Đảng", "Khối Nhà nước"])
    api_key = st.text_input("Gemini API key", type="password")
    model_name = st.selectbox("Model", ["gemini-3.6-flash"])

st.title("📝 Phần mềm Cụ thể hóa Văn bản Hành chính")

# --- KHU VỰC NHẬP DỮ LIỆU ---
col_left, col_right = st.columns([1, 1])
with col_left:
    uploaded_files = st.file_uploader(
        "Tải lên một hoặc nhiều file nguồn (.pdf, .docx, .png, .jpg...):", 
        type=["pdf", "docx", "txt", "png", "jpg", "jpeg"], 
        accept_multiple_files=True
    )
    loai_vb = st.selectbox("Chọn Loại văn bản đầu ra:", ["Công văn", "Kế hoạch", "Báo cáo", "Tờ trình", "Thông báo", "Quyết định"])

with col_right:
    co_quan = st.text_input("Cơ quan ban hành dự thảo:", value="ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH" if the_thuc == "Khối Đảng" else "UBND PHƯƠNG NHƠN TRẠCH")
    yeu_cau = st.text_area("Anh/Chị muốn cụ thể hóa hoặc tổng hợp như thế nào?:", height=120, placeholder="Ví dụ: Cụ thể hóa kế hoạch trên thành kế hoạch của phường...")

st.markdown("---")
btn_process = st.button("⚡ PHÂN TÍCH TỔNG HỢP & CỤ THỂ HÓA VĂN BẢN", type="primary", use_container_width=True)

# Khởi tạo lưu trữ nội dung dự thảo
if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""

# --- XỬ LÝ TẠO VĂN BẢN BAN ĐẦU ---
if btn_process:
    if not api_key:
        st.error("Vui lòng nhập Gemini API Key ở cột bên trái!")
    elif not uploaded_files and not yeu_cau:
        st.warning("Vui lòng tải file hoặc nhập yêu cầu!")
    else:
        with st.spinner("Đang xử lý dữ liệu và tạo văn bản..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name)
                
                content_parts = []
                extracted_texts = []
                
                for uf in uploaded_files:
                    bytes_data = uf.read()
                    if uf.name.lower().endswith('.pdf'):
                        try:
                            reader = pypdf.PdfReader(io.BytesIO(bytes_data))
                            pdf_text = "".join([page.extract_text() or "" for page in reader.pages])
                            if len(pdf_text.strip()) > 50:
                                extracted_texts.append(f"--- NỘI DUNG FILE {uf.name} ---\n" + pdf_text)
                            else:
                                content_parts.append({"mime_type": "application/pdf", "data": bytes_data})
                        except:
                            content_parts.append({"mime_type": "application/pdf", "data": bytes_data})
                    elif uf.name.lower().endswith('.docx'):
                        doc_file = docx.Document(io.BytesIO(bytes_data))
                        docx_text = "\n".join([p.text for p in doc_file.paragraphs])
                        extracted_texts.append(f"--- NỘI DUNG FILE {uf.name} ---\n" + docx_text)
                    else:
                        content_parts.append({"mime_type": uf.type, "data": bytes_data})
                
                prompt = f"""
                Bạn là chuyên gia soạn thảo văn bản hành chính công tác tại Việt Nam.
                Hãy căn cứ vào toàn bộ tài liệu đính kèm để soạn thảo 01 dự thảo văn bản hoàn chỉnh.
                
                THỂ THỨC: {the_thuc} | CƠ QUAN: {co_quan} | LOẠI VĂN BẢN: {loai_vb}
                YÊU CẦU CỤ THỂ HÓA: {yeu_cau}
                
                DỮ LIỆU TÀI LIỆU NGUỒN:
                {"".join(extracted_texts)}
                
                YÊU CẦU TRÌNH BÀY DẠNG VĂN BẢN HÀNH CHÍNH CHUẨN:
                1. Quốc hiệu, Tiêu ngữ, Tên cơ quan, Số/Ký hiệu, Địa danh ngày tháng trình bày đúng vị trí dòng.
                2. Tên loại văn bản và Trích yếu in đậm, căn giữa.
                3. Các mục I, II, III và các tiểu mục rõ ràng, đúng thể thức hành chính Việt Nam.
                """
                content_parts.insert(0, prompt)
                
                response = model.generate_content(content_parts)
                st.session_state.draft_text = response.text
                st.success("Đã cụ thể hóa văn bản thành công!")
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")

# --- HÌNH ẢNH HIỂN THỊ TRANG WORD A4 MÔ PHỎNG ---
if st.session_state.draft_text:
    res_col1, res_col2 = st.columns([1.2, 0.8])
    
    with res_col1:
        st.subheader("📄 Bản dự thảo trang Word (A4)")
        
        # Định dạng văn bản hiển thị trong khung A4 trắng
        formatted_html = st.session_state.draft_text
        formatted_html = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', formatted_html)
        formatted_html = re.sub(r'### (.*?)\n', r'<h3>\1</h3>', formatted_html)
        formatted_html = re.sub(r'#### (.*?)\n', r'<h4>\1</h4>', formatted_html)
        formatted_html = formatted_html.replace('\n', '<br>')
        
        # CSS tạo tờ giấy A4 màu trắng có bóng đổ
        a4_css = """
        <style>
        .a4-paper {
            background-color: #ffffff !important;
            color: #000000 !important;
            padding: 45px 50px;
            font-family: 'Times New Roman', Times, serif;
            font-size: 13pt;
            line-height: 1.5;
            border: 1px solid #cccccc;
            box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.25);
            border-radius: 3px;
            max-height: 550px;
            overflow-y: auto;
            word-wrap: break-word;
        }
        .a4-paper h3, .a4-paper h4 {
            color: #000000 !important;
            font-family: 'Times New Roman', Times, serif;
            margin-top: 10px;
            margin-bottom: 5px;
        }
        </style>
        """
        
        st.markdown(a4_css + f'<div class="a4-paper">{formatted_html}</div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Hàm xuất file Word (.docx)
        def generate_docx(text):
            doc = docx.Document()
            for section in doc.sections:
                section.top_margin = docx.shared.Inches(0.79)
                section.bottom_margin = docx.shared.Inches(0.79)
                section.left_margin = docx.shared.Inches(1.18)
                section.right_margin = docx.shared.Inches(0.79)
            p = doc.add_paragraph(text)
            p.style.font.name = 'Times New Roman'
            p.style.font.size = docx.shared.Pt(13)
            bio = io.BytesIO()
            doc.save(bio)
            return bio.getvalue()

        st.download_button(
            label="📥 TẢI VỀ FILE WORD (.DOCX)",
            data=generate_docx(st.session_state.draft_text),
            file_name="Du_Thao_Van_Ban.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
        
    with res_col2:
        st.subheader("💬 Chat AI sửa đổi (Google Gemini)")
        edit_instruction = st.text_area(
            "Nhập yêu cầu chỉnh sửa văn bản...", 
            height=180,
            placeholder="VD: 'Bổ sung thêm mốc thời gian hoàn thành 30/9 vào Mục II', 'Sửa lại căn cứ số 1'..."
        )
        if st.button("Chỉnh sửa dự thảo", use_container_width=True):
            if not edit_instruction:
                st.warning("Vui lòng nhập nội dung cần sửa!")
            elif not api_key:
                st.error("Vui lòng nhập Gemini API Key!")
            else:
                with st.spinner("AI đang chỉnh sửa bản dự thảo..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel(model_name)
                        edit_prompt = f"""
                        BẢN DỰ THẢO HIỆN TẠI:
                        {st.session_state.draft_text}

                        YÊU CẦU CHỈNH SỬA TỪ NGƯỜI DÙNG:
                        {edit_instruction}

                        Hãy cập nhật và hoàn thiện lại toàn bộ bản dự thảo văn bản trên theo đúng yêu cầu chỉnh sửa. Giữ nguyên thể thức văn bản hành chính.
                        """
                        res_edit = model.generate_content(edit_prompt)
                        st.session_state.draft_text = res_edit.text
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi chỉnh sửa: {str(e)}")
