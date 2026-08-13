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

if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""

# --- XỬ LÝ TẠO VĂN BẢN ---
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
                Bạn là chuyên gia soạn thảo văn bản hành chính Việt Nam.
                Hãy soạn thảo 01 dự thảo văn bản hoàn chỉnh căn cứ vào nội dung được cấp.
                
                THỂ THỨC: {the_thuc} | CƠ QUAN BAN HÀNH: {co_quan} | LOẠI VĂN BẢN: {loai_vb}
                YÊU CẦU CỤ THỂ HÓA: {yeu_cau}
                
                DỮ LIỆU TÀI LIỆU NGUỒN:
                {"".join(extracted_texts)}
                
                QUY CẮC ĐỊNH DẠNG VĂN BẢN (RẤT QUAN TRỌNG):
                1. TUYỆT ĐỐI KHÔNG DÙNG ký tự Markdown (*, #, _).
                2. Bố cục đầy đủ từ Quốc hiệu, Tiêu ngữ, Tên cơ quan, Số/Ký hiệu, Trích yếu, Nội dung (I, II, III...), Nơi nhận, Người ký.
                3. Lời văn chuẩn mực, đúng thể thức hành chính công tác Đảng/Nhà nước.
                """
                content_parts.insert(0, prompt)
                
                response = model.generate_content(content_parts)
                st.session_state.draft_text = response.text
                st.success("Đã cụ thể hóa văn bản thành công!")
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")

# --- HIỂN THỊ TRANG XEM TRƯỚC A4 NÂNG CẤP ---
if st.session_state.draft_text:
    res_col1, res_col2 = st.columns([1.2, 0.8])
    
    with res_col1:
        st.subheader("📄 Bản dự thảo trang Word (A4)")
        
        # 1. Làm sạch ký tự Markdown rác
        clean_text = re.sub(r'[\*#_]', '', st.session_state.draft_text)
        
        # 2. Phân tích dòng và tạo HTML chuẩn lề hành chính
        lines = [l.strip() for l in clean_text.split('\n') if l.strip()]
        formatted_html = ""
        
        for line in lines:
            # Tiêu đề các mục lớn I, II, III...
            if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', line):
                formatted_html += f'<p style="font-weight: bold; font-size: 13pt; margin-top: 14px; margin-bottom: 4px; text-align: left;">{line}</p>'
            # Các tiểu mục 1, 2, 3...
            elif re.match(r'^\d+\.', line) and len(line) < 80:
                formatted_html += f'<p style="font-weight: bold; font-size: 13pt; margin-top: 8px; margin-bottom: 3px; text-align: left;">{line}</p>'
            # Tên loại văn bản và trích yếu in đậm căn giữa
            elif line.isupper() and len(line) < 100:
                formatted_html += f'<p style="text-align: center; font-weight: bold; font-size: 14pt; margin-top: 12px; margin-bottom: 6px;">{line}</p>'
            # Dòng Kính gửi hoặc gạch đầu dòng
            elif line.startswith("Kính gửi:") or line.startswith("-"):
                formatted_html += f'<p style="text-align: left; margin-bottom: 4px; padding-left: 15px;">{line}</p>'
            # Đoạn văn bản thông thường: Căn đều 2 bên, lùi đầu dòng 1cm
            else:
                formatted_html += f'<p style="text-align: justify; text-indent: 1cm; margin-bottom: 6px; line-height: 1.45;">{line}</p>'
        
        # 3. CSS tạo giao diện trang A4 phẳng, chuyên nghiệp
        a4_css = """
        <style>
        .a4-container {
            background-color: #484c4f;
            padding: 25px 15px;
            display: flex;
            justify-content: center;
            border-radius: 4px;
        }
        .a4-paper {
            background-color: #ffffff !important;
            color: #000000 !important;
            width: 100%;
            padding: 40px 45px;
            font-family: 'Times New Roman', Times, serif;
            font-size: 13pt;
            box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.4);
            max-height: 620px;
            overflow-y: auto;
        }
        </style>
        """
        
        st.markdown(a4_css + f'<div class="a4-container"><div class="a4-paper">{formatted_html}</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Hàm xuất file Word (.docx)
        def generate_docx(text):
            doc = docx.Document()
            for section in doc.sections:
                section.top_margin = docx.shared.Inches(0.79)
                section.bottom_margin = docx.shared.Inches(0.79)
                section.left_margin = docx.shared.Inches(1.18)
                section.right_margin = docx.shared.Inches(0.79)
            
            p_clean = re.sub(r'[\*#_]', '', text)
            p = doc.add_paragraph(p_clean)
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

                        YÊU CẦU CHỈNH SỬA:
                        {edit_instruction}

                        Hãy cập nhật toàn bộ bản dự thảo văn bản. TUYỆT ĐỐI KHÔNG DÙNG KÝ TỰ MARKDOWN (*, #, _).
                        """
                        res_edit = model.generate_content(edit_prompt)
                        st.session_state.draft_text = res_edit.text
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi chỉnh sửa: {str(e)}")
