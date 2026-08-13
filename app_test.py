import streamlit as st
import google.generativeai as genai
import pypdf
import docx
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import re

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Phần mềm Cụ thể hóa Văn bản Hành chính", page_icon="📝", layout="wide")

# --- CỘT BÊN TRÁI: CẤU HÌNH THỂ THỨC & AI ---
with st.sidebar:
    st.title("⚙️ Cấu hình Thể thức & AI")
    
    st.subheader("📌 Thể thức Văn bản")
    the_thuc = st.radio(
        "Chọn Khối văn bản:",
        ["Khối Đảng", "Khối Nhà nước"],
        help="Khối Đảng: Hướng dẫn 05-HD/VPTW | Khối Nhà nước: Nghị định 30/2020/NĐ-CP"
    )
    
    if the_thuc == "Khối Đảng":
        st.info("📌 **Áp dụng:** Hướng dẫn 05-HD/VPTW của Văn phòng Trung ương Đảng")
    else:
        st.info("📌 **Áp dụng:** Nghị định 30/2020/NĐ-CP của Chính phủ")
        
    st.subheader("⚙️ Cấu hình AI")
    ai_provider = st.selectbox("AI xử lý chính", ["Google Gemini"])
    api_key = st.text_input("Gemini API key", type="password", help="Dán mã API của anh tại đây")
    model_name = st.selectbox("Model", ["gemini-3.6-flash"])

# --- TIÊU ĐỀ CHÍNH ---
st.title("📝 Phần mềm Cụ thể hóa Văn bản Hành chính")

# --- MỤC 1 & 2: NHẬP FILE NGUỒN VÀ YÊU CẦU ---
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. File nguồn & Loại văn bản")
    uploaded_files = st.file_uploader(
        "Tải file nguồn/Đề cương (.docx, .pdf, .png, .jpg...):",
        type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
        accept_multiple_files=True
    )
    loai_vb = st.selectbox(
        "Chọn Loại văn bản đầu ra:",
        ["Kế hoạch", "Công văn", "Báo cáo", "Tờ trình", "Thông báo", "Quyết định"]
    )

with col2:
    st.subheader("2. Yêu cầu & Cơ quan ban hành")
    yeu_cau = st.text_area(
        "Anh muốn cụ thể hóa như thế nào?:",
        height=100,
        placeholder="Ví dụ: Cụ thể hóa kế hoạch trên thành kế hoạch thực hiện của phường..."
    )
    co_quan = st.text_input(
        "Cơ quan ban hành dự thảo:",
        value="ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH" if the_thuc == "Khối Đảng" else "UBND PHƯƠNG NHƠN TRẠCH"
    )

# --- MỤC 3: FILE MẪU RIÊNG & MẪU GỢI Ý (ĐÃ KHÔI PHÚC) ---
st.subheader("3. File mẫu riêng & Mẫu gợi ý")
col3_1, col3_2 = st.columns([1, 1])

with col3_1:
    custom_template_file = st.file_uploader(
        "Tải file mẫu riêng (Nếu có file mẫu riêng):",
        type=["docx"],
        key="custom_template"
    )

with col3_2:
    de_cuong_goy_y = st.selectbox("📌 Mẫu gợi ý / Đề cương chuẩn:", [
        "(Không chọn mẫu gợi ý)",
        "Đề cương chuẩn Hướng dẫn 05-HD/VPTW (Công tác Đảng)",
        "Đề cương Kế hoạch hành động 100 ngày Chuyển đổi số",
        "Đề cương Báo cáo kết quả thực hiện nhiệm vụ chính trị",
        "Đề cương Quyết định thành lập Ban Chỉ đạo / Tổ công tác"
    ])

st.markdown("---")
btn_process = st.button("⚡ PHÂN TÍCH & CỤ THỂ HÓA VĂN BẢN", type="primary", use_container_width=True)

# Khởi tạo Session State
if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# --- XỬ LÝ TẠO VĂN BẢN BAN ĐẦU ---
if btn_process:
    if not api_key:
        st.error("Vui lòng nhập Gemini API Key ở cột bên trái!")
    elif not uploaded_files and not yeu_cau:
        st.warning("Vui lòng tải file nguồn hoặc nhập yêu cầu!")
    else:
        with st.spinner("Đang phân tích dữ liệu và tổng hợp văn bản chuẩn thể thức..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name)
                
                content_parts = []
                extracted_texts = []
                
                # Trích xuất dữ liệu đa file
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
                
                # Bổ sung thông tin mẫu gợi ý nếu có chọn
                template_context = ""
                if de_cuong_goy_y != "(Không chọn mẫu gợi ý)":
                    template_context = f"\nÁP DỤNG ĐỀ CƯỜNG MẪU: {de_cuong_goy_y}"
                
                prompt = f"""
                Bạn là chuyên gia soạn thảo văn bản hành chính Việt Nam.
                Hãy soạn thảo 01 dự thảo văn bản hoàn chỉnh căn cứ vào nội dung được cấp.
                
                THỂ THỨC: {the_thuc} ({'Hướng dẫn 05-HD/VPTW' if the_thuc == 'Khối Đảng' else 'Nghị định 30/2020/NĐ-CP'})
                CƠ QUAN BAN HÀNH: {co_quan}
                LOẠI VĂN BẢN: {loai_vb}
                YÊU CẦU CỤ THỂ HÓA: {yeu_cau}
                {template_context}
                
                DỮ LIỆU TÀI LIỆU NGUỒN:
                {"".join(extracted_texts)}
                
                QUY TẮC ĐỊNH DẠNG (RẤT QUAN TRỌNG):
                1. TUYỆT ĐỐI KHÔNG DÙNG KÝ TỰ MARKDOWN như *, #, _ trong toàn bộ văn bản.
                2. Soạn thảo trực tiếp phần Nội dung chính của văn bản (Gồm Tên loại, Trích yếu, các Mục I, II, III..., Nơi nhận, Chức vụ người ký).
                3. Lời văn chuẩn mực hành chính, đúng thể thức công tác Đảng/Nhà nước.
                """
                content_parts.insert(0, prompt)
                
                response = model.generate_content(content_parts)
                st.session_state.draft_text = response.text
                st.session_state.chat_history = []  # Reset lịch sử chat khi tạo mới
                st.success("Đã cụ thể hóa văn bản thành công!")
            except Exception as e:
                st.error(f"Lỗi xử lý: {str(e)}")

# --- HÌNH ẢNH HIỂN THỊ KẾT QUẢ & SỬA ĐỔI ---
if st.session_state.draft_text:
    res_col1, res_col2 = st.columns([1.2, 0.8])
    
    # === CỘT TRÁI: KHUNG XEM TRƯỚC TRANG WORD A4 CHUẨN ĐẸP ===
    with res_col1:
        st.subheader("📄 Bản dự thảo trang Word (A4)")
        
        # Làm sạch ký tự Markdown rác
        clean_text = re.sub(r'[\*#_]', '', st.session_state.draft_text)
        
        # Phân tích dòng để dựng HTML chuẩn
        lines = [l.strip() for l in clean_text.split('\n') if l.strip()]
        body_html = ""
        
        for line in lines:
            if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', line):
                body_html += f'<p style="font-weight: bold; font-size: 13pt; margin-top: 14px; margin-bottom: 4px; text-align: left;">{line}</p>'
            elif re.match(r'^\d+\.', line) and len(line) < 80:
                body_html += f'<p style="font-weight: bold; font-size: 13pt; margin-top: 8px; margin-bottom: 3px; text-align: left;">{line}</p>'
            elif line.isupper() and len(line) < 100:
                body_html += f'<p style="text-align: center; font-weight: bold; font-size: 14pt; margin-top: 14px; margin-bottom: 6px;">{line}</p>'
            elif line.startswith("Kính gửi:") or line.startswith("-"):
                body_html += f'<p style="text-align: left; margin-bottom: 4px; padding-left: 15px;">{line}</p>'
            else:
                body_html += f'<p style="text-align: justify; text-indent: 1.27cm; margin-bottom: 6px; line-height: 1.45;">{line}</p>'
        
        # Dựng bảng 2 cột Quốc hiệu - Tiêu ngữ chuẩn thể thức
        if the_thuc == "Khối Đảng":
            header_html = f"""
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; border: none;">
                <tr>
                    <td style="width: 48%; text-align: center; vertical-align: top; font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.2;">
                        <b>ĐẢNG BỘ THÀNH PHỐ ĐỒNG NAI</b><br>
                        <b>{co_quan.upper()}</b><br>
                        <span style="font-size: 9pt;">*</span><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;-KH/ĐU
                    </td>
                    <td style="width: 52%; text-align: center; vertical-align: top; font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.2;">
                        <b><u>ĐẢNG CỘNG SẢN VIỆT NAM</u></b><br><br>
                        <i>Nhơn Trạch, ngày &nbsp;&nbsp;&nbsp; tháng 8 năm 2026</i>
                    </td>
                </tr>
            </table>
            """
        else:
            header_html = f"""
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; border: none;">
                <tr>
                    <td style="width: 45%; text-align: center; vertical-align: top; font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.2;">
                        UBND THÀNH PHỐ ĐỒNG NAI<br>
                        <b>{co_quan.upper()}</b><br>
                        <span style="font-size: 9pt;">*</span><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;/UBND-VP
                    </td>
                    <td style="width: 55%; text-align: center; vertical-align: top; font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.2;">
                        <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b><u>Độc lập - Tự do - Hạnh phúc</u></b><br>
                        <i>Nhơn Trạch, ngày &nbsp;&nbsp;&nbsp; tháng 8 năm 2026</i>
                    </td>
                </tr>
            </table>
            """

        a4_css = """
        <style>
        .a4-wrapper {
            background-color: #3a3d40;
            padding: 20px;
            border-radius: 6px;
            display: flex;
            justify-content: center;
        }
        .a4-paper {
            background-color: #ffffff !important;
            color: #000000 !important;
            width: 100%;
            padding: 45px 50px;
            font-family: 'Times New Roman', Times, serif;
            font-size: 13pt;
            box-shadow: 0px 4px 15px rgba(0,0,0,0.5);
            max-height: 600px;
            overflow-y: auto;
        }
        </style>
        """
        
        st.markdown(a4_css + f'<div class="a4-wrapper"><div class="a4-paper">{header_html}{body_html}</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- HÀM TẠO FILE WORD (.DOCX) CHUẨN THỂ THỨC ---
        def generate_docx(text, agency_name, form_type):
            doc = docx.Document()
            
            # Căn lề A4 chuẩn thể thức hành chính (Trái 3cm, Trên/Dưới/Phải 2cm)
            for section in doc.sections:
                section.top_margin = Cm(2)
                section.bottom_margin = Cm(2)
                section.left_margin = Cm(3)
                section.right_margin = Cm(2)
            
            # 1. Tạo bảng 2 cột ẩn viền cho Phần Đầu trang
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            
            cell_left = table.cell(0, 0)
            cell_right = table.cell(0, 1)
            cell_left.width = Cm(8.5)
            cell_right.width = Cm(8.5)
            
            # Nội dung cột trái
            p_left = cell_left.paragraphs[0]
            p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_left.paragraph_format.line_spacing = 1.15
            
            if form_type == "Khối Đảng":
                run = p_left.add_run("ĐẢNG BỘ THÀNH PHỐ ĐỒNG NAI\n")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.bold = True
                
                run = p_left.add_run(f"{agency_name.upper()}\n*\n")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.bold = True
                
                run = p_left.add_run("Số:       -KH/ĐU")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
            else:
                run = p_left.add_run("UBND THÀNH PHỐ ĐỒNG NAI\n")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                
                run = p_left.add_run(f"{agency_name.upper()}\n*\n")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.bold = True
                
                run = p_left.add_run("Số:       /UBND-VP")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

            # Nội dung cột phải
            p_right = cell_right.paragraphs[0]
            p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_right.paragraph_format.line_spacing = 1.15
            
            if form_type == "Khối Đảng":
                run = p_right.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.underline = True
                
                run = p_right.add_run("\nNhơn Trạch, ngày     tháng 8 năm 2026")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.italic = True
            else:
                run = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.bold = True
                
                run = p_right.add_run("Độc lập - Tự do - Hạnh phúc\n")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12.5)
                run.font.bold = True
                run.font.underline = True
                
                run = p_right.add_run("Nhơn Trạch, ngày     tháng 8 năm 2026")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.italic = True

            # Dòng cách sau bảng
            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_after = Pt(6)

            # 2. Đưa nội dung chính vào file Word
            p_clean = re.sub(r'[\*#_]', '', text)
            word_lines = [l.strip() for l in p_clean.split('\n') if l.strip()]
            
            for line in word_lines:
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = 1.2
                p.paragraph_format.space_after = Pt(4)
                
                # Tiêu đề các mục lớn
                if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', line) or line.isupper():
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if line.isupper() else WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14 if line.isupper() else 13)
                    run.font.bold = True
                # Các tiểu mục 1, 2, 3...
                elif re.match(r'^\d+\.', line) and len(line) < 80:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(13)
                    run.font.bold = True
                # Nơi nhận / Kính gửi
                elif line.startswith("Kính gửi:") or line.startswith("-"):
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12 if line.startswith("-") else 13)
                # Nội dung đoạn văn
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.first_line_indent = Cm(1.27)  # Lùi đầu dòng 1.27cm
                    run = p.add_run(line)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(13)

            bio = io.BytesIO()
            doc.save(bio)
            return bio.getvalue()

        # Nút tải file Word (.docx)
        st.download_button(
            label="📥 TẢI VỀ FILE WORD (.DOCX)",
            data=generate_docx(st.session_state.draft_text, co_quan, the_thuc),
            file_name="Du_Thao_Van_Ban_Hanh_Chinh.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

    # === CỘT PHẢI: KHUNG CHAT AI SỬA ĐỔI (LƯU LỊCH SỬ SỬA) ===
    with res_col2:
        st.subheader("💬 Chat AI sửa đổi (Google Gemini)")
        
        edit_instruction = st.text_area(
            "Nhập yêu cầu chỉnh sửa văn bản...",
            height=120,
            placeholder="VD: 'Bổ sung mốc thời gian 30/9 vào Mục II', 'Sửa tên người ký thành Trần Văn A'..."
        )
        
        if st.button("Chỉnh sửa dự thảo", use_container_width=True):
            if not edit_instruction:
                st.warning("Vui lòng nhập nội dung cần sửa!")
            elif not api_key:
                st.error("Vui lòng nhập Gemini API Key!")
            else:
                with st.spinner("AI đang cập nhật lại dự thảo..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel(model_name)
                        edit_prompt = f"""
                        BẢN DỰ THẢO HIỆN TẠI:
                        {st.session_state.draft_text}

                        YÊU CẦU CHỈNH SỬA TỪ NGƯỜI DÙNG:
                        {edit_instruction}

                        Hãy cập nhật toàn bộ bản dự thảo văn bản trên. TUYỆT ĐỐI KHÔNG DÙNG KÝ TỰ MARKDOWN (*, #, _).
                        """
                        res_edit = model.generate_content(edit_prompt)
                        st.session_state.draft_text = res_edit.text
                        
                        # Lưu lịch sử chat
                        st.session_state.chat_history.append(edit_instruction)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi chỉnh sửa: {str(e)}")

        # Hiển thị lịch sử các câu lệnh đã sửa
        if st.session_state.chat_history:
            st.markdown("---")
            st.markdown("**Lịch sử chỉnh sửa:**")
            for idx, cmd in enumerate(reversed(st.session_state.chat_history)):
                st.info(f"🔴 **Lệnh sửa:** {cmd}")
                st.success("✅ Đã cập nhật văn bản lên trang Word!")
