import streamlit as st
import google.generativeai as genai
import pypdf
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import re

# Cấu hình trang Streamlit
st.set_page_config(page_title="Phần mềm Cụ thể hóa Văn bản Hành chính", page_icon="📝", layout="wide")

# CSS tạo trang A4 màu trắng xem trước chuẩn đẹp
st.markdown("""
<style>
.a4-wrapper {
    background-color: #525659;
    padding: 25px;
    border-radius: 6px;
    display: flex;
    justify-content: center;
}
.a4-paper {
    background-color: #ffffff !important;
    color: #000000 !important;
    width: 100%;
    max-width: 800px;
    padding: 45px 55px;
    font-family: 'Times New Roman', Times, serif;
    font-size: 13pt;
    line-height: 1.42;
    box-shadow: 0px 6px 18px rgba(0,0,0,0.4);
    max-height: 650px;
    overflow-y: auto;
}
.header-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
    border: none !important;
}
.header-table td {
    vertical-align: top;
    text-align: center;
    font-family: 'Times New Roman', Times, serif;
    font-size: 12pt;
    line-height: 1.25;
    color: #000000;
    padding: 0px;
}
.title-block {
    text-align: center;
    font-weight: bold;
    font-size: 14pt;
    margin-top: 15px;
    margin-bottom: 10px;
}
.content-para {
    text-align: justify;
    text-indent: 1.27cm;
    margin-bottom: 6px;
    line-height: 1.42;
}
.heading-para {
    font-weight: bold;
    font-size: 13pt;
    margin-top: 12px;
    margin-bottom: 4px;
}
</style>
""", unsafe_allow_html=True)

# --- CỘT BÊN TRÁI: CẤU HÌNH THỂ THỨC & AI ---
with st.sidebar:
    st.title("⚙️ Cấu hình Thể thức & AI")
    the_thuc = st.radio("Chọn Khối văn bản:", ["Khối Đảng", "Khối Nhà nước"])
    if the_thuc == "Khối Đảng":
        st.info("📌 **Áp dụng:** Hướng dẫn 05-HD/VPTW của Văn phòng Trung ương Đảng")
    else:
        st.info("📌 **Áp dụng:** Nghị định 30/2020/NĐ-CP của Chính phủ")
        
    st.subheader("⚙️ Cấu hình AI")
    ai_provider = st.selectbox("AI xử lý chính", ["Google Gemini"])
    api_key = st.text_input("Gemini API key", type="password")
    model_name = st.selectbox("Model", ["gemini-3.6-flash"])

st.title("📝 Phần mềm Cụ thể hóa Văn bản Hành chính")

# --- MỤC 1 & 2 ---
col1, col2 = st.columns([1, 1])
with col1:
    st.subheader("1. File nguồn & Loại văn bản")
    uploaded_files = st.file_uploader(
        "Tải file nguồn/Đề cương (.docx, .pdf, .png, .jpg...):",
        type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
        accept_multiple_files=True
    )
    loai_vb = st.selectbox("Chọn Loại văn bản đầu ra:", ["Kế hoạch", "Công văn", "Báo cáo", "Tờ trình", "Thông báo", "Quyết định"])

with col2:
    st.subheader("2. Yêu cầu & Cơ quan ban hành")
    yeu_cau = st.text_area("Anh muốn cụ thể hóa như thế nào?:", height=100, placeholder="Soạn kế hoạch thực hiện từ văn bản đã cung cấp...")
    co_quan = st.text_input("Cơ quan ban hành dự thảo:", value="ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH" if the_thuc == "Khối Đảng" else "UBND PHƯƠNG NHƠN TRẠCH")

# --- MỤC 3: FILE MẪU RIÊNG & MẪU GỢI Ý ---
st.subheader("3. File mẫu riêng & Mẫu gợi ý")
col3_1, col3_2 = st.columns([1, 1])
with col3_1:
    custom_template_file = st.file_uploader("Tải file mẫu riêng (Nếu có file mẫu riêng):", type=["docx"], key="custom_template")
with col3_2:
    de_cuong_goy_y = st.selectbox("📌 Mẫu gợi ý / Đề cương chuẩn:", [
        "(Không chọn mẫu gợi ý)",
        "Đề cương chuẩn Hướng dẫn 05-HD/VPTW (Công tác Đảng)",
        "Đề cương Kế hoạch hành động 100 ngày Chuyển đổi số",
        "Đề cương Báo cáo kết quả thực hiện nhiệm vụ chính trị"
    ])

st.markdown("---")
btn_process = st.button("⚡ PHÂN TÍCH & CỤ THỂ HÓA VĂN BẢN", type="primary", use_container_width=True)

if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# --- XỬ LÝ AI ---
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
                
                # 1. Đọc nội dung các file nguồn đính kèm
                for uf in uploaded_files:
                    bytes_data = uf.read()
                    if uf.name.lower().endswith('.pdf'):
                        try:
                            reader = pypdf.PdfReader(io.BytesIO(bytes_data))
                            pdf_text = "".join([page.extract_text() or "" for page in reader.pages])
                            if len(pdf_text.strip()) > 50:
                                extracted_texts.append(f"--- NỘI DUNG FILE NGUỒN {uf.name} ---\n" + pdf_text)
                            else:
                                content_parts.append({"mime_type": "application/pdf", "data": bytes_data})
                        except:
                            content_parts.append({"mime_type": "application/pdf", "data": bytes_data})
                    elif uf.name.lower().endswith('.docx'):
                        doc_file = docx.Document(io.BytesIO(bytes_data))
                        docx_text = "\n".join([p.text for p in doc_file.paragraphs])
                        extracted_texts.append(f"--- NỘI DUNG FILE NGUỒN {uf.name} ---\n" + docx_text)
                    else:
                        content_parts.append({"mime_type": uf.type, "data": bytes_data})
                
                # 2. Đọc cấu trúc từ File Mẫu Riêng (nếu có tải lên ở Mục 3)
                custom_template_prompt = ""
                if custom_template_file is not None:
                    try:
                        doc_tpl = docx.Document(io.BytesIO(custom_template_file.read()))
                        tpl_text = "\n".join([p.text for p in doc_tpl.paragraphs if p.text.strip()])
                        if tpl_text:
                            custom_template_prompt = f"\nBẮT BUỘC BÁM SÁT MẪU VĂN BẢN RIÊNG DƯỚI ĐÂY CỦA ĐƠN VỊ ĐỂ DỰ THẢO VĂN BẢN (GIỮ NGUYÊN BỐ CỤC, MỤC LỤC VÀ CÁCH XƯNG HÔ):\n--- BẮT ĐẦU MẪU RIÊNG ---\n{tpl_text}\n--- KẾT THÚC MẪU RIÊNG ---\n"
                    except Exception as tpl_err:
                        st.warning(f"Không thể đọc file mẫu riêng: {str(tpl_err)}")

                # 3. Đưa thông tin mẫu gợi ý vào prompt
                de_cuong_prompt = ""
                if de_cuong_goy_y != "(Không chọn mẫu gợi ý)":
                    de_cuong_prompt = f"\nÁP DỤNG ĐỀ CƯỜNG MẪU GỢI Ý: {de_cuong_goy_y}"

                # 4. Xây dựng Prompt tổng hợp
                prompt = f"""
                Bạn là chuyên gia soạn thảo văn bản hành chính Việt Nam.
                Hãy soạn thảo phần NỘI DUNG của 01 dự thảo văn bản hoàn chỉnh dựa trên tài liệu đính kèm.
                
                THỂ THỨC: {the_thuc} | CƠ QUAN BAN HÀNH: {co_quan} | LOẠI VĂN BẢN: {loai_vb}
                YÊU CẦU CỤ THỂ HÓA: {yeu_cau}
                {de_cuong_prompt}
                {custom_template_prompt}
                
                DỮ LIỆU TÀI LIỆU NGUỒN THAM KHẢO:
                {"".join(extracted_texts)}
                
                QUY CẮC BẮT BUỘC:
                1. BẮT ĐẦU TRỰC TIẾP BẰNG TÊN LOẠI VĂN BẢN (Viết hoa in đậm, ví dụ: KẾ HOẠCH) và Trích yếu nội dung.
                2. TUYỆT ĐỐI KHÔNG VIẾT Quốc hiệu, Tiêu ngữ, Tên cơ quan ban hành, Số/Ký hiệu, Ngày tháng ở đầu bài (Vì giao diện đã có khung bảng riêng).
                3. TUYỆT ĐỐI KHÔNG DÙNG KÝ TỰ MARKDOWN (*, #, _).
                4. Soạn thảo đầy đủ Căn cứ pháp lý, Các mục I, II, III..., Nơi nhận, Chức vụ người ký.
                """
                content_parts.insert(0, prompt)
                
                response = model.generate_content(content_parts)
                st.session_state.draft_text = response.text
                st.session_state.chat_history = []
                st.success("Đã cụ thể hóa văn bản thành công!")
            except Exception as e:
                st.error(f"Lỗi xử lý: {str(e)}")

# --- HIỂN THỊ DỰ THẢO A4 & CHAT AI ---
if st.session_state.draft_text:
    res_col1, res_col2 = st.columns([1.2, 0.8])
    
    with res_col1:
        st.subheader("📄 Bản dự thảo trang Word (A4)")
        
        # Làm sạch ký tự Markdown rác
        clean_text = re.sub(r'[\*#_]', '', st.session_state.draft_text)
        lines = [l.strip() for l in clean_text.split('\n') if l.strip()]
        
        # Lọc bỏ tiêu đề đúp ở đầu bài
        filtered_lines = []
        for l in lines:
            l_up = l.upper()
            if any(k in l_up for k in ["ĐẢNG BỘ", "ĐẢNG CỘNG SẢN", "CỘNG HÒA XÃ HỘI", "ĐỘC LẬP - TỰ DO", "NHƠN TRẠCH, NGÀY"]) and len(l) < 80:
                continue
            if ("UBND" in l_up or "ĐẢNG ỦY" in l_up) and len(l) < 60 and not l_up.startswith("KẾ HOẠCH") and not l_up.startswith("CÔNG VĂN"):
                continue
            if re.match(r'^SỐ\s*:', l_up) or re.match(r'^SỐ\s*-', l_up):
                continue
            filtered_lines.append(l)

        # 1. Khung Quốc hiệu 2 cột ẩn viền chuẩn
        if the_thuc == "Khối Đảng":
            header_table = f"""
            <table class="header-table">
                <tr>
                    <td style="width: 48%;">
                        <b>ĐẢNG BỘ THÀNH PHỐ ĐỒNG NAI</b><br>
                        <b><u>{co_quan.upper()}</u></b><br>
                        <span style="font-size: 8pt;">*</span><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;-KH/ĐU
                    </td>
                    <td style="width: 52%;">
                        <b><u>ĐẢNG CỘNG SẢN VIỆT NAM</u></b><br><br>
                        <i>Nhơn Trạch, ngày &nbsp;&nbsp;&nbsp; tháng 8 năm 2026</i>
                    </td>
                </tr>
            </table>
            """
        else:
            header_table = f"""
            <table class="header-table">
                <tr>
                    <td style="width: 45%;">
                        UBND THÀNH PHỐ ĐỒNG NAI<br>
                        <b><u>{co_quan.upper()}</u></b><br>
                        <span style="font-size: 8pt;">*</span><br>
                        Số: &nbsp;&nbsp;&nbsp;&nbsp;/UBND-VP
                    </td>
                    <td style="width: 55%;">
                        <b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br>
                        <b><u>Độc lập - Tự do - Hạnh phúc</u></b><br>
                        <i>Nhơn Trạch, ngày &nbsp;&nbsp;&nbsp; tháng 8 năm 2026</i>
                    </td>
                </tr>
            </table>
            """

        # 2. Render nội dung
        body_content = ""
        for line in filtered_lines:
            if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', line):
                body_content += f'<div class="heading-para">{line}</div>'
            elif re.match(r'^\d+\.', line) and len(line) < 80:
                body_content += f'<div class="heading-para">{line}</div>'
            elif line.isupper() and len(line) < 100:
                body_content += f'<div class="title-block">{line}</div>'
            elif line.startswith("Kính gửi:") or line.startswith("-"):
                body_content += f'<div style="text-align: left; margin-bottom: 4px; padding-left: 15px;">{line}</div>'
            else:
                body_content += f'<div class="content-para">{line}</div>'

        st.markdown(f'<div class="a4-wrapper"><div class="a4-paper">{header_table}{body_content}</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # 3. Hàm xuất file Word chuẩn thể thức hành chính
        def generate_docx(lines_data, agency_name, form_type):
            doc = docx.Document()
            for section in doc.sections:
                section.top_margin = Cm(2)
                section.bottom_margin = Cm(2)
                section.left_margin = Cm(3)
                section.right_margin = Cm(2)
            
            table = doc.add_table(rows=1, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            
            cell_left, cell_right = table.cell(0, 0), table.cell(0, 1)
            cell_left.width, cell_right.width = Cm(8.5), Cm(8.5)
            
            p_left = cell_left.paragraphs[0]
            p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_left.paragraph_format.line_spacing = 1.15
            
            if form_type == "Khối Đảng":
                r1 = p_left.add_run("ĐẢNG BỘ THÀNH PHỐ ĐỒNG NAI\n")
                r1.font.name, r1.font.size, r1.font.bold = 'Times New Roman', Pt(12), True
                r2 = p_left.add_run(f"{agency_name.upper()}\n*\n")
                r2.font.name, r2.font.size, r2.font.bold, r2.font.underline = 'Times New Roman', Pt(12), True, True
                r3 = p_left.add_run("Số:       -KH/ĐU")
                r3.font.name, r3.font.size = 'Times New Roman', Pt(12)
            else:
                r1 = p_left.add_run("UBND THÀNH PHỐ ĐỒNG NAI\n")
                r1.font.name, r1.font.size = 'Times New Roman', Pt(12)
                r2 = p_left.add_run(f"{agency_name.upper()}\n*\n")
                r2.font.name, r2.font.size, r2.font.bold, r2.font.underline = 'Times New Roman', Pt(12), True, True
                r3 = p_left.add_run("Số:       /UBND-VP")
                r3.font.name, r3.font.size = 'Times New Roman', Pt(12)

            p_right = cell_right.paragraphs[0]
            p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_right.paragraph_format.line_spacing = 1.15
            
            if form_type == "Khối Đảng":
                r1 = p_right.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n")
                r1.font.name, r1.font.size, r1.font.bold, r1.font.underline = 'Times New Roman', Pt(12), True, True
                r2 = p_right.add_run("\nNhơn Trạch, ngày     tháng 8 năm 2026")
                r2.font.name, r2.font.size, r2.font.italic = 'Times New Roman', Pt(12), True
            else:
                r1 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
                r1.font.name, r1.font.size, r1.font.bold = 'Times New Roman', Pt(12), True
                r2 = p_right.add_run("Độc lập - Tự do - Hạnh phúc\n")
                r2.font.name, r2.font.size, r2.font.bold, r2.font.underline = 'Times New Roman', Pt(12.5), True, True
                r3 = p_right.add_run("Nhơn Trạch, ngày     tháng 8 năm 2026")
                r3.font.name, r3.font.size, r3.font.italic = 'Times New Roman', Pt(12), True

            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_after = Pt(6)

            for line in lines_data:
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = 1.2
                p.paragraph_format.space_after = Pt(4)
                
                if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', line) or line.isupper():
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if line.isupper() else WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name, run.font.size, run.font.bold = 'Times New Roman', Pt(14 if line.isupper() else 13), True
                elif re.match(r'^\d+\.', line) and len(line) < 80:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name, run.font.size, run.font.bold = 'Times New Roman', Pt(13), True
                elif line.startswith("Kính gửi:") or line.startswith("-"):
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name, run.font.size = 'Times New Roman', Pt(12 if line.startswith("-") else 13)
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.first_line_indent = Cm(1.27)
                    run = p.add_run(line)
                    run.font.name, run.font.size = 'Times New Roman', Pt(13)

            bio = io.BytesIO()
            doc.save(bio)
            return bio.getvalue()

        st.download_button(
            label="📥 TẢI VỀ FILE WORD (.DOCX)",
            data=generate_docx(filtered_lines, co_quan, the_thuc),
            file_name="Du_Thao_Van_Ban_Hanh_Chinh.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

    # --- CỘT PHẢI: CHAT AI SỬA ĐỔI ---
    with res_col2:
        st.subheader("💬 Chat AI sửa đổi (Google Gemini)")
        edit_instruction = st.text_area("Nhập yêu cầu chỉnh sửa văn bản...", height=120, placeholder="VD: 'Sửa tên người ký thành Trần Văn A'...")
        
        if st.button("Chỉnh sửa dự thảo", use_container_width=True):
            if edit_instruction and api_key:
                with st.spinner("AI đang cập nhật lại dự thảo..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel(model_name)
                        edit_prompt = f"BẢN DỰ THẢO HIỆN TẠI:\n{st.session_state.draft_text}\n\nYÊU CẦU CHỈNH SỬA:\n{edit_instruction}\n\nHãy cập nhật toàn bộ bản dự thảo văn bản. TUYỆT ĐỐI KHÔNG DÙNG MARKDOWN (*, #, _)."
                        res_edit = model.generate_content(edit_prompt)
                        st.session_state.draft_text = res_edit.text
                        st.session_state.chat_history.append(edit_instruction)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi: {str(e)}")

        if st.session_state.chat_history:
            st.markdown("---")
            st.markdown("**Lịch sử chỉnh sửa:**")
            for idx, cmd in enumerate(reversed(st.session_state.chat_history)):
                st.info(f"🔴 {cmd}")
                st.success("✅ Đã cập nhật văn bản lên trang Word!")
