import streamlit as st
import google.generativeai as genai
import pypdf
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import re

st.set_page_config(page_title="Phần mềm Cụ thể hóa Văn bản Hành chính", page_icon="📝", layout="wide")

SYMBOL_MAP = {
    "Kế hoạch": "KH",
    "Công văn": "CV",
    "Báo cáo": "BC",
    "Tờ trình": "TTr",
    "Thông báo": "TB",
    "Quyết định": "QĐ"
}

# Hàm thêm đường kẻ dưới paragraph trong Word (đẩy xa dấu nặng 5pt)
def add_paragraph_bottom_border(paragraph, sz="6", space="5", color="000000"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="{space}" w:color="{color}"/></w:pBdr>')
    pPr.append(pBdr)

# CSS thiết kế giao diện A4 xem trước
a4_css = """
<style>
.a4-wrapper {
    background-color: #525659;
    padding: 15px;
    border-radius: 6px;
    display: flex;
    justify-content: center;
    width: 100%;
}
.a4-paper {
    background-color: #ffffff !important;
    color: #000000 !important;
    width: 100%;
    padding: 25px 30px;
    font-family: 'Times New Roman', Times, serif;
    font-size: 10.5pt;
    line-height: 1.35;
    box-shadow: 0px 4px 12px rgba(0,0,0,0.4);
    max-height: 600px;
    overflow-y: auto;
    box-sizing: border-box;
}
.header-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 15px;
    border: none !important;
}
.header-table td {
    vertical-align: top;
    font-family: 'Times New Roman', Times, serif;
    font-size: 10pt;
    line-height: 1.2;
    color: #000000;
    padding: 0px;
}
.custom-underline {
    display: inline-block;
    border-bottom: 1px solid #000000;
    padding-bottom: 2px;
    line-height: 1.1;
}
.title-block {
    text-align: center;
    font-weight: bold;
    font-size: 12pt;
    margin-top: 10px;
    margin-bottom: 4px;
}
.trich-yeu-block {
    text-align: center;
    font-weight: bold;
    font-size: 11pt;
    margin-top: 4px;
    margin-bottom: 4px;
}
.short-line {
    width: 35%;
    margin: 6px auto 12px auto;
    border: 0;
    border-top: 1px solid #000000;
}
.content-para {
    text-align: justify;
    text-indent: 1cm;
    margin-bottom: 5px;
    line-height: 1.35;
}
.heading-para {
    font-weight: bold;
    font-size: 10.5pt;
    margin-top: 10px;
    margin-bottom: 3px;
}
</style>
"""
st.markdown(a4_css, unsafe_allow_html=True)

# --- CỘT BÊN TRÁI: CẤU HÌNH ---
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
    yeu_cau = st.text_area("Anh muốn cụ thể hóa như thế nào?:", height=100, placeholder="Soạn thảo văn bản theo yêu cầu chỉ đạo...")
    co_quan = st.text_input("Cơ quan ban hành dự thảo:", value="ĐẢNG ỦY PHƯƠNG NHƠN TRẠCH" if the_thuc == "Khối Đảng" else "UBND PHƯƠNG NHƠN TRẠCH")

# --- MỤC 3 ---
st.subheader("3. File mẫu riêng & Mẫu gợi ý")
col3_1, col3_2 = st.columns([1, 1])
with col3_1:
    custom_template_file = st.file_uploader("Tải file mẫu riêng (Chỉ lấy thể thức/khung mẫu):", type=["docx"], key="custom_template")
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
                
                custom_template_prompt = ""
                if custom_template_file is not None:
                    try:
                        doc_tpl = docx.Document(io.BytesIO(custom_template_file.read()))
                        tpl_text = "\n".join([p.text for p in doc_tpl.paragraphs if p.text.strip()])
                        if tpl_text:
                            custom_template_prompt = f"\nBÁM SÁT KHUNG MẪU VĂN BẢN ĐÍNH KÈM (Chỉ học theo thể thức, bố cục mục I, II, III và phong cách trình bày):\n--- MẪU THỂ THỨC KHUNG ---\n{tpl_text}\n--- KẾT THÚC MẪU ---\n"
                    except Exception as tpl_err:
                        st.warning(f"Lỗi đọc file mẫu: {str(tpl_err)}")

                de_cuong_prompt = ""
                if de_cuong_goy_y != "(Không chọn mẫu gợi ý)":
                    de_cuong_prompt = f"\nÁP DỤNG ĐỀ CƯỜNG: {de_cuong_goy_y}"

                rule_doc_type = ""
                if loai_vb == "Công văn":
                    rule_doc_type = """
                    ĐÂY LÀ CÔNG VĂN HÀNH CHÍNH:
                    1. TUYỆT ĐỐI KHÔNG VIẾT tiêu đề "CÔNG VĂN" căn giữa trang.
                    2. Dòng đầu tiên phải ghi rõ Trích yếu nội dung dạng: "V/v [Tóm tắt nội dung công văn]"
                    3. Ngay sau trích yếu là dòng "Kính gửi: [Tên các cơ quan/đơn vị nhận công văn]".
                    """
                else:
                    rule_doc_type = f"""
                    ĐÂY LÀ VĂN BẢN CÓ TÊN LOẠI ({loai_vb.upper()}):
                    1. Dòng đầu tiên ghi TÊN LOẠI VĂN BẢN viết hoa căn giữa (Ví dụ: {loai_vb.upper()}).
                    2. Dòng tiếp theo ghi TRÍCH YẾU NỘI DUNG của {loai_vb}.
                    """

                prompt = f"""
                Bạn là chuyên gia soạn thảo văn bản hành chính Việt Nam.
                Hãy soạn thảo nội dung của 01 dự thảo văn bản hoàn chỉnh dựa trên tài liệu đính kèm.
                
                THỂ THỨC: {the_thuc} | CƠ QUAN BAN HÀNH: {co_quan} | LOẠI VĂN BẢN: {loai_vb}
                YÊU CẦU CỤ THỂ HÓA: {yeu_cau}
                {de_cuong_prompt}
                {custom_template_prompt}
                
                DỮ LIỆU TÀI LIỆU NGUỒN:
                {"".join(extracted_texts)}
                
                QUY CẮC THỂ THỨC BẮT BUỘC:
                {rule_doc_type}
                - TUYỆT ĐỐI KHÔNG VIẾT Quốc hiệu, Tiêu ngữ, Tên cơ quan ban hành, Số/Ký hiệu, Ngày tháng ở đầu bài (Vì giao diện đã tự chèn).
                - TUYỆT ĐỐI KHÔNG DÙNG KÝ TỰ MARKDOWN (*, #, _).
                - Nội dung bao gồm Căn cứ pháp lý, Các mục nội dung, Nơi nhận, Chức vụ người ký.
                """
                content_parts.insert(0, prompt)
                
                response = model.generate_content(content_parts)
                st.session_state.draft_text = response.text
                st.session_state.chat_history = []
                st.success("Đã cụ thể hóa văn bản thành công!")
            except Exception as e:
                st.error(f"Lỗi xử lý: {str(e)}")

# --- GIAO DIỆN HIỂN THỊ 2 CỘT ---
if st.session_state.draft_text:
    res_col1, res_col2 = st.columns([1.2, 0.8])
    
    with res_col1:
        st.subheader("📄 Bản dự thảo trang Word (A4)")
        
        clean_text = re.sub(r'[\*#_]', '', st.session_state.draft_text)
        lines = [l.strip() for l in clean_text.split('\n') if l.strip()]
        
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

        trich_yeu_cv = ""
        if loai_vb == "Công văn" and filtered_lines:
            if filtered_lines[0].startswith("V/v") or filtered_lines[0].startswith("Về việc"):
                trich_yeu_cv = filtered_lines.pop(0)

        type_code = SYMBOL_MAP.get(loai_vb, "CV")
        so_ky_hieu = f"-{type_code}/ĐU" if the_thuc == "Khối Đảng" else f"/{type_code}-UBND"

        if the_thuc == "Khối Đảng":
            sub_cv = f"<br><br><i>{trich_yeu_cv}</i>" if trich_yeu_cv else ""
            header_table = f'<table class="header-table"><tr><td style="width: 48%; text-align: center;"><b>ĐẢNG BỘ THÀNH PHỐ ĐỒNG NAI</b><br><b>{co_quan.upper()}</b><br><span style="font-size: 7pt;">*</span><br>Số: &nbsp;&nbsp;&nbsp;&nbsp;{so_ky_hieu}{sub_cv}</td><td style="width: 52%; text-align: center;"><span class="custom-underline"><b>ĐẢNG CỘNG SẢN VIỆT NAM</b></span><br><br><i>Nhơn Trạch, ngày &nbsp;&nbsp;&nbsp; tháng 8 năm 2026</i></td></tr></table>'
        else:
            sub_cv = f"<br><br><i>{trich_yeu_cv}</i>" if trich_yeu_cv else ""
            header_table = f'<table class="header-table"><tr><td style="width: 45%; text-align: center;">UBND THÀNH PHỐ ĐỒNG NAI<br><b><span class="custom-underline">{co_quan.upper()}</span></b><br><span style="font-size: 7pt;">*</span><br>Số: &nbsp;&nbsp;&nbsp;&nbsp;{so_ky_hieu}{sub_cv}</td><td style="width: 55%; text-align: center;"><b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br><b><span class="custom-underline">Độc lập - Tự do - Hạnh phúc</span></b><br><i>Nhơn Trạch, ngày &nbsp;&nbsp;&nbsp; tháng 8 năm 2026</i></td></tr></table>'

        body_content = ""
        is_trich_yeu = False
        
        for line in filtered_lines:
            if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', line):
                body_content += f'<div class="heading-para">{line}</div>'
                is_trich_yeu = False
            elif re.match(r'^\d+\.', line) and len(line) < 80:
                body_content += f'<div class="heading-para">{line}</div>'
                is_trich_yeu = False
            elif line.isupper() and len(line) < 100 and loai_vb != "Công văn":
                body_content += f'<div class="title-block">{line}</div>'
                is_trich_yeu = True
            elif is_trich_yeu and loai_vb != "Công văn":
                body_content += f'<div class="trich-yeu-block">{line}</div><hr class="short-line">'
                is_trich_yeu = False
            elif line.startswith("Kính gửi:") or line.startswith("-"):
                body_content += f'<div style="text-align: left; margin-bottom: 4px; padding-left: 10px;">{line}</div>'
                is_trich_yeu = False
            else:
                body_content += f'<div class="content-para">{line}</div>'
                is_trich_yeu = False

        full_a4_html = f'<div class="a4-wrapper"><div class="a4-paper">{header_table}{body_content}</div></div>'
        st.markdown(full_a4_html, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # Hàm xuất file Word (.docx) chuẩn không dính gạch vào dấu nặng
        def generate_docx(lines_data, agency_name, form_type, doc_type_str, cv_subject):
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
            
            t_code = SYMBOL_MAP.get(doc_type_str, "CV")
            code_str = f"-{t_code}/ĐU" if form_type == "Khối Đảng" else f"/{t_code}-UBND"

            # Xử lý Cột trái Header
            if form_type == "Khối Đảng":
                p_l1 = cell_left.paragraphs[0]
                p_l1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l1.paragraph_format.line_spacing = 1.15
                p_l1.paragraph_format.space_after = Pt(0)
                r = p_l1.add_run("ĐẢNG BỘ THÀNH PHỐ ĐỒNG NAI")
                r.font.name, r.font.size = 'Times New Roman', Pt(12)

                p_l2 = cell_left.add_paragraph()
                p_l2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l2.paragraph_format.line_spacing = 1.15
                p_l2.paragraph_format.space_after = Pt(0)
                r = p_l2.add_run(agency_name.upper())
                r.font.name, r.font.size, r.font.bold = 'Times New Roman', Pt(12), True

                p_l3 = cell_left.add_paragraph()
                p_l3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l3.paragraph_format.line_spacing = 1.15
                p_l3.paragraph_format.space_after = Pt(2)
                r = p_l3.add_run("*")
                r.font.name, r.font.size = 'Times New Roman', Pt(9)

                p_l4 = cell_left.add_paragraph()
                p_l4.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l4.paragraph_format.line_spacing = 1.15
                p_l4.paragraph_format.space_after = Pt(0)
                r = p_l4.add_run(f"Số:       {code_str}")
                r.font.name, r.font.size = 'Times New Roman', Pt(12)

                if cv_subject:
                    p_l5 = cell_left.add_paragraph()
                    p_l5.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_l5.paragraph_format.space_before = Pt(6)
                    r = p_l5.add_run(cv_subject)
                    r.font.name, r.font.size, r.font.italic = 'Times New Roman', Pt(11), True
            else:
                p_l1 = cell_left.paragraphs[0]
                p_l1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l1.paragraph_format.line_spacing = 1.15
                p_l1.paragraph_format.space_after = Pt(0)
                r = p_l1.add_run("UBND THÀNH PHỐ ĐỒNG NAI")
                r.font.name, r.font.size = 'Times New Roman', Pt(12)

                p_l2 = cell_left.add_paragraph()
                p_l2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l2.paragraph_format.line_spacing = 1.15
                p_l2.paragraph_format.space_after = Pt(0)
                r = p_l2.add_run(agency_name.upper())
                r.font.name, r.font.size, r.font.bold = 'Times New Roman', Pt(12), True
                add_paragraph_bottom_border(p_l2, sz="6", space="4")

                p_l3 = cell_left.add_paragraph()
                p_l3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l3.paragraph_format.line_spacing = 1.15
                p_l3.paragraph_format.space_after = Pt(2)
                r = p_l3.add_run("*")
                r.font.name, r.font.size = 'Times New Roman', Pt(9)

                p_l4 = cell_left.add_paragraph()
                p_l4.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l4.paragraph_format.line_spacing = 1.15
                p_l4.paragraph_format.space_after = Pt(0)
                r = p_l4.add_run(f"Số:       {code_str}")
                r.font.name, r.font.size = 'Times New Roman', Pt(12)

                if cv_subject:
                    p_l5 = cell_left.add_paragraph()
                    p_l5.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_l5.paragraph_format.space_before = Pt(6)
                    r = p_l5.add_run(cv_subject)
                    r.font.name, r.font.size, r.font.italic = 'Times New Roman', Pt(11), True

            # Xử lý Cột phải Header (Dùng viền gạch tách xa chân chữ 5pt)
            if form_type == "Khối Đảng":
                p_r1 = cell_right.paragraphs[0]
                p_r1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_r1.paragraph_format.line_spacing = 1.15
                p_r1.paragraph_format.space_after = Pt(6)
                r = p_r1.add_run("ĐẢNG CỘNG SẢN VIỆT NAM")
                r.font.name, r.font.size, r.font.bold = 'Times New Roman', Pt(12), True
                add_paragraph_bottom_border(p_r1, sz="6", space="5")

                p_r2 = cell_right.add_paragraph()
                p_r2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_r2.paragraph_format.line_spacing = 1.15
                p_r2.paragraph_format.space_after = Pt(0)
                r = p_r2.add_run("Nhơn Trạch, ngày     tháng 8 năm 2026")
                r.font.name, r.font.size, r.font.italic = 'Times New Roman', Pt(12), True
            else:
                p_r1 = cell_right.paragraphs[0]
                p_r1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_r1.paragraph_format.line_spacing = 1.15
                p_r1.paragraph_format.space_after = Pt(0)
                r = p_r1.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
                r.font.name, r.font.size, r.font.bold = 'Times New Roman', Pt(12), True

                p_r2 = cell_right.add_paragraph()
                p_r2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_r2.paragraph_format.line_spacing = 1.15
                p_r2.paragraph_format.space_after = Pt(6)
                r = p_r2.add_run("Độc lập - Tự do - Hạnh phúc")
                r.font.name, r.font.size, r.font.bold = 'Times New Roman', Pt(12.5), True
                add_paragraph_bottom_border(p_r2, sz="6", space="5")

                p_r3 = cell_right.add_paragraph()
                p_r3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_r3.paragraph_format.line_spacing = 1.15
                p_r3.paragraph_format.space_after = Pt(0)
                r = p_r3.add_run("Nhơn Trạch, ngày     tháng 8 năm 2026")
                r.font.name, r.font.size, r.font.italic = 'Times New Roman', Pt(12), True

            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_after = Pt(6)

            next_is_trich_yeu = False
            for line in lines_data:
                if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', line):
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(4)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name, run.font.size, run.font.bold = 'Times New Roman', Pt(13), True
                    next_is_trich_yeu = False
                elif line.isupper() and doc_type_str != "Công văn":
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(4)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(line)
                    run.font.name, run.font.size, run.font.bold = 'Times New Roman', Pt(14), True
                    next_is_trich_yeu = True
                elif next_is_trich_yeu and doc_type_str != "Công văn":
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(2)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(line)
                    run.font.name, run.font.size, run.font.bold = 'Times New Roman', Pt(13), True
                    
                    # Tạo đường kẻ viền mảnh chuẩn 1/3 ngắn bên dưới Trích yếu
                    p_line = doc.add_paragraph()
                    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_line.paragraph_format.space_before = Pt(2)
                    p_line.paragraph_format.space_after = Pt(8)
                    p_line.paragraph_format.left_indent = Cm(5.5)
                    p_line.paragraph_format.right_indent = Cm(5.5)
                    add_paragraph_bottom_border(p_line, sz="6", space="1")
                    next_is_trich_yeu = False
                elif re.match(r'^\d+\.', line) and len(line) < 80:
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(4)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name, run.font.size, run.font.bold = 'Times New Roman', Pt(13), True
                    next_is_trich_yeu = False
                elif line.startswith("Kính gửi:") or line.startswith("-"):
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(4)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name, run.font.size = 'Times New Roman', Pt(12 if line.startswith("-") else 13)
                    next_is_trich_yeu = False
                else:
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(4)
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.first_line_indent = Cm(1.27)
                    run = p.add_run(line)
                    run.font.name, run.font.size = 'Times New Roman', Pt(13)
                    next_is_trich_yeu = False

            bio = io.BytesIO()
            doc.save(bio)
            return bio.getvalue()

        st.download_button(
            label="📥 TẢI VỀ FILE WORD (.DOCX)",
            data=generate_docx(filtered_lines, co_quan, the_thuc, loai_vb, trich_yeu_cv),
            file_name="Du_Thao_Van_Ban_Hanh_Chinh.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

    # CỘT BÊN PHẢI: CHAT AI SỬA ĐỔI
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
