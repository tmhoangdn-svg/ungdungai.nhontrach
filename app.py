import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import requests
import json
import io
import re
import os

try:
    import openpyxl
except ImportError:
    openpyxl = None

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

st.set_page_config(
    page_title="Phần mềm Cụ thể hóa Văn bản Hành chính",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    div[data-testid="stVerticalBlock"] > div {
        gap: 0.4rem !important;
    }
    .app-header {
        background: linear-gradient(135deg, #7b0000 0%, #a81010 50%, #c41e1e 100%);
        border: 1px solid #e0a800;
        border-radius: 12px;
        padding: 18px 25px;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.4);
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .app-header-title {
        color: #ffffff;
        font-size: 22px;
        font-weight: 800;
        letter-spacing: 0.5px;
        text-shadow: 1px 1px 3px rgba(0,0,0,0.6);
        margin: 0;
    }
    .app-header-sub {
        color: #ffd700;
        font-size: 12.5px;
        margin-top: 4px;
        font-weight: 500;
        letter-spacing: 0.3px;
    }
    .section-badge {
        background: linear-gradient(90deg, #d4af37 0%, #f3e5ab 100%);
        color: #4a2c00;
        font-size: 11px;
        font-weight: bold;
        padding: 3px 8px;
        border-radius: 4px;
        text-transform: uppercase;
        display: inline-block;
        margin-bottom: 6px;
    }
    .word-page {
        background-color: #ffffff;
        color: #000000;
        width: 100%;
        max-height: 480px;
        overflow-y: auto;
        padding: 25px 35px;
        border: 1px solid #d3d3d3;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
        font-family: 'Times New Roman', Times, serif;
        font-size: 13pt;
        line-height: 1.0;
        word-wrap: break-word;
        border-radius: 4px;
    }
    .word-page table {
        width: 100% !important;
        border-collapse: collapse !important;
        border: none !important;
        margin-bottom: 8px;
    }
    .word-page td {
        border: none !important;
        vertical-align: top;
        padding: 0px 4px;
        font-family: 'Times New Roman', Times, serif;
    }
    .word-page p {
        font-family: 'Times New Roman', Times, serif;
        font-size: 13pt;
        line-height: 1.0;
        margin-top: 4pt;
        margin-bottom: 4pt;
        text-align: justify;
        text-indent: 0.98cm;
    }
    .word-page p.kinh-gui {
        text-indent: 0 !important;
        margin-left: 0cm !important;
        text-align: center !important;
        margin-top: 4pt;
        margin-bottom: 4pt;
    }
    .word-page p.noi-nhan {
        text-indent: 0cm !important;
        margin-left: 1.0cm !important;
        text-align: left !important;
        margin-top: 2pt;
        margin-bottom: 2pt;
        font-size: 11pt !important;
    }
    .word-page p.title-bold {
        text-indent: 0 !important;
        text-align: center !important;
        font-weight: bold;
        margin-top: 4pt;
        margin-bottom: 4pt;
    }
    .custom-underline {
        text-decoration: underline;
    }
</style>
""", unsafe_allow_html=True)

if "current_draft" not in st.session_state:
    st.session_state.current_draft = ""
if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []

BUILTIN_TEMPLATES_DANG = {
    "📌 Mẫu Công văn chuẩn (HD 05-HD/VPTW)": """KÍNH GỬI: Các chi, đảng bộ trực thuộc.
1. Căn cứ ban hành...
2. Nội dung chỉ đạo, triển khai...
3. Tổ chức thực hiện và báo cáo kết quả...""",

    "📌 Đề cương Kế hoạch chuẩn (HD 05-HD/VPTW)": """KẾ HOẠCH Về việc...
I. MỤC ĐÍCH, YÊU CẦU
1. Mục đích
2. Yêu cầu
II. NỘI DUNG VÀ NHIỆM VỤ CỤ THỂ
1. Nhiệm vụ trọng tâm
2. Giải pháp thực hiện
III. TỔ CHỨC THỰC HIỆN
1. Phân công trách nhiệm
2. Tiến độ và thời gian hoàn thành""",

    "📌 Đề cương Báo cáo chuẩn (HD 05-HD/VPTW)": """BÁO CÁO Tình hình...
I. KẾT QUẢ ĐẠT ĐƯỢC
1. Công tác chỉ đạo, quán triệt
2. Kết quả thực hiện các nhiệm vụ
II. HẠN CHẾ, KHUYẾT ĐIỂM VÀ NGUYÊN NHÂN
1. Hạn chế, tồn tại
2. Nguyên nhân (chủ quan, khách quan)
III. PHƯƠNG HƯỚNG, NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI""",

    "📌 Mẫu Giấy mời chuẩn (HD 05-HD/VPTW)": """GIẤY MỜI Về việc...
Đảng ủy phường Nhơn Trạch trân trọng kính mời: ...
- Thời gian: Vào lúc ... giờ ..., ngày ... tháng ... năm 2026.
- Địa điểm: Hội trường Đảng ủy phường Nhơn Trạch.
- Chủ trì: Đồng chí Bí thư Đảng ủy phường.
- Nội dung: ...""",

    "📌 Mẫu Tờ trình chuẩn (HD 05-HD/VPTW)": """TỜ TRÌNH Về việc...
KÍNH GỬI: Ban Thường vụ / Cơ quan cấp trên.
I. SỰ CẦN THIẾT / CĂN CỨ TRÌNH
II. NỘI DUNG CHÍNH CỦA TỜ TRÌNH
III. ĐỀ XUẤT, KIẾN NGHỊ"""
}

BUILTIN_TEMPLATES_NN = {
    "📌 Mẫu Công văn chuẩn (NĐ 30/2020/NĐ-CP)": """KÍNH GỬI: Các phòng, ban, đơn vị trực thuộc.
1. Căn cứ thực hiện...
2. Nội dung giao nhiệm vụ/thông báo...
3. Yêu cầu báo cáo/thời hạn hoàn thành...""",

    "📌 Đề cương Kế hoạch chuẩn (NĐ 30/2020/NĐ-CP)": """KẾ HOẠCH Về việc...
I. MỤC ĐÍCH, YÊU CẦU
II. NỘI DUNG VÀ CHỈ TIÊU NHIỆM VỤ
III. TỔ CHỨC THỰC HIỆN VÀ KINH PHÍ""",

    "📌 Đề cương Báo cáo chuẩn (NĐ 30/2020/NĐ-CP)": """BÁO CÁO Kết quả thực hiện...
I. TÌNH HÌNH VÀ KẾT QUẢ THỰC HIỆN
II. ĐÁNH GIÁ CHUNG (Ưu điểm, Hạn chế, Nguyên nhân)
III. NHIỆM VỤ GIẢI PHÁP VÀ ĐỀ XUẤT, KIẾN NGHỊ""",

    "📌 Mẫu Giấy mời chuẩn (NĐ 30/2020/NĐ-CP)": """GIẤY MỜI Về việc...
Ủy ban nhân dân phường Nhơn Trạch trân trọng kính mời: ...
- Thời gian: Vào lúc ... giờ ..., ngày ... tháng ... năm 2026.
- Địa điểm: Phòng họp UBND phường.
- Chủ trì: Đồng chí Chủ tịch UBND phường.
- Nội dung: ...""",

    "📌 Mẫu Tờ trình chuẩn (NĐ 30/2020/NĐ-CP)": """TỜ TRÌNH Về việc...
KÍNH GỬI: Ủy ban nhân dân cấp trên / Cơ quan có thẩm quyền.
I. CĂN CỨ PHÁP LÝ VÀ SỰ CẦN THIẾT
II. NỘI DUNG ĐỀ XUẤT
III. DỰ THẢO NGHỊ QUYẾT/QUYẾT ĐỊNH KÈM THEO"""
}

def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return ""
    filename = uploaded_file.name.lower()
    text = ""
    try:
        if filename.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8", errors="ignore")
        elif filename.endswith(".docx"):
            doc = Document(uploaded_file)
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif filename.endswith(".doc"):
            raw = uploaded_file.read().decode("latin-1", errors="ignore")
            text = "".join([c for c in raw if c.isprintable() or c in ["\n", "\r", "\t"]])
            text = re.sub(r"\s+", " ", text)
        elif filename.endswith((".xlsx", ".xls")):
            if openpyxl and filename.endswith(".xlsx"):
                wb = openpyxl.load_workbook(uploaded_file, data_only=True)
                sheets_text = []
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        row_vals = [str(cell) for cell in row if cell is not None]
                        if row_vals:
                            sheets_text.append("\t".join(row_vals))
                text = "\n".join(sheets_text)
            else:
                text = f"[Tệp bảng tính {uploaded_file.name} đã tải lên]"
        elif filename.endswith(".pdf"):
            try:
                import pypdf
                reader = pypdf.PdfReader(uploaded_file)
                for page in reader.pages:
                    extracted = page.extract_text()
                    if extracted:
                        text += extracted + "\n"
            except Exception:
                text = f"[Tệp PDF {uploaded_file.name} đã tải lên]"
    except Exception as e:
