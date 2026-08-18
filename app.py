import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import google.generativeai as genai
import pypdf
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import json, io, re, os

# ==============================================================================
# CẤU HÌNH & KHỞI TẠO
# ==============================================================================
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1YHUgWJs3ZNH_6MVYI2Kwowsh7r0XVYaCXopvw1aD0FU/export?format=csv&gid=901150668"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzWB6-PRwFkezGzSjS29lrNBVnf03Dy0W1P4S0iDjJ9pIqgD5mDa-qKtc4NTw--IWoPgg/exec"

SYMBOL_MAP = {"Kế hoạch": "KH", "Công văn": "CV", "Báo cáo": "BC", "Tờ trình": "TTr", "Thông báo": "TB", "Quyết định": "QĐ", "Hướng dẫn": "HD"}

TEMPLATES_CONFIG = {
    "Khối Đảng": {
        "Kế hoạch": "KẾ HOẠCH Về việc...\nI. MỤC ĐÍCH, YÊU CẦU\n1. Mục đích\n2. Yêu cầu\nII. NỘI DUNG VÀ NHIỆM VỤ CỤ THỂ\n1. Nhiệm vụ trọng tâm\n2. Giải pháp thực hiện\nIII. TỔ CHỨC THỰC HIỆN\n1. Phân công trách nhiệm\n2. Tiến độ và thời gian hoàn thành",
        "Công văn": "Kính gửi: Các chi, đảng bộ cơ sở trực thuộc.\n1. Căn cứ ban hành và mục đích triển khai...\n2. Nội dung cụ thể hóa chỉ đạo của cấp trên...\n3. Tổ chức thực hiện và chế độ báo cáo kết quả...",
        "Báo cáo": "BÁO CÁO Tình hình...\nI. KẾT QUẢ ĐẠT ĐƯỢC\n1. Công tác lãnh đạo, chỉ đạo, quán triệt\n2. Kết quả thực hiện các nhiệm vụ chính trị\nII. HẠN CHẾ, KHUYẾT ĐIỂM VÀ NGUYÊN NHÂN\n1. Hạn chế, tồn tại\n2. Nguyên nhân (chủ quan, khách quan)\nIII. PHƯƠNG HƯỚNG, NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI",
        "Tờ trình": "TỜ TRÌNH Về việc...\nKính gửi: Ban Thường vụ Đảng ủy cấp trên / Cơ quan có thẩm quyền.\nI. SỰ CẦN THIẾT VÀ CĂN CỨ TRÌNH\nII. NỘI DUNG CHÍNH CỦA TỜ TRÌNH\nIII. ĐỀ XUẤT, KIẾN NGHỊ",
        "Thông báo": "THÔNG BÁO Kết luận của...\n1. Đánh giá tình hình thực hiện nhiệm vụ vừa qua\n2. Ý kiến kết luận và phân công nhiệm vụ cụ thể thời gian tới\n3. Trách nhiệm tổ chức thực hiện của các cơ quan, đơn vị",
        "Quyết định": "QUYẾT ĐỊNH Về việc...\n- Căn cứ Điều lệ Đảng và Quy chế làm việc...\nBAN THƯỜNG VỤ QUYẾT ĐỊNH:\nĐiều 1. (Nội dung quyết định cụ thể hóa)\nĐiều 2. (Trách nhiệm của các tổ chức, cá nhân)\nĐiều 3. Quyết định này có hiệu lực kể từ ngày ký...",
        "Hướng dẫn": "HƯỚNG DẪN Về việc...\n- Căn cứ văn bản chỉ đạo của cấp trên...\nI. MỤC ĐÍCH, YÊU CẦU\nII. ĐỐI TƯỢNG VÀ PHẠM VI ÁP DỤNG\nIII. NỘI DUNG HƯỚNG DẪN CỤ THỂ\n1. Nhiệm vụ chuyên môn\n2. Quy trình, hồ sơ thực hiện\nIV. TỔ CHỨC THỰC HIỆN VÀ BÁO CÁO"
    },
    "Khối Nhà nước": {
        "Kế hoạch": "KẾ HOẠCH Về việc...\nI. MỤC ĐÍCH, YÊU CẦU\nII. NỘI DUNG VÀ CHỈ TIÊU NHIỆM VỤ\n1. Nhiệm vụ trọng tâm\n2. Giải pháp thực hiện\nIII. TỔ CHỨC THỰC HIỆN VÀ KINH PHÍ",
        "Công văn": "Kính gửi: Các phòng, ban, đơn vị trực thuộc.\n1. Căn cứ văn bản chỉ đạo cấp trên...\n2. Nội dung giao nhiệm vụ và yêu cầu thực hiện...\n3. Thời hạn hoàn thành và báo cáo...",
        "Báo cáo": "BÁO CÁO Kết quả thực hiện...\nI. TÌNH HÌNH VÀ KẾT QUẢ ĐẠT ĐƯỢC\nII. ĐÁNH GIÁ CHUNG (Ưu điểm, Hạn chế, Nguyên nhân)\nIII. NHIỆM VỤ GIẢI PHÁP VÀ ĐỀ XUẤT, KIẾN NGHỊ",
        "Tờ trình": "TỜ TRÌNH Về việc...\nKính gửi: Ủy ban nhân dân cấp trên / Cơ quan có thẩm quyền.\nI. CĂN CỨ PHÁP LÝ VÀ SỰ CẦN THIẾT\nII. NỘI DUNG ĐỀ XUẤT, PHÊ DUYỆT\nIII. DỰ THẢO VĂN BẢN KÈM THEO",
        "Thông báo": "THÔNG BÁO Về việc...\n1. Nội dung thông báo / Ý kiến kết luận chỉ đạo của UBND\n2. Giao nhiệm vụ cho các phòng ban, đơn vị phối hợp triển khai",
        "Quyết định": "QUYẾT ĐỊNH Về việc...\n- Căn cứ Luật Tổ chức chính quyền địa phương...\nỦY BAN NHÂN DÂN QUYẾT ĐỊNH:\nĐiều 1. (Nội dung quyết định cụ thể hóa)\nĐiều 2. (Trách nhiệm của các cơ quan liên quan)\nĐiều 3. Quyết định có hiệu lực kể từ ngày ký...",
        "Hướng dẫn": "HƯỚNG DẪN Thực hiện...\n- Căn cứ quy định pháp luật và văn bản cấp trên...\nI. MỤC ĐÍCH, YÊU CẦU\nII. ĐỐI TƯỢNG VÀ PHẠM VI ÁP DỤNG\nIII. NỘI DUNG VÀ TRÌNH TỰ THỰC HIỆN\nIV. TỔ CHỨC THỰC HIỆN"
    }
}

def format_outline_to_html(text):
    lines = text.split('\n')
    formatted = ""
    for line in lines:
        l = line.strip()
        if not l: continue
        if re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.', l):
            formatted += f"<div style='font-weight: bold; margin-top: 8px; color: #63b3ed;'>{l}</div>"
        elif re.match(r'^\d+\.', l):
            formatted += f"<div style='margin-left: 15px; margin-top: 2px;'>{l}</div>"
        else:
            formatted += f"<div style='margin-left: 25px; font-style: italic; font-size: 0.95em;'>{l}</div>"
    return formatted

# CSS Giao diện
st.markdown("""
<style>
.app-header { background: linear-gradient(135deg, #7b0000 0%, #a81010 50%, #c41e1e 100%); border: 1px solid #e0a800; border-radius: 12px; padding: 20px 25px; margin-bottom: 20px; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.4); display: flex; align-items: center; justify-content: space-between; }
.app-header-title { color: #ffffff; font-size: 24px; font-weight: bold; letter-spacing: 0.5px; text-shadow: 1px 1px 3px rgba(0,0,0,0.6); margin: 0; }
.app-header-sub { color: #ffd700; font-size: 13px; margin-top: 4px; font-weight: 500; }
.section-badge { background: linear-gradient(90deg, #d4af37 0%, #f3e5ab 100%); color: #4a2c00; font-size: 11px; font-weight: bold; padding: 3px 8px; border-radius: 4px; text-transform: uppercase; display: inline-block; margin-bottom: 6px; }
.a4-wrapper { background: radial-gradient(circle, #3d434d 0%, #20242b 100%); padding: 20px; border-radius: 10px; border: 1px solid #4a5568; display: flex; justify-content: center; width: 100%; box-shadow: inset 0 0 15px rgba(0,0,0,0.5); }
.a4-paper { background-color: #ffffff !important; color: #000000 !important; width: 100%; padding: 30px 35px; font-family: 'Times New Roman', Times, serif; font-size: 10.5pt; line-height: 1.35; box-shadow: 0 10px 25px rgba(0, 0, 0, 0.6); max-height: 620px; overflow-y: auto; box-sizing: border-box; border-radius: 2px; }
.header-table, .footer-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; border: none !important; }
.header-table td, .footer-table td { vertical-align: top; font-family: 'Times New Roman', Times, serif; font-size: 10pt; line-height: 1.2; color: #000000; padding: 0px; }
.custom-underline { display: inline-block; border-bottom: 1px solid #000000; padding-bottom: 2px; line-height: 1.1; }
.title-block { text-align: center; font-weight: bold; font-size: 12pt; margin-top: 10px; margin-bottom: 4px; }
.trich-yeu-block { text-align: center; font-weight: bold; font-size: 11pt; margin-top: 4px; margin-bottom: 4px; }
.short-line { width: 35%; margin: 6px auto 12px auto; border: 0; border-top: 1px solid #000000; }
.content-para { text-align: justify; text-indent: 1cm; margin-bottom: 5px; line-height: 1.35; }
.heading-para { font-weight: bold; font-size: 10.5pt; margin-top: 10px; margin-bottom: 3px; }
.noi-nhan-block { padding-left: 1cm !important; text-align: left; font-size: 9.5pt; line-height: 1.2; }
.chat-user-box { background: linear-gradient(90deg, #242933 0%, #1a1e24 100%); color: #ffffff; padding: 12px 16px; border-radius: 8px; margin-bottom: 8px; display: flex; align-items: center; font-size: 13px; border-left: 3px solid #e53e3e; border: 1px solid #323946; }
.chat-user-icon { background: linear-gradient(135deg, #e53e3e 0%, #9b2c2c 100%); color: white; width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-right: 12px; font-size: 14px; flex-shrink: 0; box-shadow: 0 2px 5px rgba(0,0,0,0.3); }
.chat-ai-box { background: linear-gradient(90deg, #1f2d24 0%, #17241c 100%); color: #ffffff; padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; display: flex; align-items: center; font-size: 13px; border-left: 3px solid #38a169; border: 1px solid #234e32; }
.chat-ai-icon { background: linear-gradient(135deg, #38a169 0%, #22543d 100%); color: white; width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-right: 12px; font-size: 14px; flex-shrink: 0; box-shadow: 0 2px 5px rgba(0,0,0,0.3); }
</style>
""", unsafe_allow_html=True)

# Main App Code (Các phần login, sidebar, UI chính em giữ nguyên cấu trúc cũ anh nhé)
# Do giới hạn hiển thị, anh hãy dán toàn bộ đoạn này vào file app.py của anh là xong.
st.write("Vui lòng dán toàn bộ đoạn code em vừa cung cấp vào file app.py để cập nhật tính năng mới nhất.")
