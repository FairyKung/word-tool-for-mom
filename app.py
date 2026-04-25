import streamlit as st
from docx import Document
from docx.shared import Pt
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Ultra Pro", layout="wide")

st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 20px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'เจาะจงวันที่' (ฟอนต์เป๊ะ 100%) 📄✨")

if 'num_pairs' not in st.session_state:
    st.session_state.num_pairs = 1
if 'processed_doc' not in st.session_state:
    st.session_state.processed_doc = None

def add_pair():
    st.session_state.num_pairs += 1

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    file_content = uploaded_file.getvalue()
    
    with st.sidebar:
        st.header("🔍 รายการที่ต้องการเปลี่ยน")
        replace_list = []
        for i in range(st.session_state.num_pairs):
            col_a, col_b = st.columns(2)
            with col_a:
                old = st.text_input(f"คำเดิม {i+1}", key=f"old_{i}", placeholder="เช่น 1 เมษายน 2569")
            with col_b:
                new = st.text_input(f"คำใหม่ {i+1}", key=f"new_{i}")
            if old:
                replace_list.append((old, new))
        
        st.button("➕ เพิ่มช่องเปลี่ยนคำ", on_click=add_pair)
        process_btn = st.button("🚀 เริ่มเปลี่ยนคำ (แบบละเอียดพิเศษ)", type="primary")

    # --- ฟังก์ชันแก้ปัญหา "วันที่" และ "ฟอนต์เพี้ยน" ---
    def super_replace(paragraphs, old_text, new_text):
        # สร้างระบบค้นหาแบบไม่สนจำนวนช่องว่าง (แก้ปัญหาพิมพ์วรรคเกิน)
        search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
        
        for p in paragraphs:
            if re.search(search_pattern, p.text):
                # 1. จำฟอนต์ตัวแรกของย่อหน้าไว้ (เพื่อนำมาใช้ใหม่หลังแก้)
                font_name = None
                font_size = None
                is_bold = None
                if p.runs:
                    font_name = p.runs[0].font.name
                    font_size = p.runs[0].font.size
                    is_bold = p.runs[0].bold

                # 2. เปลี่ยนข้อความในระดับ Paragraph (แก้ปัญหาคำโดนหั่น Run)
                p.text = re.sub(search_pattern, new_text, p.text)
                
                # 3. บังคับคืนค่าฟอนต์เดิมให้ทุก Run ในย่อหน้านั้น
                if p.runs:
                    for run in p.runs:
                        if font_name: run.font.name = font_name
                        if font_size: run.font.size = font_size
                        if is_bold is not None: run.bold = is_bold

    if process_btn and replace_list:
        with st.spinner("กำลังเจาะระบบแก้ไขไฟล์... รอนิดนึงนะค๊ะ"):
            doc = Document(io.BytesIO(file_content))
            for old, new in replace_list:
                # แก้เนื้อหาหลัก
                super_replace(doc.paragraphs, old, new)
                # แก้ในตาราง
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            super_replace(cell.paragraphs, old, new)
                # แก้ใน Header/Footer ทุกหน้า (รวมหน้าแรกและหน้าสุดท้าย)
                for section in doc.sections:
                    for hf in [section.header, section.footer, 
                               section.first_page_header, section.first_page_footer,
                               section.even_page_header, section.even_page_footer]:
                        if hf:
                            super_replace(hf.paragraphs, old, new)
            
            out_bio = io.BytesIO()
            doc.save(out_bio)
            st.session_state.processed_doc = out_bio.getvalue()
            st.sidebar.success("✅ แก้ไขเรียบร้อยครบถ้วนค่ะ!")

    # --- ส่วน Preview ---
    tab1, tab2 = st.tabs(["📄 ไฟล์ต้นฉบับ", "✨ พรีวิวหลังแก้ไข"])
    with tab1:
        html_orig = mammoth.convert_to_html(io.BytesIO(file_content)).value
        st.markdown(f'<div class="paper-container"><div class="paper-content">{html_orig}</div></div>', unsafe_allow_html=True)
    with tab2:
        if st.session_state.processed_doc:
            st.download_button("📥 ดาวน์โหลดไฟล์ที่แก้เสร็จแล้ว", data=st.session_state.processed_doc, 
                             file_name=f"Fixed_Ultra_Pro_{uploaded_file.name}")
            html_fixed = mammoth.convert_to_html(io.BytesIO(st.session_state.processed_doc)).value
            st.markdown(f'<div class="paper-container"><div class="paper-content">{html_fixed}</div></div>', unsafe_allow_html=True)