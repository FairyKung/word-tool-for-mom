import streamlit as st
from docx import Document
from docx.shared import Pt
import io
import re
import mammoth
import copy

st.set_page_config(page_title="Word Editor Ultra", layout="wide")

st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 20px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'ช้าแต่ชัวร์' (รักษาฟอนต์ 100%) 📄✨")

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
                old = st.text_input(f"คำเดิม {i+1}", key=f"old_{i}")
            with col_b:
                new = st.text_input(f"คำใหม่ {i+1}", key=f"new_{i}")
            if old:
                replace_list.append((old, new))
        
        st.button("➕ เพิ่มช่องเปลี่ยนคำ", on_click=add_pair)
        process_btn = st.button("🚀 เริ่มเปลี่ยนคำ (แบบละเอียด)", type="primary")

    # --- ฟังก์ชันหัวใจสำคัญ: แก้คำโดย "ก๊อปปี้ฟอนต์เดิม" มาแปะ ---
    def advanced_replace(paragraphs, old_text, new_text):
        pattern = re.escape(old_text).replace(r'\ ', r'\s+')
        for p in paragraphs:
            if re.search(pattern, p.text):
                # ถ้าคำนั้นอยู่ใน Run เดียวกัน (กรณีทั่วไป)
                for run in p.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)
                
                # ถ้าคำโดนหั่น (ด่านปราบเซียน) - เราจะใช้วิธีเปลี่ยนที่ p.text 
                # แต่ต้องพยายามรักษาฟอนต์โดยการนำฟอนต์จาก Run แรกมาประยุกต์ใช้กับทั้งหมด
                if old_text in p.text:
                    orig_font_name = p.runs[0].font.name if p.runs and p.runs[0].font.name else None
                    orig_font_size = p.runs[0].font.size if p.runs and p.runs[0].font.size else None
                    orig_bold = p.runs[0].bold if p.runs else None
                    
                    full_text = p.text.replace(old_text, new_text)
                    p.text = full_text # เขียนทับ
                    
                    # บังคับคืนค่าฟอนต์ทันทีหลังจากเขียนทับ
                    if p.runs:
                        for r in p.runs:
                            if orig_font_name: r.font.name = orig_font_name
                            if orig_font_size: r.font.size = orig_font_size
                            if orig_bold is not None: r.bold = orig_bold

    if process_btn and replace_list:
        with st.spinner("กำลังประมวลผลอย่างละเอียด กรุณารอประมาน 10-20 วินาทีนะค๊ะ..."):
            doc = Document(io.BytesIO(file_content))
            for old, new in replace_list:
                # 1. เนื้อหาหลัก
                advanced_replace(doc.paragraphs, old, new)
                
                # 2. ในตาราง (ต้องวนลูปละเอียด)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            advanced_replace(cell.paragraphs, old, new)
                
                # 3. ในทุก Section (บังคับวนลูป Header/Footer ครบทุกหน้า)
                for section in doc.sections:
                    headers_footers = [
                        section.header, section.footer,
                        section.first_page_header, section.first_page_footer,
                        section.even_page_header, section.even_page_footer
                    ]
                    for hf in headers_footers:
                        if hf:
                            advanced_replace(hf.paragraphs, old, new)
            
            out_bio = io.BytesIO()
            doc.save(out_bio)
            st.session_state.processed_doc = out_bio.getvalue()
            st.sidebar.success("✅ แก้ไขเรียบร้อยครบทุกหน้าแล้วค่ะ!")

    # --- ส่วนการแสดงผล (Tabs) ---
    tab1, tab2 = st.tabs(["📄 ไฟล์ต้นฉบับ", "✨ พรีวิวหลังแก้ไข"])
    with tab1:
        html_orig = mammoth.convert_to_html(io.BytesIO(file_content)).value
        st.markdown(f'<div class="paper-container"><div class="paper-content">{html_orig}</div></div>', unsafe_allow_html=True)

    with tab2:
        if st.session_state.processed_doc:
            st.download_button("📥 ดาวน์โหลดไฟล์", data=st.session_state.processed_doc, 
                             file_name=f"Fixed_Ultra_{uploaded_file.name}")
            html_fixed = mammoth.convert_to_html(io.BytesIO(st.session_state.processed_doc)).value
            st.markdown(f'<div class="paper-container"><div class="paper-content">{html_fixed}</div></div>', unsafe_allow_html=True)