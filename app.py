import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Final", layout="wide")

st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 20px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'จัดการวันที่' 📄✨")

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
        process_btn = st.button("🚀 เริ่มเปลี่ยนคำและพรีวิว", type="primary")

    # --- ฟังก์ชันแก้คำแบบแก้ปัญหา "คำโดนหั่น" ---
    def master_replace(paragraphs, old_text, new_text):
        # สร้างระบบค้นหาแบบไม่สนเว้นวรรค (หัวใจของการแก้วันที่)
        pattern = re.escape(old_text).replace(r'\ ', r'\s+')
        for p in paragraphs:
            if re.search(pattern, p.text):
                # เก็บฟอนต์ของตัวแรกไว้ก่อนเผื่อกันเพี้ยน
                inline_tags = []
                if p.runs:
                    font_name = p.runs[0].font.name
                    font_size = p.runs[0].font.size
                
                # เปลี่ยนคำในระดับ Paragraph Text (แก้ปัญหาเรื่องโดนหั่น Run)
                p.text = re.sub(pattern, new_text, p.text)
                
                # คืนค่าฟอนต์พื้นฐานให้ (เพื่อไม่ให้ขนาดเปลี่ยนจนน่าเกลียด)
                if p.runs and font_name:
                    for run in p.runs:
                        run.font.name = font_name
                        run.font.size = font_size

    if process_btn and replace_list:
        doc = Document(io.BytesIO(file_content))
        for old, new in replace_list:
            # 1. เนื้อหาหลัก
            master_replace(doc.paragraphs, old, new)
            # 2. ในตาราง
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        master_replace(cell.paragraphs, old, new)
            # 3. ในทุก Section (Header/Footer ทุกหน้า รวมถึงหน้าแรก)
            for section in doc.sections:
                for hf in [section.header, section.footer, 
                           section.first_page_header, section.first_page_footer,
                           section.even_page_header, section.even_page_footer]:
                    if hf:
                        master_replace(hf.paragraphs, old, new)
        
        out_bio = io.BytesIO()
        doc.save(out_bio)
        st.session_state.processed_doc = out_bio.getvalue()
        st.sidebar.success("✅ แก้ไขเรียบร้อย! ลองเช็ควัดพรีวิวนะค๊ะ")

    # --- ส่วนการแสดงผล (Tabs) ---
    tab1, tab2 = st.tabs(["📄 ไฟล์ต้นฉบับ", "✨ พรีวิวหลังแก้ไข"])
    
    with tab1:
        html_orig = mammoth.convert_to_html(io.BytesIO(file_content)).value
        st.markdown(f'<div class="paper-container"><div class="paper-content">{html_orig}</div></div>', unsafe_allow_html=True)

    with tab2:
        if st.session_state.processed_doc:
            st.download_button("📥 ดาวน์โหลดไฟล์", data=st.session_state.processed_doc, 
                             file_name=f"Fixed_{uploaded_file.name}")
            html_fixed = mammoth.convert_to_html(io.BytesIO(st.session_state.processed_doc)).value
            st.markdown(f'<div class="paper-container"><div class="paper-content">{html_fixed}</div></div>', unsafe_allow_html=True)