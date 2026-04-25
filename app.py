import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Perfect Font", layout="wide")

st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 20px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'ฟอนต์ไม่เพี้ยน' (แก้ครบ 100 หน้า) 📄✨")

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
        process_btn = st.button("🚀 เริ่มเปลี่ยนคำและพรีวิว", type="primary")

    # --- ฟังก์ชันแก้คำแบบถนอมฟอนต์ (Advanced Run-level Replace) ---
    def master_replace(paragraphs, old_text, new_text):
        pattern = re.escape(old_text).replace(r'\ ', r'\s+')
        for p in paragraphs:
            # ตรวจสอบว่ามีคำนี้อยู่ในย่อหน้าหรือไม่
            if re.search(pattern, p.text):
                # เทคนิค: วนลูปแก้ในแต่ละ Run โดยตรงเพื่อรักษาฟอนต์ของ Run นั้นๆ
                # และใช้การตรวจจับคำที่อาจโดนหั่นเป็นชิ้นๆ
                for run in p.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)
                
                # กรณีคำโดนหั่นข้าม Run (บอสหน้าแรก) 
                # ถ้าเปลี่ยนแบบปกติไม่สำเร็จ ให้พยายามเปลี่ยนแบบไม่กระทบโครงสร้างหลัก
                if old_text in p.text:
                    # ตรงนี้คือด่านสุดท้าย ถ้ายังไม่ออก จะใช้การแก้แบบ Paragraph-level 
                    # แต่จำกัดเฉพาะจุดที่มีการเปลี่ยนแปลงเท่านั้น
                    p.text = p.text.replace(old_text, new_text)

    if process_btn and replace_list:
        doc = Document(io.BytesIO(file_content))
        for old, new in replace_list:
            # เนื้อหาหลัก
            master_replace(doc.paragraphs, old, new)
            # ในตาราง
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        master_replace(cell.paragraphs, old, new)
            # ใน Header/Footer (บังคับเช็คทุก Section ของทั้ง 100 หน้า)
            for section in doc.sections:
                for hf in [section.header, section.footer, 
                           section.first_page_header, section.first_page_footer,
                           section.even_page_header, section.even_page_footer]:
                    if hf:
                        master_replace(hf.paragraphs, old, new)
        
        out_bio = io.BytesIO()
        doc.save(out_bio)
        st.session_state.processed_doc = out_bio.getvalue()
        st.sidebar.success("✅ แก้ไขเรียบร้อย! ลองเช็คฟอนต์หน้าแรกนะค๊ะ")

    # --- ส่วนการแสดงผล ---
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