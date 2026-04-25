import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Perfect", layout="wide")

# CSS ตกแต่ง
st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 20px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'Template ไม่เคลื่อน' 📄✨")

# ระบบจัดการช่องคำที่จะเปลี่ยน
if 'num_pairs' not in st.session_state:
    st.session_state.num_pairs = 1
if 'processed_doc' not in st.session_state:
    st.session_state.processed_doc = None

def add_pair():
    st.session_state.num_pairs += 1

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # เก็บไฟล์ไว้ใน Memory
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

    # --- ฟังก์ชันการแก้คำแบบถนอม Template (ระดับ Run) ---
    def safe_replace(doc, old_text, new_text):
        pattern = re.escape(old_text).replace(r'\ ', r'\s+')
        
        # ค้นหาใน Paragraphs และ Tables
        def search_and_replace(paragraphs):
            for p in paragraphs:
                if re.search(pattern, p.text):
                    # วนลูปแก้ทีละ Run เพื่อรักษา Formatting
                    for run in p.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)
                        # กรณีคำโดนตัดแบ่ง Run (แก้แบบ Regex เบื้องต้น)
                        elif re.search(pattern, run.text):
                             run.text = re.sub(pattern, new_text, run.text)

        # ทำงานกับทุกส่วนของเอกสาร
        search_and_replace(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    search_and_replace(cell.paragraphs)
        for section in doc.sections:
            for hf in [section.header, section.footer, section.first_page_header, 
                       section.first_page_footer, section.even_page_header, section.even_page_footer]:
                if hf: search_and_replace(hf.paragraphs)

    # --- เมื่อกดปุ่มประมวลผล ---
    if process_btn and replace_list:
        working_doc = Document(io.BytesIO(file_content))
        for old, new in replace_list:
            safe_replace(working_doc, old, new)
        
        # บันทึกไฟล์ที่แก้แล้วลงใน session_state
        out_bio = io.BytesIO()
        working_doc.save(out_bio)
        st.session_state.processed_doc = out_bio.getvalue()
        st.sidebar.success("✅ แก้ไขเรียบร้อย! ตรวจพรีวิวด้านขวาได้เลยค่ะ")

    # --- ส่วนการแสดงผล (Preview) ---
    tab1, tab2 = st.tabs(["📄 ไฟล์ต้นฉบับ", "✨ พรีวิวหลังแก้ไข (ตัวอย่าง)"])
    
    with tab1:
        html_orig = mammoth.convert_to_html(io.BytesIO(file_content)).value
        st.markdown(f'<div class="paper-container"><div class="paper-content">{html_orig}</div></div>', unsafe_allow_html=True)

    with tab2:
        if st.session_state.processed_doc:
            # ปุ่มดาวน์โหลด (เอาไว้ข้างบนพรีวิวให้หาง่าย)
            st.download_button("📥 ดาวน์โหลดไฟล์ที่แก้ไขแล้ว", 
                             data=st.session_state.processed_doc, 
                             file_name=f"Fixed_{uploaded_file.name}",
                             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
            html_fixed = mammoth.convert_to_html(io.BytesIO(st.session_state.processed_doc)).value
            st.markdown(f'<div class="paper-container"><div class="paper-content">{html_fixed}</div></div>', unsafe_allow_html=True)
        else:
            st.warning("ยังไม่มีการแก้ไข กรุณากรอกข้อมูลในแถบด้านซ้ายแล้วกดปุ่มประมวลผลค่ะ")