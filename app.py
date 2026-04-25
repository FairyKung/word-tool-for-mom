import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Pro", layout="wide")

# CSS ตกแต่ง
st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 20px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'ปราบหน้าแรก' (Template ไม่เคลื่อน) 📄✨")

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

    # --- ฟังก์ชันแก้คำแบบโหดแต่ถนอมไฟล์ (Hybrid Method) ---
    def master_replace(paragraphs, old_text, new_text):
        pattern = re.escape(old_text).replace(r'\ ', r'\s+')
        for p in paragraphs:
            if re.search(pattern, p.text):
                # วิธีรวมร่าง: เปลี่ยนที่ข้อความรวมของ Paragraph แล้วล้าง Run ที่เหลือ
                # วิธีนี้จะรักษาฟอนต์ของ Run แรกไว้ และทำให้คำที่โดนหั่นถูกเปลี่ยนแน่นอน
                new_full_text = re.sub(pattern, new_text, p.text)
                if p.runs:
                    p.runs[0].text = new_full_text
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = new_full_text

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
            # 3. ใน Header/Footer ทุกแบบ (จุดที่หน้าแรกชอบซ่อนตัว)
            for section in doc.sections:
                sections_to_check = [
                    section.header, section.footer,
                    section.first_page_header, section.first_page_footer,
                    section.even_page_header, section.even_page_footer
                ]
                for hf in sections_to_check:
                    if hf:
                        master_replace(hf.paragraphs, old, new)
        
        out_bio = io.BytesIO()
        doc.save(out_bio)
        st.session_state.processed_doc = out_bio.getvalue()
        st.sidebar.success("✅ แก้ไขเรียบร้อย! เช็คหน้า 1 ในพรีวิวนะค๊ะ")

    # --- ส่วนการแสดงผล ---
    tab1, tab2 = st.tabs(["📄 ไฟล์ต้นฉบับ", "✨ พรีวิวหลังแก้ไข"])
    
    with tab1:
        html_orig = mammoth.convert_to_html(io.BytesIO(file_content)).value
        st.markdown(f'<div class="paper-container"><div class="paper-content">{html_orig}</div></div>', unsafe_allow_html=True)

    with tab2:
        if st.session_state.processed_doc:
            st.download_button("📥 ดาวน์โหลดไฟล์ที่แก้ไขแล้ว", 
                             data=st.session_state.processed_doc, 
                             file_name=f"Fixed_{uploaded_file.name}",
                             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            html_fixed = mammoth.convert_to_html(io.BytesIO(st.session_state.processed_doc)).value
            st.markdown(f'<div class="paper-container"><div class="paper-content">{html_fixed}</div></div>', unsafe_allow_html=True)