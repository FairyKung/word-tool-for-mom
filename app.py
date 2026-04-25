import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Pro", layout="wide")

# CSS ตกแต่งให้สวยงาม
st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 20px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 40px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; }
    .stButton>button { width: 100%; border-radius: 20px; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำแบบถนอมรูปภาพและเทมเพลต 📄✨")

# --- ระบบปุ่มบวกเพิ่มคำ (Session State) ---
if 'num_pairs' not in st.session_state:
    st.session_state.num_pairs = 1

def add_pair():
    st.session_state.num_pairs += 1

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    with st.sidebar:
        st.header("🔍 รายการคำที่ต้องการเปลี่ยน")
        
        replace_list = []
        for i in range(st.session_state.num_pairs):
            st.markdown(f"**ชุดที่ {i+1}**")
            col_a, col_b = st.columns(2)
            with col_a:
                old = st.text_input(f"คำเดิม", key=f"old_{i}")
            with col_b:
                new = st.text_input(f"คำใหม่", key=f"new_{i}")
            if old:
                replace_list.append((old, new))
            st.divider()
        
        st.button("➕ เพิ่มช่องคำที่จะเปลี่ยน", on_click=add_pair)
        
        process_btn = st.button("🚀 เริ่มเปลี่ยนคำทั้งหมด", type="primary")

    if process_btn and replace_list:
        doc = Document(io.BytesIO(file_bytes))
        
        def smart_replace(paragraphs):
            count = 0
            for p in paragraphs:
                for old_val, new_val in replace_list:
                    # ใช้ Regex เพื่อรองรับเว้นวรรคที่ไม่เท่ากัน
                    pattern = re.escape(old_val).replace(r'\ ', r'\s+')
                    if re.search(pattern, p.text):
                        # เปลี่ยนที่ระดับ Runs เพื่อรักษารูปภาพและ Format
                        # หมายเหตุ: หากคำโดนตัดแบ่ง Run ระบบจะรวมชั่วคราวเพื่อหา แต่พยายามรักษา Object อื่นไว้
                        full_text = re.sub(pattern, new_val, p.text)
                        # วิธีที่ปลอดภัยที่สุดในการรักษา Template คือการเขียนทับ text ของย่อหน้า
                        # แต่ถ้ามีย่อหน้าที่มีรูปภาพ เราจะใช้วิธีเปลี่ยนที่ Run แรกและล้างที่เหลือ
                        if len(p.runs) > 0:
                            p.runs[0].text = full_text
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = full_text
                        count += 1
            return count

        # สแกนทุกส่วนของเอกสาร
        total = smart_replace(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    total += smart_replace(cell.paragraphs)
        
        # ส่วน Header/Footer (ครอบคลุมหน้าแรกและหน้าคู่คี่)
        for section in doc.sections:
            for hf in [section.header, section.footer, section.first_page_header, 
                       section.first_page_footer, section.even_page_header, section.even_page_footer]:
                if hf:
                    total += smart_replace(hf.paragraphs)

        st.sidebar.success(f"✅ แก้ไขเสร็จสิ้นทั้งหมด {total} จุด")
        out_bio = io.BytesIO()
        doc.save(out_bio)
        st.sidebar.download_button("📥 ดาวน์โหลดไฟล์", data=out_bio.getvalue(), file_name=f"Fixed_{uploaded_file.name}")

    # --- ส่วน Preview (โชว์เนื้อหาจริง) ---
    st.subheader("👁️ ตรวจสอบความเรียบร้อย")
    html_res = mammoth.convert_to_html(io.BytesIO(file_bytes))
    st.markdown(f'<div class="paper-container"><div class="paper-content">{html_res.value}</div></div>', unsafe_allow_html=True)