import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Professional Word Editor", layout="wide")

# --- การตกแต่งหน้าจอให้เหมือนกระดาษ ---
st.markdown("""
    <style>
    .paper {
        background-color: white;
        padding: 40px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        margin: auto;
        min-height: 500px;
        width: 100%;
        color: black;
        font-family: 'Sarabun', sans-serif;
    }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชันหน้ากระดาษสมจริง 📄")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # อ่านไฟล์ต้นฉบับ
    file_bytes = uploaded_file.read()
    
    # --- ส่วนการตั้งค่าการเปลี่ยนคำ ---
    with st.sidebar:
        st.header("⚙️ เมนูแก้ไข")
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน")
        new_text = st.text_input("คำใหม่ที่ต้องการ")
        process_btn = st.button("🚀 ยืนยันการเปลี่ยนคำ")

    # --- ส่วนการทำงาน ---
    doc = Document(io.BytesIO(file_bytes))
    
    if process_btn and old_text and new_text:
        search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
        
        # ฟังก์ชันสแกนเปลี่ยนคำแบบรักษารูปภาพ
        def replace_all(paragraphs):
            for p in paragraphs:
                if re.search(search_pattern, p.text):
                    p.text = re.sub(search_pattern, new_text, p.text)

        # สั่งลุยทุกส่วนของไฟล์ (100 หน้าก็ทำหมด)
        replace_all(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_all(cell.paragraphs)
        for section in doc.sections:
            replace_all(section.header.paragraphs)
            replace_all(section.footer.paragraphs)
        
        st.sidebar.success("✅ เปลี่ยนคำในไฟล์เรียบร้อยแล้ว!")
        
        # ปุ่มดาวน์โหลด
        bio = io.BytesIO()
        doc.save(bio)
        st.sidebar.download_button("📥 ดาวน์โหลดไฟล์ที่แก้แล้ว", data=bio.getvalue(), file_name=f"fixed_{uploaded_file.name}")

    # --- ส่วนแสดงผลหน้ากระดาษ (Preview) ---
    st.subheader("👁️ ตัวอย่างหน้ากระดาษ (Preview)")
    
    # แปลง Word เป็น HTML เพื่อโชว์บนเว็บ (จะเห็นรูปและตารางด้วย)
    result = mammoth.convert_to_html(io.BytesIO(file_bytes))
    html_content = result.value
    
    # ใส่เนื้อหาลงในกล่องที่จัดสไตล์เป็นกระดาษ
    st.markdown(f'<div class="paper">{html_content}</div>', unsafe_allow_html=True)