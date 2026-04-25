import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Pro", layout="wide")

# CSS ตกแต่ง (คงเดิมไว้เพราะแม่ชอบ)
st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 30px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 50px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; line-height: 1.6; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'ทะลวงทุก Section' (แก้ครบ 100%) 📄✨")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx) ของคุณแม่", type="docx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    with st.sidebar:
        st.header("🔍 ตั้งค่าการเปลี่ยนคำ")
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน")
        new_text = st.text_input("คำใหม่ที่ต้องการ")
        
        if st.button("🚀 เริ่มเปลี่ยนคำทั้งไฟล์"):
            if old_text and new_text:
                doc = Document(io.BytesIO(file_bytes))
                search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
                total = 0

                # ฟังก์ชันช่วยเปลี่ยน (เปลี่ยนแบบยืดหยุ่น)
                def change_text(paragraphs):
                    c = 0
                    for p in paragraphs:
                        if re.search(search_pattern, p.text):
                            p.text = re.sub(search_pattern, new_text, p.text)
                            c += 1
                    return c

                # 1. เนื้อหาหลัก (Body) - ปกติจะแก้ได้หมดทุกหน้า
                total += change_text(doc.paragraphs)

                # 2. ในตาราง (Tables)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            total += change_text(cell.paragraphs)

                # 3. จุดสำคัญ! หัว/ท้ายกระดาษ (ต้องวนลูปทุก Section เพราะแต่ละหน้าอาจไม่เหมือนกัน)
                for i, section in enumerate(doc.sections):
                    # แก้ใน Header ปกติ, Header หน้าแรก และ Header หน้าคู่/คี่
                    total += change_text(section.header.paragraphs)
                    if section.first_page_header:
                        total += change_text(section.first_page_header.paragraphs)
                    if section.even_page_header:
                        total += change_text(section.even_page_header.paragraphs)
                        
                    # แก้ใน Footer ให้ครบทุกแบบ
                    total += change_text(section.footer.paragraphs)
                    if section.first_page_footer:
                        total += change_text(section.first_page_footer.paragraphs)
                    if section.even_page_footer:
                        total += change_text(section.even_page_footer.paragraphs)

                if total > 0:
                    st.success(f"สำเร็จ! แก้ไขไปทั้งหมด {total} จุด (ตรวจสอบหน้า 9-10 ได้เลยค่ะ)")
                    out_bio = io.BytesIO()
                    doc.save(out_bio)
                    st.download_button("📥 ดาวน์โหลดไฟล์", data=out_bio.getvalue(), file_name=f"fixed_{uploaded_file.name}")
                else:
                    st.warning("หาคำไม่เจอเลยค่ะ ลองเช็คเว้นวรรคดูอีกทีนะค๊ะ")

    # ส่วน Preview
    st.subheader("👁️ ตรวจสอบเนื้อหาในหน้ากระดาษ")
    html_res = mammoth.convert_to_html(io.BytesIO(file_bytes))
    st.markdown(f'<div class="paper-container"><div class="paper-content">{html_res.value}</div></div>', unsafe_allow_html=True)