import streamlit as st
from docx import Document
import io
import re

st.set_page_config(page_title="Word Editor for Mom", layout="wide")
st.title("โปรแกรมแก้คำในไฟล์ Word ")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # อ่านไฟล์
    doc = Document(uploaded_file)
    
    # --- ส่วนการแก้ไขเรื่องเว้นวรรค (Flexible Search) ---
    st.sidebar.header("⚙️ ตั้งค่าการค้นหา")
    use_regex = st.sidebar.checkbox("ค้นหาแบบยืดหยุ่น (แก้ปัญหาวรรคนเกิน)", value=True)

    col1, col2 = st.columns(2)
    with col1:
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน", placeholder="ก๊อปปี้คำจากข้างล่างมาวางจะแม่นที่สุด")
    with col2:
        new_text = st.text_input("คำใหม่ที่ต้องการ")

    # --- ส่วนแสดงหน้ากระดาษ (Preview เนื้อหาแยกตาม Paragraph) ---
    st.subheader("📄 ตรวจสอบเนื้อหาในไฟล์")
    
    # ฟังก์ชันช่วยเปลี่ยนคำแบบจัดการเรื่อง Space
    def smart_replace(text, search, replace):
        if use_regex:
            # สร้าง pattern ที่ยอมรับเว้นวรรคกี่ทีก็ได้ระหว่างคำ
            pattern = re.escape(search).replace(r'\ ', r'\s+')
            return re.sub(pattern, replace, text)
        return text.replace(search, replace)

    if st.button("🚀 เริ่มเปลี่ยนคำและบันทึกไฟล์"):
        if old_text and new_text:
            # เปลี่ยนใน Paragraphs
            for p in doc.paragraphs:
                p.text = smart_replace(p.text, old_text, new_text)
            
            # เปลี่ยนใน Tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            p.text = smart_replace(p.text, old_text, new_text)
            
            st.success("แก้ไขคำเรียบร้อยแล้ว!")
            
            # เตรียมดาวน์โหลด
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button("📥 ดาวน์โหลดไฟล์ใหม่", data=bio.getvalue(), file_name=f"edited_{uploaded_file.name}")

    # แสดง Preview (เนื่องจากการ Render เป็นภาพจริงทำได้ยากบน Cloud ฟรี 
    # จึงโชว์เป็นลำดับ Paragraph เพื่อให้เห็นภาพรวมของหน้ากระดาษแทน)
    st.info("💡 ด้านล่างนี้คือลำดับข้อความในไฟล์ ")
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            with st.container():
                st.markdown(f"**บรรทัดที่ {i+1}:** {p.text}")
                st.divider()