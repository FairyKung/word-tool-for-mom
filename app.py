import streamlit as st
from docx import Document
import io

# หัวข้อหน้าเว็บ
st.title("โปรแกรมแก้คำในไฟล์ Word ")
st.write("อัปโหลดไฟล์ พิมพ์คำที่ต้องการเปลี่ยน ")

# 1. ส่วนการอัปโหลดไฟล์
uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # 2. ส่วนกรอกข้อมูล
    old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน", placeholder="เช่น 1 เมษายน 2569")
    new_text = st.text_input("คำใหม่ที่ต้องการ", placeholder="เช่น 5 เมษายน 2569")

    if st.button("เริ่มเปลี่ยนคำทั้งหมด"):
        if old_text and new_text:
            # อ่านไฟล์ Word
            doc = Document(uploaded_file)
            
            # Logic การเปลี่ยนคำ (Paragraphs)
            for p in doc.paragraphs:
                if old_text in p.text:
                    p.text = p.text.replace(old_text, new_text)
            
            # Logic การเปลี่ยนคำ (Tables)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if old_text in p.text:
                                p.text = p.text.replace(old_text, new_text)

            # เตรียมไฟล์สำหรับให้ดาวน์โหลดกลับ
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("แก้ไขเรียบร้อยแล้ว!")
            
            # 3. ปุ่มดาวน์โหลด
            st.download_button(
                label="ดาวน์โหลดไฟล์ที่แก้ไขแล้ว",
                data=bio.getvalue(),
                file_name=f"fixed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("กรุณากรอกคำให้ครบทั้งสองช่อง")