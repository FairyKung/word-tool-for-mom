import streamlit as st
from docx import Document
import io
import re

st.set_page_config(page_title="Word Editor for Mom", layout="wide")
st.title("โปรแกรมแก้คำเวอร์ชัน 'แก้ทุกอย่างแต่รักษารูป' ❤️")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(file_bytes))
    
    col1, col2 = st.columns(2)
    with col1:
        old_text = st.text_input("คำเดิม (เช่น วันที่ที่มีเว้นวรรคเยอะๆ)")
    with col2:
        new_text = st.text_input("คำใหม่ที่ต้องการ")

    if st.button("🚀 เริ่มเปลี่ยนคำทั้งหมด"):
        if old_text and new_text:
            
            def better_replace(paragraphs):
                count = 0
                for p in paragraphs:
                    # ตรวจสอบว่ามีคำที่ต้องการอยู่ในย่อหน้าหรือไม่ (แบบรวมร่าง)
                    if old_text in p.text:
                        # สร้าง Pattern สำหรับค้นหาแบบไม่สนจำนวนช่องว่าง (Regex)
                        # เช่น '1   เมษายน' หรือ '1 เมษายน' จะมองว่าเหมือนกัน
                        search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
                        
                        if re.search(search_pattern, p.text):
                            # เปลี่ยนคำในระดับ Paragraph text โดยตรง
                            # วิธีนี้จะรักษารูปภาพไว้ได้ "ถ้า" รูปนั้นไม่ได้ถูกตั้งค่าเป็น Inline กับตัวอักษรที่ถูกแก้
                            new_paragraph_text = re.sub(search_pattern, new_text, p.text)
                            
                            # เทคนิคการเปลี่ยนเนื้อหาโดยรักษาโครงสร้าง:
                            # เราจะล้างข้อความใน Runs ทั้งหมดทิ้ง แต่เก็บรูปภาพ (ที่มักจะอยู่ใน Run พิเศษ) ไว้
                            # หรือใช้วิธีเปลี่ยนที่ Run แรกแล้วล้าง Run ที่เหลือ
                            if len(p.runs) > 0:
                                p.runs[0].text = new_paragraph_text
                                for r in p.runs[1:]:
                                    r.text = ""
                            count += 1
                return count

            # สั่งลุยทุกส่วน
            found = better_replace(doc.paragraphs)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        found += better_replace(cell.paragraphs)

            if found > 0:
                st.success(f"เปลี่ยนสำเร็จ {found} จุดแล้วค่ะ! รูปยังอยู่ดีไหมค๊ะ?")
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("📥 ดาวน์โหลดไฟล์", data=bio.getvalue(), file_name=f"fixed_{uploaded_file.name}")
            else:
                st.warning("ยังหาคำไม่เจอเลยค่ะ ลองก๊อปปี้คำจาก Preview ด้านล่างมาวางดูนะค๊ะ")

    # ส่วน Preview เนื้อหาแบบดิบๆ เพื่อให้ก๊อปปี้ง่าย
    with st.expander("🔍 ดูข้อความดิบในไฟล์ (สำหรับก๊อปปี้มาวาง)"):
        for p in doc.paragraphs:
            if p.text.strip():
                st.code(p.text) # ใช้ st.code เพื่อให้เห็นเว้นวรรคชัดๆ และก๊อปง่าย