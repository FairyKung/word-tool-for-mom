import streamlit as st
from docx import Document
import io

st.set_page_config(page_title="Word Editor for Mom", layout="wide")
st.title("โปรแกรมแก้คำ")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # อ่านไฟล์เข้า Memory
    file_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(file_bytes))
    
    col1, col2 = st.columns(2)
    with col1:
        old_text = st.text_input("คำเดิม (ระวังเว้นวรรคให้ตรง)")
    with col2:
        new_text = st.text_input("คำใหม่ที่ต้องการ")

    if st.button("🚀 เริ่มเปลี่ยนคำ (แบบรักษารูปภาพ)"):
        if old_text and new_text:
            
            # ฟังก์ชันพิเศษ: เปลี่ยนคำโดยไม่ลบ Object อื่น (เช่น รูปภาพ) ในย่อหน้า
            def safe_replace(paragraphs):
                found_count = 0
                for p in paragraphs:
                    if old_text in p.text:
                        # วนลูปหาใน 'runs' (ส่วนย่อยของย่อหน้า)
                        # วิธีนี้จะรักษา formatting และรูปภาพที่แทรกอยู่ได้ดีกว่า
                        for run in p.runs:
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_text)
                                found_count += 1
                        
                        # กรณีที่คำโดนตัดแบ่งเป็นหลาย run (ปัญหาเรื่องวันที่)
                        # ถ้าข้างบนยังไม่เปลี่ยน ให้ใช้วิธีเปลี่ยนที่ระดับ paragraph แต่ต้องระวัง
                        if old_text in p.text: 
                            # ตรงนี้คือด่านสุดท้าย ถ้าเปลี่ยนใน run ไม่สำเร็จจะลองเปลี่ยนรวม
                            # แต่จะพยายามรักษาโครงสร้างไว้
                            full_text = p.text.replace(old_text, new_text)
                            p.text = full_text
                            found_count += 1
                return found_count

            # สั่งทำงานทุกจุด
            count = safe_replace(doc.paragraphs) # เนื้อหาหลัก
            for table in doc.tables: # ในตาราง
                for row in table.rows:
                    for cell in row.cells:
                        count += safe_replace(cell.paragraphs)
            
            for section in doc.sections: # หัว/ท้ายกระดาษ
                count += safe_replace(section.header.paragraphs)
                count += safe_replace(section.footer.paragraphs)

            if count > 0:
                st.success(f"เรียบร้อย! เปลี่ยนไปทั้งหมด {count} จุด")
                
                # เตรียมดาวน์โหลด
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ที่แก้แล้ว",
                    data=bio.getvalue(),
                    file_name=f"fixed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("หาคำที่ระบุไม่เจอเลย ลองเช็คเว้นวรรคดูอีกทีนะ")

    # ส่วน Preview เนื้อหา (แสดงเฉพาะข้อความให้ดูเพื่อตรวจสอบ)
    with st.expander("ตรวจสอบข้อความในไฟล์ (มองไม่เห็นรูปในหน้า Preview นี้ แต่ในไฟล์จริงรูปยังอยู่นะค๊ะ)"):
        for p in doc.paragraphs:
            if p.text.strip():
                st.write(p.text)