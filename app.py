import streamlit as st
from docx import Document
import io

st.set_page_config(page_title="Word Editor for Mom", layout="wide")
st.title("โปรแกรมแก้คำในไฟล์ Word ")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # อ่านไฟล์เข้าเครื่อง
    doc = Document(uploaded_file)
    
    # --- ส่วนที่ 1: แสดงเนื้อหาในไฟล์ให้ดู ---
    st.subheader("📄 เนื้อหาที่พบในไฟล์ปัจจุบัน:")
    all_text = []
    for p in doc.paragraphs:
        if p.text.strip():
            all_text.append(p.text)
    
    # แสดงเป็นกล่องข้อความให้อ่านง่าย
    with st.expander("คลิกเพื่อดูเนื้อหาทั้งหมดในไฟล์"):
        for line in all_text:
            st.write(line)

    st.divider()

    # --- ส่วนที่ 2: ตั้งค่าการเปลี่ยนคำ ---
    col1, col2 = st.columns(2)
    with col1:
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน", placeholder="เช่น 1 เมษายน 2569")
    with col2:
        new_text = st.text_input("คำใหม่ที่ต้องการ", placeholder="เช่น 5 เมษายน 2569")

    if st.button("เริ่มเปลี่ยนคำและตรวจสอบทั้งหมด"):
        if old_text and new_text:
            count = 0
            
            # ฟังก์ชันช่วยเปลี่ยนคำแบบละเอียด (แก้ปัญหา Runs โดนตัด)
            def replace_in_paragraphs(paragraphs):
                changes = 0
                for p in paragraphs:
                    if old_text in p.text:
                        # วิธีที่ชัวร์ที่สุดคือแทนที่ที่ระดับ paragraph.text ไปเลย
                        p.text = p.text.replace(old_text, new_text)
                        changes += 1
                return changes

            # 1. เปลี่ยนในเนื้อหาหลัก
            count += replace_in_paragraphs(doc.paragraphs)
            
            # 2. เปลี่ยนในตาราง
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        count += replace_in_paragraphs(cell.paragraphs)
            
            # 3. เปลี่ยนในหัวกระดาษ/ท้ายกระดาษ (Headers/Footers)
            for section in doc.sections:
                count += replace_in_paragraphs(section.header.paragraphs)
                count += replace_in_paragraphs(section.footer.paragraphs)

            if count > 0:
                st.success(f"สำเร็จ! พบและเปลี่ยนคำไปทั้งหมด {count} จุด")
                
                # เตรียมดาวน์โหลด
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ที่แก้ไขแล้ว",
                    data=bio.getvalue(),
                    file_name=f"fixed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning(f"หาคำว่า '{old_text}' ไม่เจอในไฟล์นี้ ลองเช็คตัวสะกดหรือเว้นวรรคอีกทีนะค๊ะ")
        else:
            st.error("กรุณากรอกคำให้ครบทั้งสองช่อง")