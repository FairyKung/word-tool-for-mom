import streamlit as st
from docx import Document
import io
import re

st.set_page_config(page_title="Word Editor for Mom", layout="wide")
st.title("โปรแกรมแก้คำเวอร์ชัน 'สแกนละเอียดทุกหน้า' ❤️")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # อ่านไฟล์
    file_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(file_bytes))
    
    col1, col2 = st.columns(2)
    with col1:
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน", help="แนะนำให้ก๊อปปี้จากด้านล่างมาวาง")
    with col2:
        new_text = st.text_input("คำใหม่ที่ต้องการ")

    # ฟังก์ชันช่วยเปลี่ยนคำแบบ Regex (จัดการเรื่องเว้นวรรค)
    def universal_replace(paragraphs, search, replace):
        count = 0
        search_pattern = re.escape(search).replace(r'\ ', r'\s+')
        for p in paragraphs:
            if re.search(search_pattern, p.text):
                # รวมร่าง Run เพื่อป้องกันคำขาด แล้วค่อยเปลี่ยน
                full_text = re.sub(search_pattern, replace, p.text)
                if len(p.runs) > 0:
                    p.runs[0].text = full_text
                    for r in p.runs[1:]:
                        r.text = ""
                count += 1
        return count

    if st.button("🚀 เริ่มเปลี่ยนคำทั้งหมดในทุกหน้า"):
        if old_text and new_text:
            total_found = 0
            
            # 1. สแกนเนื้อหาหลัก (Paragraphs)
            total_found += universal_replace(doc.paragraphs, old_text, new_text)
            
            # 2. สแกนในตารางทั้งหมด (Tables)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        total_found += universal_replace(cell.paragraphs, old_text, new_text)
            
            # 3. สแกนในหัวและท้ายกระดาษ (Headers & Footers) - จุดที่มักจะตกหล่น
            for section in doc.sections:
                total_found += universal_replace(section.header.paragraphs, old_text, new_text)
                total_found += universal_replace(section.footer.paragraphs, old_text, new_text)

            if total_found > 0:
                st.success(f"สำเร็จ! พบและเปลี่ยนคำไปทั้งหมด {total_found} จุด ในทุกส่วนของไฟล์")
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("📥 ดาวน์โหลดไฟล์ที่แก้ครบทุกหน้า", data=bio.getvalue(), file_name=f"full_fixed_{uploaded_file.name}")
            else:
                st.warning("ไม่พบคำที่ระบุในส่วนใดของไฟล์เลยค่ะ")

    # --- ส่วนการแสดงผล Preview ให้ครบทุกส่วน ---
    st.divider()
    st.subheader("🔍 ตรวจสอบเนื้อหาทั้งหมดที่ตรวจพบในไฟล์")
    
    tab1, tab2, tab3 = st.tabs(["เนื้อหาหลัก", "ในตาราง", "หัว/ท้ายกระดาษ"])
    
    with tab1:
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip():
                st.code(f"[ย่อหน้าที่ {i+1}]: {p.text}")

    with tab2:
        for i, table in enumerate(doc.tables):
            st.write(f"📊 ตารางที่ {i+1}")
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                st.write(" | ".join(row_data))

    with tab3:
        for i, section in enumerate(doc.sections):
            st.write(f"📌 ส่วนที่ {i+1}")
            st.text(f"Header: {section.header.paragraphs[0].text if section.header.paragraphs else 'ว่าง'}")
            st.text(f"Footer: {section.footer.paragraphs[0].text if section.footer.paragraphs else 'ว่าง'}")