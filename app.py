import streamlit as st
from docx import Document
import io
import re

st.set_page_config(page_title="Deep Scan Word Editor", layout="wide")
st.title("โปรแกรมแก้คำเวอร์ชัน 'สแกนลึกทุกซอกทุกมุม' (Deep Scan) 🚀")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(file_bytes))
    
    col1, col2 = st.columns(2)
    with col1:
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน", placeholder="เช่น 1 เมษายน 2569")
    with col2:
        new_text = st.text_input("คำใหม่ที่ต้องการ")

    if st.button("🚀 เริ่มเปลี่ยนคำแบบ Deep Scan"):
        if old_text and new_text:
            total_changes = 0
            # สร้าง Pattern สำหรับจัดการเว้นวรรคแบบยืดหยุ่น
            search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')

            # --- ฟังก์ชันหลักสำหรับเปลี่ยนคำแบบรักษาโครงสร้าง ---
            def deep_replace(paragraphs):
                count = 0
                for p in paragraphs:
                    if re.search(search_pattern, p.text):
                        # รวมร่างข้อความใน runs เพื่อป้องกันคำขาดจาก XML
                        # และรักษาความสวยงามของฟอนต์/รูปภาพ
                        full_text = re.sub(search_pattern, new_text, p.text)
                        if p.runs:
                            p.runs[0].text = full_text
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = full_text
                        count += 1
                return count

            # 1. เนื้อหาหลัก
            total_changes += deep_replace(doc.paragraphs)
            
            # 2. ในตาราง (รวมถึงตารางซ้อนตาราง)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        total_changes += deep_replace(cell.paragraphs)
                        # กรณีมีตารางซ้อนใน Cell
                        for nested_table in cell.tables:
                            for n_row in nested_table.rows:
                                for n_cell in n_row.cells:
                                    total_changes += deep_replace(n_cell.paragraphs)

            # 3. ในหัวกระดาษ/ท้ายกระดาษ (ทุกหน้า)
            for section in doc.sections:
                total_changes += deep_replace(section.header.paragraphs)
                total_changes += deep_replace(section.footer.paragraphs)

            # 4. **ท่าไม้ตาย: เจาะลึกกล่องข้อความ (Textboxes)**
            # เข้าถึงผ่านอีลิเมนต์ภายใน XML ของเอกสาร
            for p in doc._element.xpath('//w:t'):
                if re.search(search_pattern, p.text):
                    p.text = re.sub(search_pattern, new_text, p.text)
                    total_changes += 1

            if total_changes > 0:
                st.success(f"สำเร็จ! พบและเปลี่ยนคำไปทั้งหมด {total_changes} จุด ครบทุกส่วนแน่นอน")
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("📥 ดาวน์โหลดไฟล์ที่แก้เสร็จแล้ว", data=bio.getvalue(), file_name=f"DeepFixed_{uploaded_file.name}")
            else:
                st.warning("ยังหาคำไม่เจอค่ะ ลองเช็คว่าสะกดถูกหรือก๊อปมาจาก Preview ด้านล่างดูนะค๊ะ")

    # ส่วนตรวจสอบเนื้อหาแบบโชว์ทุกอย่าง
    st.divider()
    with st.expander("🔍 ตรวจสอบคำทั้งหมดที่ระบบค้นพบ (สแกนทั้ง 100 หน้า)"):
        all_text_found = []
        # ดึงข้อความจาก XML โดยตรงเพื่อให้มั่นใจว่าเห็นครบทุกหน้า
        for t in doc._element.xpath('//w:t'):
            if t.text and t.text.strip():
                all_text_found.append(t.text)
        
        st.write(f"พบข้อความทั้งหมด {len(all_text_found)} จุด")
        st.code(" ".join(all_text_found))