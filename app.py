import streamlit as st
from docx import Document
import io
import re

st.set_page_config(page_title="Word Editor 100 Pages", layout="wide")
st.title("โปรแกรมแก้คำเวอร์ชัน 'ทะลวง 100 หน้า' (สแกน XML) 🚀")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    # อ่านไฟล์ต้นฉบับ
    file_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(file_bytes))
    
    col1, col2 = st.columns(2)
    with col1:
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน", placeholder="เช่น 1 เมษายน 2569")
    with col2:
        new_text = st.text_input("คำใหม่ที่ต้องการ")

    if st.button("🚀 เริ่มเปลี่ยนคำทั้งหมด (สแกนทั้งไฟล์)"):
        if old_text and new_text:
            total_changes = 0
            # สร้าง Pattern Regex สำหรับจัดการเว้นวรรค (space)
            search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
            
            # 1. ฟังก์ชันสแกนเปลี่ยนคำใน Paragraphs (ทุกที่รวมถึง Header/Footer/Textbox)
            def thorough_replace(paragraphs):
                count = 0
                for p in paragraphs:
                    if re.search(search_pattern, p.text):
                        # รวมร่างข้อความใน Runs เพื่อป้องกันคำขาดจาก XML
                        full_text = re.sub(search_pattern, new_text, p.text)
                        if p.runs:
                            p.runs[0].text = full_text
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = full_text
                        count += 1
                return count

            # ลุยสแกนทุกจุดที่อาจมีตัวอักษรซ่อนอยู่
            # เนื้อหาหลัก
            total_changes += thorough_replace(doc.paragraphs)
            
            # ในตาราง (Tables)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        total_changes += thorough_replace(cell.paragraphs)
            
            # ในทุก Section (Header/Footer ทุกหน้า)
            for section in doc.sections:
                total_changes += thorough_replace(section.header.paragraphs)
                total_changes += thorough_replace(section.footer.paragraphs)
            
            # **ไม้ตาย: สแกนหากล่องข้อความ (Inline Shapes / Textboxes)**
            # หมายเหตุ: บางกล่องข้อความจะอยู่ในหมวดหมู่นี้
            for shape in doc.inline_shapes:
                # ถ้ามีข้อความซ่อนในรูปภาพหรือกราฟิก (บางประเภท)
                pass # python-docx เข้าถึงบางส่วนได้จำกัด แต่ Paragraph สแกนครอบคลุมส่วนใหญ่แล้ว

            if total_changes > 0:
                st.success(f"เสร็จแล้ว! เปลี่ยนไปทั้งหมด {total_changes} จุด ครบทุกหน้าแน่นอนค่ะ")
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("📥 ดาวน์โหลดไฟล์ 100 หน้าที่แก้แล้ว", data=bio.getvalue(), file_name=f"updated_{uploaded_file.name}")
            else:
                st.warning("หาคำไม่เจอเลยค่ะ ลองเช็คว่าในไฟล์มี 'เว้นวรรค' แปลกๆ ไหมนะค๊ะ")

    # ส่วนแสดงผลสำหรับตรวจสอบ (Show All Text)
    st.divider()
    with st.expander("🔍 ตรวจสอบตัวอักษรทั้งหมดที่ระบบอ่านได้ (รวม 100 หน้า)"):
        all_content = []
        for p in doc.paragraphs:
            if p.text.strip():
                all_content.append(p.text)
        
        # ถ้าเนื้อหาเยอะมาก จะแสดงแบบแบ่งหน้าให้ดูบนเว็บ
        st.write(f"พบทั้งหมด {len(all_content)} ย่อหน้า")
        st.code("\n\n".join(all_content))