import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Pro", layout="wide")

# CSS ตกแต่งหน้าจอ (คงเดิมไว้เพื่อให้คุณแม่ใช้งานง่าย)
st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 30px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 50px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; line-height: 1.6; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'ปลดล็อกหน้าแรก' 📄✨")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx) ของคุณแม่", type="docx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    with st.sidebar:
        st.header("🔍 ตั้งค่าการเปลี่ยนคำ")
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน")
        new_text = st.text_input("คำใหม่ที่ต้องการ")
        
        if st.button("🚀 เริ่มเปลี่ยนคำทั้งไฟล์ (รวมหน้าแรก)"):
            if old_text and new_text:
                doc = Document(io.BytesIO(file_bytes))
                search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
                total = 0

                # --- 1. ท่าไม้ตายล้างบาง (แก้ทุกจุดที่เป็นตัวอักษรใน XML) ---
                # วิธีนี้ปกติควรจะครอบคลุมหน้าแรกด้วย แต่เราจะเสริมทัพในข้อ 2
                for t in doc._element.xpath('//w:t'):
                    if re.search(search_pattern, t.text):
                        new_val = re.sub(search_pattern, new_text, t.text)
                        if t.text != new_val:
                            t.text = new_val
                            total += 1

                # --- 2. บังคับตรวจเช็ค Header/Footer หน้าแรก (กรณีหน้าแรกตั้งค่าแยกไว้) ---
                def force_replace(paragraphs):
                    c = 0
                    for p in paragraphs:
                        if re.search(search_pattern, p.text):
                            p.text = re.sub(search_pattern, new_text, p.text)
                            c += 1
                    return c

                for section in doc.sections:
                    # เจาะจงหน้าแรก (First Page)
                    if section.first_page_header:
                        total += force_replace(section.first_page_header.paragraphs)
                    if section.first_page_footer:
                        total += force_replace(section.first_page_footer.paragraphs)
                    # เจาะจงหน้าคู่/คี่ (Odd/Even)
                    if section.even_page_header:
                        total += force_replace(section.even_page_header.paragraphs)
                    if section.even_page_footer:
                        total += force_replace(section.even_page_footer.paragraphs)

                if total > 0:
                    st.success(f"สำเร็จ! แก้ไขหน้าแรกและหน้าอื่นๆ รวม {total} จุดเรียบร้อยค่ะ")
                    out_bio = io.BytesIO()
                    doc.save(out_bio)
                    st.download_button("📥 ดาวน์โหลดไฟล์ที่สมบูรณ์", data=out_bio.getvalue(), file_name=f"Fixed_All_{uploaded_file.name}")
                else:
                    st.warning("ยังหาคำไม่เจอในหน้าแรกเลยค่ะ ลองเช็คเว้นวรรคดูอีกทีนะค๊ะ")

    # ส่วน Preview
    st.subheader("👁️ ตรวจสอบเนื้อหา (ลองเช็คหน้าแรกดูนะค๊ะ)")
    html_res = mammoth.convert_to_html(io.BytesIO(file_bytes))
    st.markdown(f'<div class="paper-container"><div class="paper-content">{html_res.value}</div></div>', unsafe_allow_html=True)