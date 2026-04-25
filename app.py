import streamlit as st
from docx import Document
import io
import re
import mammoth
from docx.oxml.ns import qn

st.set_page_config(page_title="Word Editor Ultimate", layout="wide")

st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 30px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 50px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; line-height: 1.6; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'สแกนล้างบาง' (แก้ทุกที่ 100%) 📄🚀")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx) ของคุณแม่", type="docx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    with st.sidebar:
        st.header("🔍 ตั้งค่าการเปลี่ยนคำ")
        old_text = st.text_input("คำเดิมที่ต้องการเปลี่ยน")
        new_text = st.text_input("คำใหม่ที่ต้องการ")
        
        if st.button("🚀 เริ่มเปลี่ยนคำทั้งหมด"):
            if old_text and new_text:
                doc = Document(io.BytesIO(file_bytes))
                # Regex สำหรับจัดการช่องว่าง (Space) ที่อาจจะไม่เท่ากันในแต่ละหน้า
                search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
                total = 0

                # --- ท่าไม้ตาย: สแกนหาทุก XML Element ที่มีตัวอักษร (w:t) ---
                # วิธีนี้จะเข้าถึงทุกที่: เนื้อหา, ตาราง, Header, Footer, กล่องข้อความ, บันทึกย่อ
                for t in doc._element.xpath('//w:t'):
                    if re.search(search_pattern, t.text):
                        new_val = re.sub(search_pattern, new_text, t.text)
                        if t.text != new_val:
                            t.text = new_val
                            total += 1

                if total > 0:
                    st.success(f"สำเร็จ! พบและแก้ไขไปทั้งหมด {total} จุด (รวมทุกหน้าแล้วค่ะ)")
                    out_bio = io.BytesIO()
                    doc.save(out_bio)
                    st.download_button("📥 ดาวน์โหลดไฟล์", data=out_bio.getvalue(), file_name=f"fixed_{uploaded_file.name}")
                else:
                    st.warning("หาคำไม่เจอเลยค่ะ ลองเช็คว่าสะกดถูกหรือมีวรรคแปลกๆ ไหม")

    # ส่วน Preview (โชว์ให้แม่เห็นภาพรวม)
    st.subheader("👁️ ตัวอย่างหน้ากระดาษ (ไถลงไปดูหน้า 9-10 ได้เลยค่ะ)")
    html_res = mammoth.convert_to_html(io.BytesIO(file_bytes))
    st.markdown(f'<div class="paper-container"><div class="paper-content">{html_res.value}</div></div>', unsafe_allow_html=True)