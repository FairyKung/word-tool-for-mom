import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Ultimate Pro", layout="wide")

st.markdown("""
    <style>
    .paper-container { background-color: #f0f2f6; padding: 30px; border-radius: 10px; }
    .paper-content { background-color: white; padding: 50px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); 
                    margin: auto; min-height: 800px; color: black; font-family: 'Sarabun', sans-serif; line-height: 1.6; }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชัน 'เปลี่ยนยกชุด' 📄🚀")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx)", type="docx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    with st.sidebar:
        st.header("🔍 ตั้งค่าการเปลี่ยนคำหลายชุด")
        st.info("พิมพ์คำเดิม [เครื่องหมายเท่ากับ] คำใหม่ (บรรทัดละคู่)\nเช่น:\n1 เมษายน = 5 พฤษภาคม\nนาย ก = นาย ข")
        
        # ช่องกรอกแบบหลายบรรทัด
        mapping_text = st.text_area("รายการคำที่ต้องการเปลี่ยน", height=200)
        
        if st.button("🚀 เริ่มเปลี่ยนคำทั้งหมดทันที"):
            if mapping_text:
                doc = Document(io.BytesIO(file_bytes))
                total_changes = 0
                
                # แปลงข้อความที่กรอกมาเป็น Dictionary
                replace_map = {}
                for line in mapping_text.split('\n'):
                    if '=' in line:
                        parts = line.split('=')
                        old_val = parts[0].strip()
                        new_val = parts[1].strip()
                        if old_val:
                            replace_map[old_val] = new_val

                if replace_map:
                    # ฟังก์ชันหลักสำหรับการเปลี่ยนคำ (ใช้ Regex แบบทนทานพิเศษ)
                    def perform_replace(text):
                        if not text: return text
                        new_text = text
                        for old, new in replace_map.items():
                            pattern = re.escape(old).replace(r'\ ', r'\s+')
                            new_text = re.sub(pattern, new, new_text)
                        return new_text

                    # 1. ท่าไม้ตาย: สแกน XML ทุกอณู (รวมถึงหน้าแรกและกล่องข้อความ)
                    for t in doc._element.xpath('//w:t'):
                        original = t.text
                        replaced = perform_replace(t.text)
                        if original != replaced:
                            t.text = replaced
                            total_changes += 1

                    # 2. บังคับย้ำที่ Paragraphs และ Tables (เพื่อความชัวร์เรื่อง Formatting)
                    for p in doc.paragraphs:
                        p.text = perform_replace(p.text)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                cell.text = perform_replace(cell.text)

                    # 3. ย้ำที่ Header/Footer ทุกหน้า (รวมหน้าแรก)
                    for section in doc.sections:
                        for hf in [section.header, section.footer, 
                                   section.first_page_header, section.first_page_footer,
                                   section.even_page_header, section.even_page_footer]:
                            if hf:
                                for p in hf.paragraphs:
                                    p.text = perform_replace(p.text)

                    st.success(f"เสร็จเรียบร้อย! แก้ไขไปทั้งหมด {total_changes} จุด")
                    out_bio = io.BytesIO()
                    doc.save(out_bio)
                    st.download_button("📥 ดาวน์โหลดไฟล์", data=out_bio.getvalue(), file_name=f"multi_fixed_{uploaded_file.name}")
                else:
                    st.warning("กรุณากรอกรูปแบบให้ถูกต้อง (คำเดิม = คำใหม่)")

    # ส่วน Preview
    st.subheader("👁️ ตัวอย่างหน้ากระดาษ")
    html_res = mammoth.convert_to_html(io.BytesIO(file_bytes))
    st.markdown(f'<div class="paper-container"><div class="paper-content">{html_res.value}</div></div>', unsafe_allow_html=True)