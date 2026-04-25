import streamlit as st
from docx import Document
import io
import re
import mammoth

st.set_page_config(page_title="Word Editor Pro", layout="wide")

# ตกแต่ง CSS ให้ดูเหมือนกระดาษขาว
st.markdown("""
    <style>
    .paper-container {
        background-color: #f0f2f6;
        padding: 30px;
        border-radius: 10px;
    }
    .paper-content {
        background-color: white;
        padding: 50px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        margin: auto;
        min-height: 800px;
        color: black;
        font-family: 'Sarabun', sans-serif;
        line-height: 1.6;
    }
    </style>
""", unsafe_allow_html=True)

st.title("โปรแกรมแก้คำเวอร์ชันอ่านครบ 100 หน้า 📄✨")

uploaded_file = st.file_uploader("เลือกไฟล์ Word (.docx) ของคุณแม่", type="docx")

if uploaded_file is not None:
    # อ่านไฟล์ต้นฉบับ
    file_bytes = uploaded_file.read()
    
    with st.sidebar:
        st.header("🔍 ตั้งค่าการเปลี่ยนคำ")
        old_text = st.text_input("คำเดิม (เช่น วันที่ที่มีวรรคเยอะๆ)")
        new_text = st.text_input("คำใหม่ที่ต้องการ")
        
        st.divider()
        if st.button("🚀 เริ่มเปลี่ยนคำทั้งไฟล์"):
            if old_text and new_text:
                # โหลด docx
                doc = Document(io.BytesIO(file_bytes))
                
                # ใช้ Regex แบบยืดหยุ่น (หาเจอแม้เว้นวรรคไม่เท่ากัน)
                search_pattern = re.escape(old_text).replace(r'\ ', r'\s+')
                
                def deep_clean_replace(paragraphs):
                    count = 0
                    for p in paragraphs:
                        if re.search(search_pattern, p.text):
                            # เปลี่ยนแบบรวมร่างเพื่อความชัวร์
                            p.text = re.sub(search_pattern, new_text, p.text)
                            count += 1
                    return count

                # สแกนทุกจุด (Body, Tables, Header, Footer)
                total = deep_clean_replace(doc.paragraphs)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            total += deep_clean_replace(cell.paragraphs)
                for section in doc.sections:
                    total += deep_clean_replace(section.header.paragraphs)
                    total += deep_clean_replace(section.footer.paragraphs)
                
                if total > 0:
                    st.success(f"สำเร็จ! พบและแก้ไขไปทั้งหมด {total} จุดค่ะ")
                    out_bio = io.BytesIO()
                    doc.save(out_bio)
                    st.download_button("📥 ดาวน์โหลดไฟล์ที่แก้แล้ว", data=out_bio.getvalue(), file_name=f"fixed_{uploaded_file.name}")
                else:
                    st.warning("หาคำไม่เจอเลยค่ะ ลองเช็คตัวสะกดอีกทีนะค๊ะ")

    # --- ส่วนแสดง Preview แบบสวยงาม ---
    st.subheader("👁️ ตรวจสอบเนื้อหาในหน้ากระดาษ")
    st.info("💡 ระบบจะแสดงเนื้อหาทั้งหมดที่มีในไฟล์ (รวมถึงหน้าหลังๆ) ให้คุณแม่ตรวจดูก่อนได้ที่นี่ค่ะ")
    
    with st.container():
        st.markdown('<div class="paper-container">', unsafe_allow_html=True)
        # ใช้ mammoth แปลง Word เป็น HTML เพื่อโชว์บนเว็บให้ครบถ้วนที่สุด
        html_res = mammoth.convert_to_html(io.BytesIO(file_bytes))
        st.markdown(f'<div class="paper-content">{html_res.value}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)