import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
import copy
import io
import re
import mammoth

st.set_page_config(page_title="แก้คำใน Word", page_icon="📝", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Sarabun', sans-serif !important; }
.stApp { background: linear-gradient(135deg, #f0f4ff 0%, #faf5ff 100%); }
h1 { font-size: 2.2rem !important; font-weight: 700 !important; color: #2d2d6b !important; text-align: center; padding: 0.5rem 0 1.5rem 0; }
h2 { font-size: 1.4rem !important; color: #3b3b8f !important; font-weight: 600 !important; border-left: 5px solid #6c63ff; padding-left: 12px; margin-top: 1.5rem !important; }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #2d2d6b 0%, #3b3b8f 100%) !important; }
[data-testid="stSidebar"] * { color: white !important; }
[data-testid="stSidebar"] .stTextInput input { background: rgba(255,255,255,0.15) !important; color: white !important; border: 1px solid rgba(255,255,255,0.4) !important; border-radius: 8px !important; font-size: 1rem !important; padding: 0.5rem 0.7rem !important; }
[data-testid="stSidebar"] .stTextInput input::placeholder { color: rgba(255,255,255,0.5) !important; }
[data-testid="stSidebar"] label { font-size: 1rem !important; font-weight: 600 !important; }
.stButton > button { border-radius: 12px !important; font-family: 'Sarabun', sans-serif !important; font-size: 1.05rem !important; font-weight: 600 !important; padding: 0.55rem 1.2rem !important; border: none !important; transition: all 0.2s ease !important; }
.stButton > button[kind="primary"] { background: linear-gradient(135deg, #6c63ff, #4e46d4) !important; color: white !important; width: 100% !important; padding: 0.7rem !important; font-size: 1.1rem !important; box-shadow: 0 4px 14px rgba(108,99,255,0.35) !important; }
.stButton > button:not([kind="primary"]) { background: rgba(255,255,255,0.25) !important; color: white !important; border: 1px solid rgba(255,255,255,0.4) !important; }
[data-testid="stFileUploader"] { background: white !important; border: 2.5px dashed #6c63ff !important; border-radius: 16px !important; padding: 1.5rem !important; }
.paper-wrap { background: #e8eaf6; border-radius: 16px; padding: 24px; box-shadow: inset 0 2px 8px rgba(0,0,0,0.08); }
.paper-page { background: white; border-radius: 8px; padding: 48px 52px; box-shadow: 0 4px 20px rgba(0,0,0,0.12); min-height: 600px; max-width: 820px; margin: 0 auto; font-family: 'Sarabun', sans-serif !important; font-size: 16px; line-height: 1.9; color: #111; }
[data-testid="stDownloadButton"] > button { background: linear-gradient(135deg, #22c55e, #16a34a) !important; color: white !important; font-size: 1.15rem !important; padding: 0.7rem 1.5rem !important; border-radius: 12px !important; font-family: 'Sarabun', sans-serif !important; font-weight: 700 !important; width: 100% !important; box-shadow: 0 4px 14px rgba(34,197,94,0.3) !important; }
[data-testid="stTabs"] [role="tablist"] { background: white; border-radius: 12px; padding: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
[data-testid="stTabs"] [role="tab"] { font-family: 'Sarabun', sans-serif !important; font-size: 1rem !important; font-weight: 600 !important; border-radius: 10px !important; padding: 0.5rem 1.2rem !important; }
[data-testid="stTabs"] [role="tab"][aria-selected="true"] { background: #6c63ff !important; color: white !important; }
.replace-card { background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.25); border-radius: 12px; padding: 14px 14px 4px 14px; margin-bottom: 12px; }
.badge { display: inline-block; background: #6c63ff; color: white; font-size: 0.75rem; font-weight: 700; padding: 2px 8px; border-radius: 20px; margin-bottom: 6px; }
.empty-state { text-align: center; padding: 60px 20px; color: #9ca3af; }
.empty-state .icon { font-size: 4rem; }
.empty-state p { font-size: 1.1rem; margin-top: 12px; }
</style>
""", unsafe_allow_html=True)

if 'num_pairs' not in st.session_state:
    st.session_state.num_pairs = 3
if 'processed_doc' not in st.session_state:
    st.session_state.processed_doc = None


# ============================================================
#  ENGINE: แก้ปัญหาทุกอย่างจากไฟล์จริง
#
#  ปัญหาที่พบใน docx นี้:
#  1. คำถูกหั่นข้าม run หลายอัน เช่น "10" | "  เมษายน" | "  2569"
#  2. บาง run มี font.size = None (inherit จาก style) — ต้องเก็บไว้เป็น None
#     ไม่ใช่ copy ค่าจาก run อื่น
#  3. rPr (run properties XML) ต้องถูก clone เต็ม ๆ ไม่ใช่แค่ font
# ============================================================

def _clone_rpr(run_elem):
    """Clone <w:rPr> element ทั้งหมดจาก run XML (วิธีที่ถูกที่สุด)"""
    rpr = run_elem.find(qn('w:rPr'))
    if rpr is not None:
        return copy.deepcopy(rpr)
    return None


def _make_run_with_rpr(paragraph, text, rpr_clone):
    """สร้าง run ใหม่พร้อม rPr ที่ clone มา (รักษา font/size/bold/color ทั้งหมด)"""
    new_run = paragraph.add_run(text)
    # ลบ rPr เดิมออกก่อน แล้วใส่ rPr ที่ clone มาแทน
    existing_rpr = new_run._r.find(qn('w:rPr'))
    if existing_rpr is not None:
        new_run._r.remove(existing_rpr)
    if rpr_clone is not None:
        new_run._r.insert(0, copy.deepcopy(rpr_clone))
    return new_run


def _replace_in_paragraph(paragraph, old_text, new_text):
    """
    แทนที่คำโดย:
    1. รวม text ทุก run + เก็บ map ว่าตัวอักษร i อยู่ใน run index ไหน
    2. regex replace บน full_text
    3. rebuild runs โดย clone <w:rPr> เต็ม ๆ จาก run เดิม
    """
    runs = paragraph.runs
    if not runs:
        return 0

    # เก็บ rPr clone ของทุก run
    rpr_list = [_clone_rpr(r._r) for r in runs]

    # สร้าง full_text + char→run map
    full_text = ""
    char_to_run = []
    for r_idx, run in enumerate(runs):
        for _ in run.text:
            char_to_run.append(r_idx)
        full_text += run.text

    pattern = re.escape(old_text).replace(r'\ ', r'\s+')
    matches = list(re.finditer(pattern, full_text))
    if not matches:
        return 0

    # สร้าง segments: list of (text, run_index_สำหรับ_font)
    segments = []
    cursor = 0
    for m in matches:
        # ส่วนก่อน match
        for i in range(cursor, m.start()):
            r_idx = char_to_run[i]
            if segments and segments[-1][1] == r_idx:
                segments[-1] = (segments[-1][0] + full_text[i], r_idx)
            else:
                segments.append((full_text[i], r_idx))
        # ส่วนที่แทนที่ — ใช้ rPr ของ run แรกของ match
        first_r_idx = char_to_run[m.start()] if m.start() < len(char_to_run) else 0
        if new_text:
            if segments and segments[-1][1] == first_r_idx:
                segments[-1] = (segments[-1][0] + new_text, first_r_idx)
            else:
                segments.append((new_text, first_r_idx))
        cursor = m.end()

    # ส่วนท้าย
    for i in range(cursor, len(full_text)):
        r_idx = char_to_run[i] if i < len(char_to_run) else len(runs) - 1
        if segments and segments[-1][1] == r_idx:
            segments[-1] = (segments[-1][0] + full_text[i], r_idx)
        else:
            segments.append((full_text[i], r_idx))

    # ลบ runs เดิมออก
    p_elem = paragraph._p
    for run in runs:
        p_elem.remove(run._r)

    # สร้าง runs ใหม่ด้วย rPr ที่ clone มา (รักษาทุก property เป๊ะ)
    for seg_text, r_idx in segments:
        if not seg_text:
            continue
        _make_run_with_rpr(paragraph, seg_text, rpr_list[r_idx])

    return len(matches)


def replace_in_paragraphs(paragraphs, old_text, new_text):
    count = 0
    for p in paragraphs:
        count += _replace_in_paragraph(p, old_text, new_text)
    return count


def process_document(file_bytes, replace_list):
    doc = Document(io.BytesIO(file_bytes))
    total = 0
    for old, new in replace_list:
        if not old.strip():
            continue
        total += replace_in_paragraphs(doc.paragraphs, old, new)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    total += replace_in_paragraphs(cell.paragraphs, old, new)
        for section in doc.sections:
            for hf in [section.header, section.footer,
                       section.first_page_header, section.first_page_footer,
                       section.even_page_header, section.even_page_footer]:
                if hf:
                    total += replace_in_paragraphs(hf.paragraphs, old, new)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), total


# ===== UI =====
st.markdown("# 📝 โปรแกรมแก้คำใน Word")
st.markdown("<p style='text-align:center;color:#6b7280;font-size:1.05rem;margin-top:-12px;margin-bottom:24px;'>อัปโหลดไฟล์ · ใส่คำที่ต้องการเปลี่ยน · ดาวน์โหลดไฟล์ใหม่ได้เลย</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂  เลือกหรือลากไฟล์ Word มาวางที่นี่", type=["docx"], help="รองรับไฟล์ .docx เท่านั้น")

with st.sidebar:
    st.markdown("## 🔄 รายการคำที่ต้องการเปลี่ยน")
    st.markdown("<p style='font-size:0.9rem;opacity:0.8;margin-top:-8px;'>ใส่คำเดิมและคำใหม่ที่ต้องการแทน</p>", unsafe_allow_html=True)
    st.markdown("---")

    replace_list = []
    for i in range(st.session_state.num_pairs):
        st.markdown(f'<div class="replace-card"><span class="badge">คู่ที่ {i+1}</span>', unsafe_allow_html=True)
        old = st.text_input("คำเดิม", key=f"old_{i}", placeholder="เช่น 1 เมษายน 2568")
        new = st.text_input("คำใหม่", key=f"new_{i}", placeholder="เช่น 1 พฤษภาคม 2568")
        st.markdown('</div>', unsafe_allow_html=True)
        if old.strip():
            replace_list.append((old.strip(), new.strip()))

    col_add, col_reset = st.columns(2)
    with col_add:
        if st.button("➕ เพิ่มช่อง"):
            st.session_state.num_pairs += 1
            st.rerun()
    with col_reset:
        if st.button("🗑️ ล้างทั้งหมด"):
            st.session_state.num_pairs = 3
            for k in list(st.session_state.keys()):
                if k.startswith("old_") or k.startswith("new_"):
                    del st.session_state[k]
            st.session_state.processed_doc = None
            st.rerun()

    st.markdown("---")
    process_btn = st.button("🚀 เริ่มเปลี่ยนคำ", type="primary",
                            disabled=(uploaded_file is None or len(replace_list) == 0))
    if uploaded_file is None:
        st.caption("⬆️ กรุณาอัปโหลดไฟล์ก่อน")
    elif len(replace_list) == 0:
        st.caption("⬆️ กรุณาใส่คำที่ต้องการเปลี่ยนก่อน")

    if replace_list:
        st.markdown("---")
        st.markdown("**📋 สรุปรายการที่จะเปลี่ยน:**")
        for old, new in replace_list:
            disp_new = f'`{new}`' if new else "_ลบออก_"
            st.markdown(f"• `{old}` → {disp_new}")

if uploaded_file is not None:
    file_content = uploaded_file.getvalue()

    if process_btn and replace_list:
        with st.spinner("⏳ กำลังแก้ไขไฟล์... รอสักครู่นะคะ"):
            result_bytes, count = process_document(file_content, replace_list)
            st.session_state.processed_doc = result_bytes
        st.success(f"✅ แก้ไขเรียบร้อยแล้ว! พบและเปลี่ยนใน {count} ตำแหน่ง")

    if st.session_state.processed_doc:
        st.download_button(label="📥  ดาวน์โหลดไฟล์ที่แก้เสร็จแล้ว",
                           data=st.session_state.processed_doc,
                           file_name=f"แก้แล้ว_{uploaded_file.name}",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.markdown("---")

    if st.session_state.processed_doc:
        tab1, tab2 = st.tabs(["📄 ต้นฉบับ", "✨ หลังแก้ไข"])
    else:
        tab1, = st.tabs(["📄 ต้นฉบับ"])
        tab2 = None

    with tab1:
        st.markdown("## 📄 ไฟล์ต้นฉบับ")
        try:
            html_orig = mammoth.convert_to_html(io.BytesIO(file_content)).value
            st.markdown(f'<div class="paper-wrap"><div class="paper-page">{html_orig}</div></div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"ไม่สามารถแสดง Preview ได้: {e}")

    if tab2 and st.session_state.processed_doc:
        with tab2:
            st.markdown("## ✨ ไฟล์หลังแก้ไข")
            if replace_list:
                changes = " · ".join([f"`{o}` → `{n}`" if n else f"`{o}` → _(ลบ)_" for o, n in replace_list])
                st.info(f"🔄 เปลี่ยนแล้ว: {changes}")
            try:
                html_fixed = mammoth.convert_to_html(io.BytesIO(st.session_state.processed_doc)).value
                st.markdown(f'<div class="paper-wrap"><div class="paper-page">{html_fixed}</div></div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(f"ไม่สามารถแสดง Preview ได้: {e}")
else:
    st.markdown('<div class="empty-state"><div class="icon">📂</div><p>ยังไม่ได้เลือกไฟล์<br>กรุณาอัปโหลดไฟล์ Word (.docx) ด้านบนก่อนนะคะ</p></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("## 📖 วิธีใช้งาน")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("### 1️⃣ อัปโหลดไฟล์\nคลิกที่กล่องด้านบน แล้วเลือกไฟล์ Word ที่ต้องการแก้ไข")
    with col2:
        st.markdown("### 2️⃣ ใส่คำที่ต้องการเปลี่ยน\nใส่คำเดิมและคำใหม่ในแถบด้านซ้าย สามารถเพิ่มได้หลายคู่")
    with col3:
        st.markdown("### 3️⃣ ดาวน์โหลด\nกด **เริ่มเปลี่ยนคำ** แล้วดาวน์โหลดไฟล์ที่แก้เสร็จแล้ว")