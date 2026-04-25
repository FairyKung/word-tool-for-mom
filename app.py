import streamlit as st
from docx import Document
from docx.shared import Pt
import io
import re
import mammoth

# ===== Page Config =====
st.set_page_config(
    page_title="แก้คำใน Word",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===== Custom CSS =====
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif !important;
}

/* ── Background ── */
.stApp {
    background: linear-gradient(135deg, #f0f4ff 0%, #faf5ff 100%);
}

/* ── Main title ── */
h1 {
    font-size: 2.2rem !important;
    font-weight: 700 !important;
    color: #2d2d6b !important;
    text-align: center;
    padding: 0.5rem 0 1.5rem 0;
}

/* ── Section headers ── */
h2 {
    font-size: 1.4rem !important;
    color: #3b3b8f !important;
    font-weight: 600 !important;
    border-left: 5px solid #6c63ff;
    padding-left: 12px;
    margin-top: 1.5rem !important;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #2d2d6b 0%, #3b3b8f 100%) !important;
}
[data-testid="stSidebar"] * {
    color: white !important;
}
[data-testid="stSidebar"] .stTextInput input {
    background: rgba(255,255,255,0.15) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.4) !important;
    border-radius: 8px !important;
    font-size: 1rem !important;
    padding: 0.5rem 0.7rem !important;
}
[data-testid="stSidebar"] .stTextInput input::placeholder {
    color: rgba(255,255,255,0.5) !important;
}
[data-testid="stSidebar"] label {
    font-size: 1rem !important;
    font-weight: 600 !important;
}

/* ── Buttons ── */
.stButton > button {
    border-radius: 12px !important;
    font-family: 'Sarabun', sans-serif !important;
    font-size: 1.05rem !important;
    font-weight: 600 !important;
    padding: 0.55rem 1.2rem !important;
    border: none !important;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #6c63ff, #4e46d4) !important;
    color: white !important;
    width: 100% !important;
    padding: 0.7rem !important;
    font-size: 1.1rem !important;
    box-shadow: 0 4px 14px rgba(108,99,255,0.35) !important;
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 18px rgba(108,99,255,0.45) !important;
}
.stButton > button:not([kind="primary"]) {
    background: rgba(255,255,255,0.25) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.4) !important;
}

/* ── File Uploader ── */
[data-testid="stFileUploader"] {
    background: white !important;
    border: 2.5px dashed #6c63ff !important;
    border-radius: 16px !important;
    padding: 1.5rem !important;
    text-align: center !important;
}
[data-testid="stFileUploader"] label {
    font-size: 1.1rem !important;
    color: #3b3b8f !important;
    font-weight: 600 !important;
}

/* ── Paper Preview ── */
.paper-wrap {
    background: #e8eaf6;
    border-radius: 16px;
    padding: 24px;
    box-shadow: inset 0 2px 8px rgba(0,0,0,0.08);
}
.paper-page {
    background: white;
    border-radius: 8px;
    padding: 48px 52px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.12);
    min-height: 600px;
    max-width: 820px;
    margin: 0 auto;
    font-family: 'Sarabun', sans-serif !important;
    font-size: 16px;
    line-height: 1.9;
    color: #111;
}

/* ── Download button ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #22c55e, #16a34a) !important;
    color: white !important;
    font-size: 1.15rem !important;
    padding: 0.7rem 1.5rem !important;
    border-radius: 12px !important;
    font-family: 'Sarabun', sans-serif !important;
    font-weight: 700 !important;
    width: 100% !important;
    box-shadow: 0 4px 14px rgba(34,197,94,0.3) !important;
    transition: all 0.2s ease !important;
}
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 20px rgba(34,197,94,0.4) !important;
}

/* ── Tab styling ── */
[data-testid="stTabs"] [role="tablist"] {
    background: white;
    border-radius: 12px;
    padding: 4px;
    gap: 4px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}
[data-testid="stTabs"] [role="tab"] {
    font-family: 'Sarabun', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    padding: 0.5rem 1.2rem !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    background: #6c63ff !important;
    color: white !important;
}

/* ── Info / success boxes ── */
.stAlert {
    border-radius: 12px !important;
    font-size: 1rem !important;
}

/* ── Replace card ── */
.replace-card {
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 12px;
    padding: 14px 14px 4px 14px;
    margin-bottom: 12px;
}

/* ── Badge ── */
.badge {
    display: inline-block;
    background: #6c63ff;
    color: white;
    font-size: 0.75rem;
    font-weight: 700;
    padding: 2px 8px;
    border-radius: 20px;
    margin-bottom: 6px;
}

/* ── Empty state ── */
.empty-state {
    text-align: center;
    padding: 60px 20px;
    color: #9ca3af;
}
.empty-state .icon { font-size: 4rem; }
.empty-state p { font-size: 1.1rem; margin-top: 12px; }
</style>
""", unsafe_allow_html=True)


# ===== Session State =====
if 'num_pairs' not in st.session_state:
    st.session_state.num_pairs = 3
if 'processed_doc' not in st.session_state:
    st.session_state.processed_doc = None
if 'replace_count' not in st.session_state:
    st.session_state.replace_count = 0


# ===== Helper: Replace function =====
def super_replace(paragraphs, old_text, new_text):
    """แก้คำโดยรักษาฟอนต์เดิม รองรับช่องว่างไม่เท่ากัน"""
    pattern = re.escape(old_text).replace(r'\ ', r'\s+')
    count = 0
    for p in paragraphs:
        if re.search(pattern, p.text, flags=re.IGNORECASE):
            font_name = font_size = is_bold = None
            if p.runs:
                font_name = p.runs[0].font.name
                font_size = p.runs[0].font.size
                is_bold   = p.runs[0].bold
            p.text = re.sub(pattern, new_text, p.text, flags=re.IGNORECASE)
            if p.runs:
                for run in p.runs:
                    if font_name: run.font.name = font_name
                    if font_size: run.font.size = font_size
                    if is_bold is not None: run.bold = is_bold
            count += 1
    return count


def process_document(file_bytes, replace_list):
    doc = Document(io.BytesIO(file_bytes))
    total = 0
    for old, new in replace_list:
        if not old.strip():
            continue
        total += super_replace(doc.paragraphs, old, new)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    total += super_replace(cell.paragraphs, old, new)
        for section in doc.sections:
            for hf in [section.header, section.footer,
                       section.first_page_header, section.first_page_footer,
                       section.even_page_header, section.even_page_footer]:
                if hf:
                    total += super_replace(hf.paragraphs, old, new)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), total


# ===== TITLE =====
st.markdown("# 📝 โปรแกรมแก้คำใน Word")
st.markdown(
    "<p style='text-align:center;color:#6b7280;font-size:1.05rem;margin-top:-12px;margin-bottom:24px;'>"
    "อัปโหลดไฟล์ · ใส่คำที่ต้องการเปลี่ยน · ดาวน์โหลดไฟล์ใหม่ได้เลย</p>",
    unsafe_allow_html=True
)

# ===== FILE UPLOAD (main area) =====
uploaded_file = st.file_uploader(
    "📂  เลือกหรือลากไฟล์ Word มาวางที่นี่",
    type=["docx"],
    help="รองรับไฟล์ .docx เท่านั้น"
)

# ===== SIDEBAR =====
with st.sidebar:
    st.markdown("## 🔄 รายการคำที่ต้องการเปลี่ยน")
    st.markdown(
        "<p style='font-size:0.9rem;opacity:0.8;margin-top:-8px;'>ใส่คำเดิมและคำใหม่ที่ต้องการแทน</p>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    replace_list = []
    for i in range(st.session_state.num_pairs):
        st.markdown(f'<div class="replace-card"><span class="badge">คู่ที่ {i+1}</span>', unsafe_allow_html=True)
        old = st.text_input(f"คำเดิม", key=f"old_{i}",
                            placeholder="เช่น 1 เมษายน 2568",
                            label_visibility="visible")
        new = st.text_input(f"คำใหม่", key=f"new_{i}",
                            placeholder="เช่น 1 พฤษภาคม 2568",
                            label_visibility="visible")
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

    # Process button
    process_btn = st.button(
        "🚀 เริ่มเปลี่ยนคำ",
        type="primary",
        disabled=(uploaded_file is None or len(replace_list) == 0)
    )
    if uploaded_file is None:
        st.caption("⬆️ กรุณาอัปโหลดไฟล์ก่อน")
    elif len(replace_list) == 0:
        st.caption("⬆️ กรุณาใส่คำที่ต้องการเปลี่ยนก่อน")

    # Summary of pairs
    if replace_list:
        st.markdown("---")
        st.markdown("**📋 สรุปรายการที่จะเปลี่ยน:**")
        for old, new in replace_list:
            arrow = "→" if new else "→ (ลบคำ)"
            disp_new = f'`{new}`' if new else "_ลบออก_"
            st.markdown(f"• `{old}` → {disp_new}")


# ===== MAIN CONTENT =====
if uploaded_file is not None:
    file_content = uploaded_file.getvalue()

    # Process on button click
    if process_btn and replace_list:
        with st.spinner("⏳ กำลังแก้ไขไฟล์... รอสักครู่นะคะ"):
            result_bytes, count = process_document(file_content, replace_list)
            st.session_state.processed_doc = result_bytes
            st.session_state.replace_count = count
        st.success(f"✅ แก้ไขเรียบร้อยแล้ว! พบและเปลี่ยนใน {count} ย่อหน้า")

    # Download button (if ready)
    if st.session_state.processed_doc:
        st.download_button(
            label="📥  ดาวน์โหลดไฟล์ที่แก้เสร็จแล้ว",
            data=st.session_state.processed_doc,
            file_name=f"แก้แล้ว_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.markdown("---")

    # Tabs
    if st.session_state.processed_doc:
        tab1, tab2 = st.tabs(["📄 ต้นฉบับ", "✨ หลังแก้ไข"])
    else:
        tab1, = st.tabs(["📄 ต้นฉบับ"])
        tab2 = None

    with tab1:
        st.markdown("## 📄 ไฟล์ต้นฉบับ")
        try:
            html_orig = mammoth.convert_to_html(io.BytesIO(file_content)).value
            st.markdown(
                f'<div class="paper-wrap"><div class="paper-page">{html_orig}</div></div>',
                unsafe_allow_html=True
            )
        except Exception as e:
            st.error(f"ไม่สามารถแสดง Preview ได้: {e}")

    if tab2 and st.session_state.processed_doc:
        with tab2:
            st.markdown("## ✨ ไฟล์หลังแก้ไข")
            # Highlight summary
            if replace_list:
                changes = " · ".join([f"`{o}` → `{n}`" if n else f"`{o}` → _(ลบ)_" for o, n in replace_list])
                st.info(f"🔄 เปลี่ยนแล้ว: {changes}")
            try:
                html_fixed = mammoth.convert_to_html(io.BytesIO(st.session_state.processed_doc)).value
                st.markdown(
                    f'<div class="paper-wrap"><div class="paper-page">{html_fixed}</div></div>',
                    unsafe_allow_html=True
                )
            except Exception as e:
                st.error(f"ไม่สามารถแสดง Preview ได้: {e}")

else:
    # Empty state
    st.markdown("""
    <div class="empty-state">
        <div class="icon">📂</div>
        <p>ยังไม่ได้เลือกไฟล์<br>กรุณาอัปโหลดไฟล์ Word (.docx) ด้านบนก่อนนะคะ</p>
    </div>
    """, unsafe_allow_html=True)

    # How-to guide
    st.markdown("---")
    st.markdown("## 📖 วิธีใช้งาน")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        ### 1️⃣ อัปโหลดไฟล์
        คลิกที่กล่องด้านบน แล้วเลือกไฟล์ Word ที่ต้องการแก้ไข
        """)
    with col2:
        st.markdown("""
        ### 2️⃣ ใส่คำที่ต้องการเปลี่ยน
        ใส่คำเดิมและคำใหม่ในแถบด้านซ้าย สามารถเพิ่มได้หลายคู่
        """)
    with col3:
        st.markdown("""
        ### 3️⃣ ดาวน์โหลด
        กด **เริ่มเปลี่ยนคำ** แล้วดาวน์โหลดไฟล์ที่แก้เสร็จแล้ว
        """)