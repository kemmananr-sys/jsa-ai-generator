import streamlit as st
import json
import io
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader

# ==========================================
# 1. ฟังก์ชันช่วยจัดรูปแบบตาราง Word (.docx)
# ==========================================
def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def format_cell_text(cell, text, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = str(text) if text is not None else ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for paragraph in cell.paragraphs:
        paragraph.alignment = align
        paragraph.paragraph_format.space_before = Pt(2)
        paragraph.paragraph_format.space_after = Pt(2)
        for run in paragraph.runs:
            run.font.name = 'TH SarabunPSK'
            run.font.size = Pt(size)
            run.font.bold = bold

# ==========================================
# 2. ฟังก์ชันสร้างไฟล์ Word รองรับทั้งภาษาไทยและอังกฤษ
# ==========================================
def create_jsa_word_file(data, project_name, doc_title, doc_ref, lang):
    doc = Document()
    
    # --- ส่วนที่ 1: ตาราง JSA 4 คอลัมน์ ---
    section1 = doc.sections[0]
    section1.orientation = WD_ORIENT.PORTRAIT
    section1.page_width = Inches(8.27)
    section1.page_height = Inches(11.69)
    section1.top_margin = Inches(0.6)
    section1.bottom_margin = Inches(0.6)
    section1.left_margin = Inches(0.6)
    section1.right_margin = Inches(0.6)

    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = h1.add_run(f"JSA\n{doc_title}")
    r1.bold = True
    r1.font.size = Pt(16)
    r1.font.name = 'TH SarabunPSK'

    if lang == "ภาษาไทย":
        headers_4col = ["ขั้นตอนการทำงาน", "แหล่งอันตราย", "ลักษณะอันตราย", "มาตรการป้องกัน"]
    else:
        headers_4col = ["Step of work", "Hazard", "Hazard identified", "Prevention measures"]

    table_jsa = doc.add_table(rows=1, cols=4)
    table_jsa.style = 'Table Grid'
    table_jsa.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, h_text in enumerate(headers_4col):
        cell = table_jsa.rows[0].cells[i]
        format_cell_text(cell, h_text, bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_bg(cell, "D9E1F2")

    for step_idx, step_item in enumerate(data.get("steps", []), start=1):
        step_name = f"{step_idx}. {step_item.get('step_name', '')}"
        hazards = step_item.get("hazards", [])
        start_row_idx = len(table_jsa.rows)

        for h_idx, h_item in enumerate(hazards, start=1):
            row_cells = table_jsa.add_row().cells
            format_cell_text(row_cells[0], step_name if h_idx == 1 else "", bold=(h_idx == 1), size=12)
            format_cell_text(row_cells[1], f"{h_idx}. {h_item.get('hazard_source', '')}", size=12)
            format_cell_text(row_cells[2], f"{h_idx}. {h_item.get('risk_impact', '')}", size=12)
            
            controls_list = h_item.get("controls", [])
            controls_str = "\n".join([f"{c_i}. {c_text}" for c_i, c_text in enumerate(controls_list, start=1)])
            format_cell_text(row_cells[3], controls_str, size=12)

        end_row_idx = len(table_jsa.rows) - 1
        if end_row_idx > start_row_idx:
            table_jsa.cell(start_row_idx, 0).merge(table_jsa.cell(end_row_idx, 0))

    # --- ส่วนที่ 2: ตาราง Risk Matrix & Assessment ---
    section2 = doc.add_section(WD_ORIENT.LANDSCAPE)
    section2.orientation = WD_ORIENT.LANDSCAPE
    section2.page_width = Inches(11.69)
    section2.page_height = Inches(8.27)
    section2.top_margin = Inches(0.5)
    section2.bottom_margin = Inches(0.5)
    section2.left_margin = Inches(0.5)
    section2.right_margin = Inches(0.5)

    p_rm = doc.add_paragraph()
    p_rm.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_rm = p_rm.add_run(f"{project_name}\nRisk Matrix - {doc_title}\nRef: {doc_ref}")
    r_rm.bold = True
    r_rm.font.size = Pt(14)
    r_rm.font.name = 'TH SarabunPSK'

    matrix_table = doc.add_table(rows=6, cols=6)
    matrix_table.style = 'Table Grid'
    matrix_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    format_cell_text(matrix_table.cell(0, 0), "Severity (S) \\ Probability (P)", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_bg(matrix_table.cell(0, 0), "D9E1F2")

    p_labels = ["P=1 (Unlikely)", "P=2 (May Happen)", "P=3 (Likely)", "P=4 (Very Likely)", "P=5 (Certain)"]
    s_labels = ["S=1 (Minor)", "S=2 (Moderate)", "S=3 (Major)", "S=4 (Catastrophic)", "S=5 (Fatalities)"]

    for p_i in range(1, 6):
        format_cell_text(matrix_table.cell(0, p_i), p_labels[p_i-1], bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_bg(matrix_table.cell(0, p_i), "D9E1F2")

    for s_i in range(1, 6):
        format_cell_text(matrix_table.cell(s_i, 0), s_labels[s_i-1], bold=True, size=11)
        set_cell_bg(matrix_table.cell(s_i, 0), "D9E1F2")
        for p_i in range(1, 6):
            r_val = s_i * p_i
            c = matrix_table.cell(s_i, p_i)
            format_cell_text(c, str(r_val), bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            if r_val <= 3: set_cell_bg(c, "C6E0B4")
            elif r_val <= 9: set_cell_bg(c, "FFE699")
            elif r_val <= 16: set_cell_bg(c, "F8CBAD")
            else: set_cell_bg(c, "FF7C80")

    doc.add_paragraph("Risk Level Criteria: 1-3 = Low (Acceptable) | 4-9 = Medium | 10-16 = High | 20-25 = Very High")

    doc.add_page_break()
    p_ra = doc.add_paragraph()
    p_ra.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ra = p_ra.add_run(f"RISK ASSESSMENT TABLE\n{project_name}\n{doc_title}")
    r_ra.bold = True
    r_ra.font.size = Pt(14)
    r_ra.font.name = 'TH SarabunPSK'

    ra_headers = [
        "No.", "Activity (Description)", "Hazard / Aspect", "Risk / Impact",
        "P (Before)", "S (Before)", "R=SxP",
        "Existing Controls (OHS & E)",
        "P (After)", "S (After)", "R=SxP",
        "Legal Reqd. (Y/N)", "Sign. Resp. Person"
    ]

    table_ra = doc.add_table(rows=1, cols=13)
    table_ra.style = 'Table Grid'
    table_ra.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, h_text in enumerate(ra_headers):
        cell = table_ra.rows[0].cells[i]
        format_cell_text(cell, h_text, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_bg(cell, "D9E1F2")

    for step_idx, step_item in enumerate(data.get("steps", []), start=1):
        step_name = step_item.get("step_name", "")
        hazards = step_item.get("hazards", [])
        start_row_ra = len(table_ra.rows)

        for h_idx, h_item in enumerate(hazards, start=1):
            row_c = table_ra.add_row().cells
            p_init = int(h_item.get("p_before", 2))
            s_init = int(h_item.get("s_before", 2))
            r_init = p_init * s_init

            p_res = int(h_item.get("p_after", 1))
            s_res = int(h_item.get("s_after", 2))
            r_res = p_res * s_res

            controls_list = h_item.get("controls", [])
            controls_str = "\n".join([f"{c_i}. {c_text}" for c_i, c_text in enumerate(controls_list, start=1)])

            format_cell_text(row_c[0], str(step_idx) if h_idx == 1 else "", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[1], step_name if h_idx == 1 else "", bold=(h_idx == 1), size=11)
            format_cell_text(row_c[2], h_item.get("hazard_source", ""), size=11)
            format_cell_text(row_c[3], h_item.get("risk_impact", ""), size=11)
            format_cell_text(row_c[4], str(p_init), size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[5], str(s_init), size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[6], str(r_init), bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[7], controls_str, size=11)
            format_cell_text(row_c[8], str(p_res), size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[9], str(s_res), size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[10], str(r_res), bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[11], h_item.get("legal_req", "N"), size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell_text(row_c[12], h_item.get("resp_person", "ER, FM, ST"), size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

        end_row_ra = len(table_ra.rows) - 1
        if end_row_ra > start_row_ra:
            table_ra.cell(start_row_ra, 0).merge(table_ra.cell(end_row_ra, 0))
            table_ra.cell(start_row_ra, 1).merge(table_ra.cell(end_row_ra, 1))

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ==========================================
# 3. ระบบคำสั่ง AI ที่รองรับทั้งภาษาไทยและอังกฤษตามผู้ใช้เลือก
# ==========================================
def get_system_prompt(lang):
    if lang == "English":
        return """
        You are a Senior Safety Officer for construction and engineering projects. 
        Your task is to read the provided Method Statement (Section 7) and generate a professional JSA and Risk Assessment table in ENGLISH.
        
        [Rules]:
        1. Steps must start with material/machinery preparation and site inspection, followed sequentially to site handover.
        2. Analyze 2-4 hazards per step covering machinery, environment, and operations.
        3. Prevention measures must be numbered lists (3-5 items) using standard safety terms (e.g., Toolbox talk, regular inspection, Tag Out "Do Not Use", proper PPE, flagman).
        4. Rating: Initial P(2-4) x S(2-4), Residual P(1-2) x S(1-2), Risk Value R = P x S.
        5. Responsible persons: "ER, FM, ST".
        
        Respond strictly in valid JSON format:
        {
          "steps": [
            {
              "step_name": "Step description in English",
              "hazards": [
                {
                  "hazard_source": "Hazard source",
                  "risk_impact": "Risk / Impact description",
                  "p_before": 2,
                  "s_before": 3,
                  "controls": ["Control 1", "Control 2", "Control 3"],
                  "p_after": 1,
                  "s_after": 2,
                  "legal_req": "N",
                  "resp_person": "ER, FM, ST"
                }
              ]
            }
          ]
        }
        """
    else:
        return """
        คุณคือ จป.วิชาชีพอาวุโส หน้าที่ของคุณคืออ่าน Method Statement แล้วแปลงเป็นตาราง JSA และ Risk Assessment เป็นภาษาไทย
        ตอบกลับเป็น JSON ตามโครงสร้างนี้เท่านั้น:
        {
          "steps": [
            {
              "step_name": "ชื่อขั้นตอนการทำงาน",
              "hazards": [
                {
                  "hazard_source": "แหล่งอันตราย",
                  "risk_impact": "ลักษณะอันตราย",
                  "p_before": 2,
                  "s_before": 3,
                  "controls": ["มาตรการ 1", "มาตรการ 2"],
                  "p_after": 1,
                  "s_after": 2,
                  "legal_req": "N",
                  "resp_person": "ER, FM, ST"
                }
              ]
            }
          ]
        }
        """

# ==========================================
# 4. หน้าจอเว็บไซต์ (Streamlit UI)
# ==========================================
st.set_page_config(page_title="Safety AI - JSA & Risk Assessment Generator", layout="wide")

st.title("👷‍♂️ ระบบ AI สร้างตาราง JSA & Risk Assessment อัตโนมัติ")
st.caption("รองรับทั้งภาษาไทยและภาษาอังกฤษ พร้อมออกไฟล์ Word ตามฟอร์มบริษัท")

with st.sidebar:
    st.header("⚙️ ตั้งค่าระบบ")
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    lang_choice = st.radio("🌐 เลือกภาษาของผลลัพธ์", ["ภาษาไทย", "English"])
    
    st.divider()
    st.subheader("📄 ข้อมูลหัวกระดาษ")
    project_name = st.text_input("ชื่อโครงการ", value="PROJECT NAME / โครงการระบบรถไฟชานเมือง")
    doc_title = st.text_input("ชื่องาน", value="Method Statement for Pile Repair Works")
    doc_ref = st.text_input("รหัสเอกสาร", value="UN-DRL/ACW/MST/...")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1️⃣ นำเข้าข้อมูล Method Statement")
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF หรือ Word", type=["pdf", "docx"])
    
    extracted_text = ""
    if uploaded_file is not None:
        if uploaded_file.name.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            for page in reader.pages:
                extracted_text += page.extract_text() + "\n"
        elif uploaded_file.name.endswith(".docx"):
            doc_in = Document(uploaded_file)
            for para in doc_in.paragraphs:
                extracted_text += para.text + "\n"
        st.success(f"ดึงข้อความสำเร็จ!")

    method_input = st.text_area("หรือวางข้อความขั้นตอนการทำงานที่นี่:", value=extracted_text, height=320)
    generate_btn = st.button("🚀 ให้ AI วิเคราะห์และสร้าง JSA + Risk Matrix", type="primary", use_container_width=True)

if generate_btn:
    if not api_key:
        st.error("⚠️ กรุณาใส่ Google Gemini API Key ด้านซ้าย")
    elif not method_input.strip():
        st.warning("⚠️ กรุณาวางข้อความขั้นตอนการทำงาน")
    else:
        with st.spinner("🤖 AI กำลังประมวลผล (หากข้อมูลเยอะอาจใช้เวลา 10-20 วินาที)..."):
            try:
                client = genai.Client(api_key=api_key)
                user_prompt = f"""
                Target Language: {lang_choice}
                Document Title: {doc_title}
                Method Statement Content:
                {method_input}
                """
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=user_prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=get_system_prompt(lang_choice),
                        response_mime_type="application/json",
                        temperature=0.2,
                    ),
                )
                st.session_state["jsa_result"] = json.loads(response.text)
                st.success("✅ สำเร็จ! ตรวจสอบตารางด้านขวาได้เลย")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")

with col2:
    st.subheader("2️⃣ ตรวจสอบและดาวน์โหลดไฟล์ Word")
    if "jsa_result" in st.session_state:
        result_data = st.session_state["jsa_result"]
        word_file = create_jsa_word_file(result_data, project_name, doc_title, doc_ref, lang_choice)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Word (.docx) ฉบับสมบูรณ์",
            data=word_file,
            file_name=f"JSA_{doc_title.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

        preview_rows = []
        for s_i, step in enumerate(result_data.get("steps", []), start=1):
            for h_i, haz in enumerate(step.get("hazards", []), start=1):
                p_b = int(haz.get("p_before", 2))
                s_b = int(haz.get("s_before", 2))
                preview_rows.append({
                    "No.": f"{s_i}.{h_i}",
                    "Activity": step.get("step_name", ""),
                    "Hazard": haz.get("hazard_source", ""),
                    "Risk / Impact": haz.get("risk_impact", ""),
                    "R (Before)": f"{p_b}x{s_b}={p_b*s_b}"
                })
        st.dataframe(preview_rows, use_container_width=True, height=420)
    else:
        st.info("👈 กรอกข้อมูลและกดปุ่มสร้างรายงานเพื่อดูผลลัพธ์ที่นี่")