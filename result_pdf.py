import streamlit as st
import pdfplumber
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO
import tempfile
import matplotlib.pyplot as plt
from fpdf import FPDF

grade_list = ['O', 'A+', 'A', 'B+', 'B', 'C', 'D', 'F', 'SA']
grade_set = set(grade_list)
pass_grades = set(['O', 'A+', 'A', 'B+', 'B', 'C', 'D', 'SA'])
absent_labels = set(['Ab', 'AB', 'ab', 'ABSENT', 'Absent', '-', ''])

def parse_subject_chunk(chunk):
    parts = chunk.strip().split()
    if not parts or (not parts[0].startswith('MS-') and not parts[0].startswith('PGA')):
        return None
    match = next((i for i, p in enumerate(parts) if re.match(r'^-?[\d\*]+$', p) or p == '-'), None)
    if match is None or match < 2:
        return None
    code = parts[0]
    name = " ".join(parts[1:match])
    scores = []
    idx = match
    while idx < len(parts) and (re.match(r'^-?[\d\*]+$', parts[idx]) or parts[idx] == '-'):
        scores.append(parts[idx])
        idx += 1
    grade = None
    while idx < len(parts):
        if parts[idx] in grade_set or parts[idx] in absent_labels:
            grade = parts[idx]
            idx += 1
            break
        idx += 1
    marks = None
    if scores:
        try:
            marks = float(scores[-1].replace('*',''))
        except:
            marks = None
    return {'code': code, 'name': name, 'marks': marks, 'grade': grade}

@st.cache_data(show_spinner=True)
def analyze_pdf(file):
    subject_studentdata = defaultdict(list)
    all_students = set()
    with pdfplumber.open(file) as pdf:
        current_student = None
        for page in pdf.pages:
            lines = (page.extract_text() or '').split('\n')
            for line in lines:
                if 'Name :' in line and 'Mother' not in line:
                    name_match = re.search(
                        r'Name\s*:\s*([A-Za-z\s\.\'\-]+?)(?:\s+Eligibility|$)', line)
                    if name_match:
                        current_student = name_match.group(1).strip().title()
                        all_students.add(current_student)
                if not current_student:
                    continue
                if 'MS-' in line or 'PGA' in line:
                    found = [m.start() for m in re.finditer(r'(?:MS-|PGA)', line)]
                    halves = []
                    if len(found) >= 2:
                        halves = [line[found[0]:found[1]].strip(), line[found[1]:].strip()]
                    elif len(found) == 1:
                        halves = [line[found[0]:].strip()]
                    else:
                        halves = [line.strip()]
                    for chunk in halves:
                        subinfo = parse_subject_chunk(chunk)
                        if subinfo and subinfo["grade"]:
                            k = (subinfo["code"], subinfo["name"])
                            entry = {
                                "marks": subinfo["marks"],
                                "grade": subinfo["grade"],
                                "student": current_student
                            }
                            subject_studentdata[k].append(entry)
    result_rows = []
    topper_blocks = []
    total_passed_students = set()
    for sno, (k, entries) in enumerate(subject_studentdata.items(), 1):
        grade_counts = {gr: 0 for gr in grade_list}
        present, absent, passed = 0, 0, 0
        topper_list = []
        for d in entries:
            grade = d.get('grade')
            marks = d.get('marks')
            student = d.get('student', '')
            if not grade:
                continue
            if grade in absent_labels:
                absent += 1
            else:
                present += 1
                grade_counts[grade] += 1
                if grade in pass_grades:
                    passed += 1
                    total_passed_students.add(student)
                if marks is not None and grade not in absent_labels:
                    topper_list.append({"student": student, "marks": marks, "grade": grade})
        total_present = present
        total_absent = absent
        perc_passed = round((passed/total_present)*100, 2) if total_present else 0
        row = {
            "S.No.": sno,
            "Subject Code": k[0],
            "Subject Name": k[1],
            **grade_counts,
            "total Present": total_present,
            "Absent": total_absent,
            "Pass": passed,
            "Percentage passed": perc_passed
        }
        result_rows.append(row)
        top_entries = sorted(topper_list, key=lambda x: x["marks"], reverse=True)[:3]
        toppers = [
            {"Rank": i+1, "Student Name": e["student"], "Marks Obtain": e["marks"], "Grade": e["grade"]}
            for i, e in enumerate(top_entries)
        ]
        topper_blocks.append({
            "S.No.": sno,
            "Subject Code": k[0],
            "Subject Name": k[1],
            "Total Present": total_present,
            "Total Absent": total_absent,
            "Pass": passed,
            "Pass %": perc_passed,
            "Toppers": toppers
        })
    gradesheet_df = pd.DataFrame(result_rows)
    total_students = len(all_students)
    overall_pass_perc = round(len(total_passed_students) / total_students * 100, 2) if total_students else 0
    return gradesheet_df, topper_blocks, total_students, overall_pass_perc

def wrap_text(text, width):
    import textwrap
    return '\n'.join(textwrap.wrap(str(text), width=width, break_long_words=False))

def render_table_image(df, title="", subject_col_wrap=36, subject_col_idx=None, wider=False):
    df_mod = df.copy()
    if subject_col_idx is None:
        subject_col_idx = (
            df.columns.get_loc("Subject Name")
            if "Subject Name" in df.columns
            else (df.columns.get_loc("Student Name") if "Student Name" in df.columns else 1)
        )
    colnames = list(df_mod.columns)
    if len(df_mod) > 0:
        for idx in df_mod.index:
            # handle empty and non-string values gracefully
            cellval = df_mod.iloc[idx, subject_col_idx]
            if pd.isnull(cellval): cellval = ''
            df_mod.iloc[idx, subject_col_idx] = wrap_text(cellval, subject_col_wrap)
    fig, ax = plt.subplots(figsize=(15 if wider else 10, 1.2 + 0.50*len(df_mod)))
    ax.axis('off')
    tbl = ax.table(cellText=df_mod.values, colLabels=colnames, loc='center', cellLoc='center', bbox=[0,0,1,1])
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(13)
    for c in range(df_mod.shape[1]):
        tbl[(0,c)].get_text().set_weight('bold')
        tbl[(0,c)].set_facecolor('#dbeffd')
        tbl.auto_set_column_width(col=c)
    for key, cell in tbl.get_celld().items():
        if key[0] == 0:
            cell.set_height(0.075)
        else:
            cell.set_height(0.055)
    if title:
        plt.title(title, fontsize=15, backgroundcolor='#f5fafd', y=1.09)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
        plt.savefig(tmpfile, bbox_inches='tight', dpi=185)
        plt.close(fig)
        return tmpfile.name

def generate_pdf(gradesheet_df, toppers):
    pdf = FPDF(unit="pt", format='A4')
    imgfile = render_table_image(
        gradesheet_df, title="Gradesheet Overview", wider=True, subject_col_wrap=40, subject_col_idx=2
    )
    pdf.add_page()
    pdf.image(imgfile, x=20, y=40, w=pdf.w-40)
    y_cursor = pdf.get_y() + (gradesheet_df.shape[0]+4)*21  # Approx height
    pdf.set_y(y_cursor)
    for block in toppers:
        df_top = pd.DataFrame(block["Toppers"])
        if df_top.empty: continue
        subject_title = f"{block['S.No.']}) {block['Subject Code']} {block['Subject Name']}"
        imgfile = render_table_image(
            df_top,
            title=f"Top 3 Toppers | {subject_title}",
            subject_col_wrap=32,
            subject_col_idx=df_top.columns.get_loc("Student Name") if "Student Name" in df_top.columns else 1,
            wider=False
        )
        y0 = pdf.get_y()
        # If not enough space left, new page
        if y0 + 290 + 32 * len(df_top) > pdf.h:
            pdf.add_page()
            y0 = 40
            pdf.set_y(y0)
        pdf.ln(8)
        pdf.set_font("Arial", size=13)
        pdf.multi_cell(0, 15,
            f"Total Present: {block['Total Present']} | "
            f"Total Absent: {block['Total Absent']} | "
            f"Pass: {block['Pass']} | "
            f"Pass %: {block['Pass %']}\n", align='C')
        y0 = pdf.get_y()
        pdf.image(imgfile, x=50, y=y0, w=pdf.w-100)
        pdf.set_y(y0 + 65 + 32 * len(df_top))
    output = BytesIO(pdf.output(dest='S').encode('latin1'))
    return output

st.set_page_config(page_title='Result Dashboard', layout='wide')
st.markdown("""
    <style>
    [data-testid="column"] > div {
        background: #f8fafd;
        border-radius: 18px;
        padding: 20px 18px 16px 18px;
        margin-bottom: 10px;
        box-shadow: 0 1px 4px 0 #e6eef7;
    }
    .big-font { font-size:22px !important; font-weight:bold; }
    .metric-label { color: #555; font-size:15px; }
    .metric-value { color: #0366d6; font-weight:bold; font-size:24px;}
    .sidebar .sidebar-content { background-color:#f3f8fe !important;}
    .block-container { margin-top:1.5rem;}
    </style>
""", unsafe_allow_html=True)

st.title("🏆 Result Ledger Analyzer (Excel & PDF Export)")

st.write("Upload your PDF to view the comprehensive grade sheet and subject-wise topper tables. Download beautiful Excel/PDF reports!")

uploaded_file = st.file_uploader("Upload your result PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("Processing..."):
        gradesheet_df, toppers, total_students, overall_pass_perc = analyze_pdf(uploaded_file)

    # KPI Tiles
    kcol1, kcol2, kcol3 = st.columns([1,1,1])
    with kcol1:
        st.markdown('<div class="metric-label">Total Students</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{total_students}</div>', unsafe_allow_html=True)
    with kcol2:
        st.markdown('<div class="metric-label">Number of Subjects</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{len(gradesheet_df)}</div>', unsafe_allow_html=True)
    with kcol3:
        st.markdown('<div class="metric-label">Overall Pass %</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{overall_pass_perc}</div>', unsafe_allow_html=True)

    # Excel download (grade sheet & toppers)
    def make_safe_sheetname(name):
        return re.sub(r'[\[\]\:\*\?\/\\]', '', name)[:28]
    def to_excel(gradesheet_df, toppers):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            gradesheet_df.to_excel(writer, sheet_name='GradeSheet', index=False)
            used_names = set()
            for block in toppers:
                sub_name = make_safe_sheetname(block['Subject Code'] + "_" + block['Subject Name'])
                base = sub_name[:28]
                base = base.rstrip('. ')
                i = 1
                sheetname = base
                while sheetname in used_names:
                    sheetname = (base + str(i))[:31]
                    i += 1
                used_names.add(sheetname)
                df_top = pd.DataFrame(block["Toppers"])
                if not df_top.empty:
                    df_top.to_excel(writer, sheet_name=sheetname, index=False)
        output.seek(0)
        return output
    excel_file = to_excel(gradesheet_df, toppers)

    st.download_button(
        label="⬇️ Download Excel (GradeSheet + All Toppers)",
        data=excel_file,
        file_name="result_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    pdf_file = generate_pdf(gradesheet_df, toppers)
    st.download_button(
        label="⬇️ Download Report as PDF",
        data=pdf_file,
        file_name="result_analysis.pdf",
        mime="application/pdf"
    )

    # -- UI Dashboard Two Columns --
    left, right = st.columns([1.3,1.7])
    with left:
        st.markdown('<div class="big-font">Gradesheet Overview</div>', unsafe_allow_html=True)
        st.dataframe(gradesheet_df.fillna(0), use_container_width=True, height=600)
    with right:
        st.markdown('<div class="big-font">Subject-wise Toppers</div>', unsafe_allow_html=True)
        for block in toppers:
            st.markdown(
                f"<div style='margin-top:1rem; margin-bottom:0.2rem;'><b>{block['S.No.']}) {block['Subject Code']} {block['Subject Name']}</b></div>"
                f"<div style='color:#444; margin-bottom:5px;'>"
                f"Total Present: <b>{block['Total Present']}</b> | "
                f"Total Absent: <b>{block['Total Absent']}</b> | "
                f"Pass: <b>{block['Pass']}</b> | "
                f"Pass %: <b>{block['Pass %']}</b>"
                f"</div>", unsafe_allow_html=True)
            if block["Toppers"]:
                st.dataframe(pd.DataFrame(block["Toppers"]), use_container_width=True, hide_index=True)
            else:
                st.info("No present students with marks found.")
