import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="DHS Child Care Certificate Parser",
    page_icon="🦉",
    layout="wide",
)

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}
h1, h2, h3 {
    font-family: 'DM Serif Display', serif;
}

.main-header {
    background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%);
    padding: 2rem 2.5rem;
    border-radius: 12px;
    margin-bottom: 2rem;
    color: white;
}
.main-header h1 { color: white; margin: 0; font-size: 2rem; }
.main-header p  { color: #b8d4f0; margin: 0.4rem 0 0; font-size: 1rem; }

.info-card {
    background: #f0f6ff;
    border-left: 4px solid #2e6da4;
    border-radius: 0 8px 8px 0;
    padding: 1rem 1.25rem;
    margin-bottom: 1rem;
}
.info-card h4 { margin: 0 0 0.25rem; color: #1a3a5c; font-size: 0.85rem; text-transform: uppercase; letter-spacing: 0.05em; }
.info-card p  { margin: 0; color: #1a3a5c; font-size: 1rem; font-weight: 500; }

.stDataFrame { border-radius: 8px; overflow: hidden; }

footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>🦉 DHS Child Care Certificate Parser</h1>
  <p>Upload an Illinois DHS Child Care Certificate Report PDF to extract data and download as Excel.</p>
</div>
""", unsafe_allow_html=True)


# ── PDF Parsing ────────────────────────────────────────────────────────────────

def parse_certificate_pdf(pdf_file) -> dict:
    """
    Extract provider info and child records from a DHS Child Care Certificate PDF.
    Returns a dict with keys: provider_info, children, summary.
    """
    provider_info = {}
    children = []
    summary = {}

    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = [l.strip() for l in full_text.splitlines() if l.strip()]

    # ── Provider info ──────────────────────────────────────────────────────────
    for line in lines:
        if "CHILD CARE CERTIFICATE REPORT FOR:" in line:
            m = re.search(r'FOR:\s+(\w+)\s+(\d{4})', line)
            if m:
                provider_info["period"] = f"{m.group(1)} {m.group(2)}"
        if "PROVIDER NAME:" in line:
            provider_info["name"] = line.split("PROVIDER NAME:")[-1].strip()
        if "DHS PROVIDER ID#:" in line:
            m = re.search(r'DHS PROVIDER ID#:\s*([\d\s]+)', line)
            if m:
                provider_info["dhs_id"] = m.group(1).strip()
        if "LOCATION:" in line and "TYPE OF CARE:" in line:
            m = re.search(r'LOCATION:\s*(\w+)\s+TYPE OF CARE:\s*(\d+)', line)
            if m:
                provider_info["location"] = m.group(1)
                provider_info["type_of_care"] = m.group(2)
        if "Date of Issue:" in line or "DATE OF ISSUE:" in line.upper():
            m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', line)
            if m:
                provider_info["date_of_issue"] = m.group(1)
        if "HOW MANY DAYS WAS YOUR CENTER" in line or "OPEN FOR BUSINESS" in line:
            m = re.search(r'(\d+)\s*$', line)
            if m:
                provider_info["days_open"] = int(m.group(1))

    # ── Children records ───────────────────────────────────────────────────────
    # Pattern: case number line followed by child name + DOB + rate rows
    # We look for lines matching an IVR CASE# pattern and build records from context
    case_pattern = re.compile(
        r'^(\d{5}\s+\d{5}\s+\d{5})\s+'           # case no
        r'([A-Z][A-Z\s]+?)\s+'                     # child name
        r'(\d{1,2}/\d{2}/\d{4})'                   # DOB
    )
    client_pattern = re.compile(r'^([A-Z][A-Z\s\-]+[A-Z])$')
    rate_pattern   = re.compile(
        r'^(\d+\.\d{2})\s+(F|P|S)\s+'             # rate + schedule
        r'(\d+\.\d{2})?\s*'                        # co-pay (optional)
        r'(\d+)?\s*'                               # elig days (optional)
        r'(\d+)?\s*'                               # attd days (optional)
        r'(\d+)?'                                  # adj days (optional)
    )

    # Walk lines to find child blocks
    i = 0
    current_client = None
    current_case = None

    while i < len(lines):
        line = lines[i]

        # Detect client name lines (all-caps name on its own line, after IVR line)
        if re.match(r'^IVR CASE#:', line):
            i += 1
            continue

        # Case number + child name + DOB on same line
        m = case_pattern.match(line)
        if m:
            case_no   = m.group(1).strip()
            child_name = m.group(2).strip()
            dob        = m.group(3).strip()
            current_case = case_no

            # Next non-IVR line that is all-caps is the client name
            client_name = None
            for j in range(i + 1, min(i + 4, len(lines))):
                if re.match(r'^IVR CASE#:', lines[j]):
                    continue
                if re.match(r'^[A-Z][A-Z\s\-]+$', lines[j]) and len(lines[j]) > 3:
                    client_name = lines[j].strip()
                    break

            current_client = client_name

            # Find the F (full-day) rate row — that's the one with elig days
            rate = copay = elig = attd = adj = None
            for j in range(i + 1, min(i + 8, len(lines))):
                rm = rate_pattern.match(lines[j])
                if rm and rm.group(2) == "F":
                    rate  = float(rm.group(1))
                    copay = float(rm.group(3)) if rm.group(3) else 0.0
                    elig  = int(rm.group(4))   if rm.group(4) else None
                    attd  = int(rm.group(5))   if rm.group(5) else None
                    adj   = int(rm.group(6))   if rm.group(6) else None
                    break

            if rate is not None:
                children.append({
                    "Parent/Client":  current_client or "",
                    "Child's Name":   child_name,
                    "Date of Birth":  dob,
                    "Rate":           rate,
                    "Co-Pay":         copay,
                    "Elig Days":      elig,
                    "Attd Days":      attd,
                    "Adj Days":       adj,
                })

        # Summary totals (page 4)
        if "Total Number of Attended Days:" in line:
            m = re.search(r'(\d+)', line)
            if m:
                summary["total_attended"] = int(m.group(1))
        if "Total Number of Adjusted Eligible Days:" in line:
            m = re.search(r'(\d+)', line)
            if m:
                summary["total_eligible"] = int(m.group(1))
        if "Percentage of Attended Days" in line and "%" in line:
            m = re.search(r'(\d+)\s*%', line)
            if m:
                summary["attendance_pct"] = int(m.group(1))
        if "PROVIDER'S SIGNATURE:" in line or "PRINT PROVIDER'S NAME:" in line:
            m = re.search(r'NAME:\s*(.+)', line)
            if m:
                summary["signed_by"] = m.group(1).strip()
        if "DATE:" in line and "SIGNATURE" not in line and "10/" in line:
            m = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})', line)
            if m:
                summary["signature_date"] = m.group(1)

        i += 1

    return {"provider_info": provider_info, "children": children, "summary": summary}


def compute_payment(row):
    rate  = row["Rate"]
    copay = row["Co-Pay"] or 0.0
    adj   = row["Adj Days"]
    elig  = row["Elig Days"] or 0
    days  = adj if (adj is not None and adj != "") else elig
    return round(rate * days - copay, 2)


# ── Excel Builder ──────────────────────────────────────────────────────────────

def build_excel(provider_info: dict, children: list, summary: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Children Detail"

    header_fill    = PatternFill("solid", start_color="1F4E79")
    subheader_fill = PatternFill("solid", start_color="2E75B6")
    alt_fill       = PatternFill("solid", start_color="D6E4F0")
    white_fill     = PatternFill("solid", start_color="FFFFFF")
    thin = Side(style="thin", color="AAAAAA")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    period = provider_info.get("period", "")

    # Title
    ws.merge_cells("A1:I1")
    ws["A1"] = f"Child Care Certificate Report – {period}"
    ws["A1"].font      = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill      = header_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Provider meta
    meta = [
        ("Provider:",       provider_info.get("name", "")),
        ("DHS Provider ID#:", provider_info.get("dhs_id", "")),
        ("Report Period:",  period),
        ("Days Open:",      provider_info.get("days_open", "")),
        ("Date of Issue:",  provider_info.get("date_of_issue", "")),
    ]
    for i, (label, val) in enumerate(meta, start=2):
        ws[f"A{i}"] = label
        ws[f"A{i}"].font = Font(name="Arial", bold=True, size=10)
        ws[f"B{i}"] = val
        ws[f"B{i}"].font = Font(name="Arial", size=10)

    # Column headers
    headers = ["Parent/Client", "Child's Name", "Date of Birth",
               "Rate ($)", "Co-Pay ($)", "Elig Days", "Attd Days", "Adj Days", "Payment ($)"]
    header_row = 8
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill      = subheader_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border    = bdr
    ws.row_dimensions[header_row].height = 30

    # Data rows
    for r_idx, child in enumerate(children, start=header_row + 1):
        fill = alt_fill if r_idx % 2 == 0 else white_fill
        row_vals = [
            child.get("Parent/Client", ""),
            child.get("Child's Name", ""),
            child.get("Date of Birth", ""),
            child.get("Rate"),
            child.get("Co-Pay"),
            child.get("Elig Days"),
            child.get("Attd Days"),
            child.get("Adj Days"),
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = fill
            cell.border    = bdr
            cell.alignment = Alignment(horizontal="center")
            if c_idx in (4, 5):
                cell.number_format = '$#,##0.00'

        # Payment formula
        pay = ws.cell(row=r_idx, column=9)
        pay.value          = f'=IF(H{r_idx}<>"",D{r_idx}*H{r_idx},D{r_idx}*F{r_idx})-E{r_idx}'
        pay.font           = Font(name="Arial", size=10)
        pay.fill           = fill
        pay.border         = bdr
        pay.alignment      = Alignment(horizontal="center")
        pay.number_format  = '$#,##0.00'

    # Totals
    total_row = header_row + len(children) + 1
    for c in range(1, 10):
        cell = ws.cell(total_row, c)
        cell.border    = bdr
        cell.fill      = PatternFill("solid", start_color="BDD7EE")
        cell.font      = Font(name="Arial", bold=True, size=10)
        cell.alignment = Alignment(horizontal="center")
    ws.cell(total_row, 1, "TOTALS").fill = subheader_fill
    ws.cell(total_row, 1).font = Font(bold=True, color="FFFFFF", name="Arial")
    ws.cell(total_row, 6, f"=SUM(F{header_row+1}:F{header_row+len(children)})")
    ws.cell(total_row, 7, f"=SUM(G{header_row+1}:G{header_row+len(children)})")
    pt = ws.cell(total_row, 9, f"=SUM(I{header_row+1}:I{header_row+len(children)})")
    pt.number_format = '$#,##0.00'

    # Column widths
    for i, w in enumerate([24, 20, 14, 11, 11, 10, 10, 10, 13], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Sheet 2: Attendance Summary
    ws2 = wb.create_sheet("Attendance Summary")
    ws2.merge_cells("A1:D1")
    ws2["A1"] = f"Attendance Summary – {period}"
    ws2["A1"].font      = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    ws2["A1"].fill      = header_fill
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    att = summary.get("total_attended", "")
    elig = summary.get("total_eligible", "")
    pct  = f"{summary.get('attendance_pct', '')}%" if summary.get('attendance_pct') else ""
    summ_rows = [
        ("Total Attended Days",           att),
        ("Total Adjusted Eligible Days",   elig),
        ("Attendance Percentage",          pct),
        ("", ""),
        ("Signed By",       summary.get("signed_by", "")),
        ("Signature Date",  summary.get("signature_date", "")),
    ]
    for r, (label, val) in enumerate(summ_rows, start=2):
        ws2[f"A{r}"] = label
        ws2[f"A{r}"].font = Font(name="Arial", bold=True, size=11)
        ws2[f"B{r}"] = val
        ws2[f"B{r}"].font = Font(name="Arial", size=11)
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Streamlit UI ───────────────────────────────────────────────────────────────

uploaded = st.file_uploader(
    "Upload DHS Child Care Certificate Report (PDF)",
    type=["pdf"],
    help="Illinois DHS IL444-3492A form"
)

if uploaded:
    with st.spinner("Parsing PDF…"):
        try:
            result = parse_certificate_pdf(uploaded)
        except Exception as e:
            st.error(f"Could not parse PDF: {e}")
            st.stop()

    pi       = result["provider_info"]
    children = result["children"]
    summary  = result["summary"]

    # ── Provider info cards ────────────────────────────────────────────────────
    st.subheader("Provider Information")
    cols = st.columns(4)
    fields = [
        ("Provider",      pi.get("name", "—")),
        ("Report Period", pi.get("period", "—")),
        ("DHS Provider ID", pi.get("dhs_id", "—")),
        ("Days Open",     pi.get("days_open", "—")),
    ]
    for col, (label, val) in zip(cols, fields):
        col.markdown(f"""
        <div class="info-card">
          <h4>{label}</h4>
          <p>{val}</p>
        </div>""", unsafe_allow_html=True)

    st.divider()

    # ── Children table ─────────────────────────────────────────────────────────
    if children:
        st.subheader(f"Children ({len(children)} records)")
        df = pd.DataFrame(children)
        df["Payment"] = df.apply(compute_payment, axis=1)

        # Format for display
        df_display = df.copy()
        df_display["Rate"]    = df_display["Rate"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "")
        df_display["Co-Pay"]  = df_display["Co-Pay"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "")
        df_display["Payment"] = df_display["Payment"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "")
        df_display.columns = ["Parent/Client", "Child's Name", "Date of Birth",
                               "Rate", "Co-Pay", "Elig Days", "Attd Days", "Adj Days", "Payment"]
        st.dataframe(df_display, use_container_width=True, hide_index=True)

        total_payment = df["Payment"].sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Elig Days",  int(df["Elig Days"].fillna(0).sum()))
        c2.metric("Total Attd Days",  int(df["Attd Days"].fillna(0).sum()))
        c3.metric("Total Payment",    f"${total_payment:,.2f}")
    else:
        st.warning("No child records could be extracted. The PDF layout may differ from expected — you can edit the data manually below.")
        children = []

    st.divider()

    # ── Download ───────────────────────────────────────────────────────────────
    st.subheader("Download Excel")

    # Build filename: e.g. "Wise Owl Oct 2025 Certs.xlsx"
    raw_name   = pi.get("name", "ChildCare")          # e.g. "WISE OWL DAYCARE LLC"
    raw_period = pi.get("period", "")                  # e.g. "OCTOBER 2025"

    # Shorten name: drop generic suffixes, title-case
    name_clean = re.sub(r'\b(DAYCARE|DAY CARE|LLC|INC|CORP|CENTER)\b', '', raw_name, flags=re.IGNORECASE).strip()
    name_clean = " ".join(name_clean.split()).title()   # collapse spaces, title-case

    # Shorten month: "OCTOBER 2025" → "Oct 2025"
    parts = raw_period.split()
    if len(parts) == 2:
        period_short = f"{parts[0][:3].title()} {parts[1]}"
    else:
        period_short = raw_period.title()

    filename = f"{name_clean} {period_short} Certs.xlsx"
    excel_bytes = build_excel(pi, children, summary)

    st.download_button(
        label="⬇️  Download Excel File",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.info("👆 Upload a PDF to get started.")
