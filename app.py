import streamlit as st
import pandas as pd
import re
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

st.set_page_config(page_title="DHS Child Care Certificate Parser", page_icon="🦉", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
h1, h2, h3 { font-family: 'DM Serif Display', serif; }
.main-header { background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%);
    padding: 2rem 2.5rem; border-radius: 12px; margin-bottom: 2rem; color: white; }
.main-header h1 { color: white; margin: 0; font-size: 2rem; }
.main-header p  { color: #b8d4f0; margin: 0.4rem 0 0; font-size: 1rem; }
.info-card { background: #f0f6ff; border-left: 4px solid #2e6da4;
    border-radius: 0 8px 8px 0; padding: 1rem 1.25rem; margin-bottom: 1rem; }
.info-card h4 { margin: 0 0 0.25rem; color: #1a3a5c; font-size: 0.85rem;
    text-transform: uppercase; letter-spacing: 0.05em; }
.info-card p  { margin: 0; color: #1a3a5c; font-size: 1rem; font-weight: 500; }
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>🦉 DHS Child Care Certificate Parser</h1>
  <p>Upload an Illinois DHS Child Care Certificate Report PDF to extract data and download as Excel.</p>
</div>
""", unsafe_allow_html=True)


# ── Helpers ──────────────────────────────────────────────────────────────────

def ocr_words(page_image):
    """Run tesseract on one page image, return list of word dicts."""
    data = pytesseract.image_to_data(page_image, output_type=pytesseract.Output.DICT)
    result = []
    for i in range(len(data['text'])):
        txt = data['text'][i].strip()
        if txt and int(data['conf'][i]) > 0:
            result.append({
                'text': txt, 'left': data['left'][i], 'top': data['top'][i],
                'width': data['width'][i], 'height': data['height'][i],
                'conf': int(data['conf'][i]),
            })
    return result

def ocr_page(args):
    """Worker: OCR one page. Returns (page_idx, words, full_text)."""
    page_idx, page_image = args
    words = ocr_words(page_image)
    full_text = pytesseract.image_to_string(page_image)
    return page_idx, words, full_text

def clean_number(s):
    s = s.strip().replace(',', '')
    m = re.search(r'(\d+\.\d{1,2})', s)
    if m:
        return float(m.group(1))
    digits = re.sub(r'[^\d]', '', s)
    if len(digits) >= 3:
        return float(digits[:-2] + '.' + digits[-2:])
    return None

# Column x-ranges as fractions of page width
COL_FRACS = {
    'rate':  (0.43, 0.52),
    'sch':   (0.52, 0.58),
    'copay': (0.535, 0.665),
    'elig':  (0.665, 0.735),
    'attd':  (0.735, 0.815),
    'adj':   (0.815, 0.895),
}


# ── Main parser ───────────────────────────────────────────────────────────────

def parse_certificate_pdf(pdf_bytes: bytes, progress_bar=None) -> dict:
    provider_info = {}
    children = []
    summary = {}

    # Convert PDF to images at 150 DPI (fast, still readable for printed forms)
    pages = convert_from_bytes(pdf_bytes, dpi=150)
    n = len(pages)

    if progress_bar:
        progress_bar.progress(0.05, text=f"Converting PDF ({n} pages)…")

    # OCR all pages in parallel (up to 4 threads)
    all_results = [None] * n
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(ocr_page, (i, p)): i for i, p in enumerate(pages)}
        done = 0
        for future in as_completed(futures):
            page_idx, words, full_text = future.result()
            all_results[page_idx] = (words, full_text)
            done += 1
            if progress_bar:
                pct = 0.05 + 0.75 * (done / n)
                progress_bar.progress(pct, text=f"OCR: page {done}/{n}…")

    all_words = [r[0] for r in all_results]
    all_texts = [r[1] for r in all_results]

    if progress_bar:
        progress_bar.progress(0.82, text="Extracting provider info…")

    # ── Provider info ─────────────────────────────────────────────────────────
    for txt in all_texts[:4]:
        for line in txt.splitlines():
            line = line.strip()
            if not provider_info.get('period'):
                m = re.search(r'CERTIFICATE REPORT FOR[:\s]+(\w+)\s+(\d{4})', line, re.I)
                if m:
                    provider_info['period'] = f"{m.group(1).upper()} {m.group(2)}"
            if not provider_info.get('name'):
                m = re.search(r'PROVIDER NAME[:\s]+(.+)', line, re.I)
                if m:
                    provider_info['name'] = m.group(1).strip()
            if not provider_info.get('dhs_id'):
                m = re.search(r'DHS PROVIDER ID#[:\s]+([\d\s]+)', line, re.I)
                if m:
                    provider_info['dhs_id'] = m.group(1).strip()
            if not provider_info.get('date_of_issue'):
                m = re.search(r'Date of Issue[:\s]+(\d{1,2}/\d{1,2}/\d{4})', line, re.I)
                if m:
                    provider_info['date_of_issue'] = m.group(1)

    # Page 1 fallback: standalone provider name
    if not provider_info.get('name'):
        for line in all_texts[0].splitlines():
            line = line.strip()
            if re.match(r'^[A-Z][A-Z\s]+(LLC|INC|CORP|DAYCARE|CENTER)$', line):
                provider_info['name'] = line
                break

    # Summary from last page
    for line in all_texts[-1].splitlines():
        line = line.strip()
        m = re.search(r'PRINT PROVIDER.S NAME[:\s]+(.+)', line, re.I)
        if m: summary['signed_by'] = m.group(1).strip()
        m = re.search(r'DATE[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})', line, re.I)
        if m: summary['signature_date'] = m.group(1)
        m = re.search(r'Total Number of Attended Days[:\s]+(\d+)', line, re.I)
        if m: summary['total_attended'] = int(m.group(1))
        m = re.search(r'Total Number of Adjusted Eligible Days[:\s]+(\d+)', line, re.I)
        if m: summary['total_eligible'] = int(m.group(1))
        m = re.search(r'Percentage.*?(\d+)\s*%', line, re.I)
        if m: summary['attendance_pct'] = int(m.group(1))

    if progress_bar:
        progress_bar.progress(0.88, text="Extracting child records…")

    # ── Child records via positional OCR ──────────────────────────────────────
    for page_idx, words in enumerate(all_words[:-1]):
        if not words:
            continue
        page_w = pages[page_idx].width

        def xr(col):
            lo, hi = COL_FRACS[col]
            return lo * page_w, hi * page_w

        # Find case-number anchors: 5-digit numbers near left edge
        case_words = [w for w in words
                      if re.match(r'^\d{5}$', w['text']) and w['left'] < 0.18 * page_w]

        used = set()
        case_rows = []
        for i, w in enumerate(case_words):
            if i in used:
                continue
            group = [w]
            for j, w2 in enumerate(case_words):
                if j != i and j not in used and abs(w2['top'] - w['top']) < 40:
                    group.append(w2)
                    used.add(j)
            used.add(i)
            if len(group) >= 2:
                y_c = sum(x['top'] + x['height']/2 for x in group) / len(group)
                case_rows.append(y_c)

        for ci, y_case in enumerate(case_rows):
            y_end = case_rows[ci + 1] if ci + 1 < len(case_rows) else y_case + 700

            row_words = [w for w in words
                         if (w['top'] + w['height']/2) > y_case - 30
                         and (w['top'] + w['height']/2) < y_end]

            # Child name
            name_ws = [w for w in row_words
                       if w['left'] > 0.20 * page_w and w['left'] < 0.45 * page_w
                       and abs((w['top'] + w['height']/2) - y_case) < 60
                       and w['conf'] > 30 and re.match(r'^[A-Z]{2,}', w['text'])]
            child_name = ' '.join(w['text'] for w in sorted(name_ws, key=lambda x: x['left']))

            # DOB
            dob_ws = [w for w in row_words
                      if w['left'] > 0.20 * page_w and w['left'] < 0.45 * page_w
                      and (w['top'] + w['height']/2) > y_case + 30
                      and (w['top'] + w['height']/2) < y_case + 250
                      and re.match(r'\d{1,2}/\d{2}/\d{4}', w['text'])]
            dob = dob_ws[0]['text'] if dob_ws else ''

            # Client name
            client_ws = [w for w in row_words
                         if w['left'] < 0.20 * page_w
                         and (w['top'] + w['height']/2) > y_case + 60
                         and (w['top'] + w['height']/2) < y_case + 300
                         and w['conf'] > 35
                         and re.match(r'^[A-Z]{2,}', w['text'])
                         and w['text'] not in ('IVR', 'CASE', 'CHILD', 'CARE', 'CLIENT')]
            client_name = ' '.join(w['text'] for w in sorted(client_ws, key=lambda x: x['left']))

            # Find F-row y
            xlo_sch, xhi_sch = xr('sch')
            f_row_y = None
            for w in row_words:
                cx = w['left'] + w['width']/2
                if xlo_sch <= cx <= xhi_sch and w['text'].upper() in ('F', 'E', '='):
                    f_row_y = w['top'] + w['height']/2
                    break
            if f_row_y is None:
                xlo_r, xhi_r = xr('rate')
                for w in sorted(row_words, key=lambda x: x['top']):
                    cx = w['left'] + w['width']/2
                    if xlo_r <= cx <= xhi_r:
                        v = clean_number(w['text'])
                        if v and 10 < v < 300:
                            f_row_y = w['top'] + w['height']/2
                            break

            rate = copay = elig = attd = adj = None
            if f_row_y is not None:
                tol = 70

                def pick_number(col, min_v=0, max_v=500):
                    xlo, xhi = xr(col)
                    cands = [w for w in row_words
                             if xlo <= (w['left']+w['width']/2) <= xhi
                             and abs((w['top']+w['height']/2) - f_row_y) < tol]
                    for w in cands:
                        v = clean_number(w['text'])
                        if v is not None and min_v <= v <= max_v:
                            return v
                    return None

                def pick_int(col, min_v=0, max_v=31):
                    xlo, xhi = xr(col)
                    cands = [w for w in row_words
                             if xlo <= (w['left']+w['width']/2) <= xhi
                             and abs((w['top']+w['height']/2) - f_row_y) < tol]
                    for w in cands:
                        v = re.sub(r'[^\d]', '', w['text'])
                        if v and min_v <= int(v) <= max_v:
                            return int(v)
                    return None

                rate  = pick_number('rate',  10, 300)
                copay = pick_number('copay', 0,  500)
                elig  = pick_int('elig')
                attd  = pick_int('attd')
                adj   = pick_int('adj')

            if child_name:
                children.append({
                    'Parent/Client': client_name,
                    "Child's Name":  child_name,
                    'Date of Birth': dob,
                    'Rate':          rate,
                    'Co-Pay':        copay if copay is not None else 0.0,
                    'Elig Days':     elig,
                    'Attd Days':     attd,
                    'Adj Days':      adj,
                })

    if progress_bar:
        progress_bar.progress(1.0, text="Done!")

    return {'provider_info': provider_info, 'children': children, 'summary': summary}


def compute_payment(row):
    rate  = row.get('Rate')  or 0.0
    copay = row.get('Co-Pay') or 0.0
    adj   = row.get('Adj Days')
    elig  = row.get('Elig Days') or 0
    days  = adj if adj is not None else elig
    return round(rate * days - copay, 2)


# ── Excel builder ─────────────────────────────────────────────────────────────

def build_excel(provider_info, children, summary) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Children Detail"

    hf  = PatternFill("solid", start_color="1F4E79")
    shf = PatternFill("solid", start_color="2E75B6")
    af  = PatternFill("solid", start_color="D6E4F0")
    wf  = PatternFill("solid", start_color="FFFFFF")
    thin = Side(style="thin", color="AAAAAA")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    period = provider_info.get("period", "")

    ws.merge_cells("A1:I1")
    ws["A1"] = f"Child Care Certificate Report – {period}"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = hf
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for i, (lbl, val) in enumerate([
        ("Provider:", provider_info.get("name", "")),
        ("DHS Provider ID#:", provider_info.get("dhs_id", "")),
        ("Report Period:", period),
        ("Date of Issue:", provider_info.get("date_of_issue", "")),
    ], start=2):
        ws[f"A{i}"] = lbl; ws[f"A{i}"].font = Font(name="Arial", bold=True, size=10)
        ws[f"B{i}"] = val; ws[f"B{i}"].font = Font(name="Arial", size=10)

    hr = 7
    for c, h in enumerate(["Parent/Client", "Child's Name", "Date of Birth",
                            "Rate ($)", "Co-Pay ($)", "Elig Days", "Attd Days",
                            "Adj Days", "Payment ($)"], 1):
        cell = ws.cell(hr, c, h)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = shf; cell.border = bdr
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[hr].height = 30

    for ri, child in enumerate(children, start=hr + 1):
        fill = af if ri % 2 == 0 else wf
        for ci, val in enumerate([
            child.get("Parent/Client", ""), child.get("Child's Name", ""),
            child.get("Date of Birth", ""), child.get("Rate"),
            child.get("Co-Pay"), child.get("Elig Days"),
            child.get("Attd Days"), child.get("Adj Days"),
        ], 1):
            cell = ws.cell(ri, ci, val)
            cell.font = Font(name="Arial", size=10); cell.fill = fill
            cell.border = bdr; cell.alignment = Alignment(horizontal="center")
            if ci in (4, 5): cell.number_format = '$#,##0.00'

        pay = ws.cell(ri, 9)
        pay.value = f'=IF(H{ri}<>"",D{ri}*H{ri},D{ri}*F{ri})-E{ri}'
        pay.font = Font(name="Arial", size=10); pay.fill = fill
        pay.border = bdr; pay.alignment = Alignment(horizontal="center")
        pay.number_format = '$#,##0.00'

    tr = hr + len(children) + 1
    for c in range(1, 10):
        cell = ws.cell(tr, c)
        cell.border = bdr
        cell.fill = PatternFill("solid", start_color="BDD7EE")
        cell.font = Font(name="Arial", bold=True, size=10)
        cell.alignment = Alignment(horizontal="center")
    ws.cell(tr, 1, "TOTALS").fill = shf
    ws.cell(tr, 1).font = Font(bold=True, color="FFFFFF", name="Arial")
    ws.cell(tr, 6, f"=SUM(F{hr+1}:F{hr+len(children)})")
    ws.cell(tr, 7, f"=SUM(G{hr+1}:G{hr+len(children)})")
    ws.cell(tr, 9, f"=SUM(I{hr+1}:I{hr+len(children)})").number_format = '$#,##0.00'

    for i, w in enumerate([24, 20, 14, 11, 11, 10, 10, 10, 13], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet("Attendance Summary")
    ws2.merge_cells("A1:D1")
    ws2["A1"] = f"Attendance Summary – {period}"
    ws2["A1"].font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    ws2["A1"].fill = hf
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    pct = f"{summary.get('attendance_pct', '')}%" if summary.get('attendance_pct') else ""
    for r, (lbl, val) in enumerate([
        ("Total Attended Days",          summary.get("total_attended", "")),
        ("Total Adjusted Eligible Days", summary.get("total_eligible", "")),
        ("Attendance Percentage",        pct),
        ("", ""),
        ("Signed By",      summary.get("signed_by", "")),
        ("Signature Date", summary.get("signature_date", "")),
    ], 2):
        ws2[f"A{r}"] = lbl; ws2[f"A{r}"].font = Font(name="Arial", bold=True, size=11)
        ws2[f"B{r}"] = val; ws2[f"B{r}"].font = Font(name="Arial", size=11)
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── UI ────────────────────────────────────────────────────────────────────────

if not OCR_AVAILABLE:
    st.error("pdf2image and pytesseract must be installed. Check requirements.txt.")
    st.stop()

uploaded = st.file_uploader("Upload DHS Child Care Certificate Report (PDF)", type=["pdf"])

if uploaded:
    pdf_bytes = uploaded.read()

    progress_bar = st.progress(0, text="Starting…")
    try:
        result = parse_certificate_pdf(pdf_bytes, progress_bar=progress_bar)
    except Exception as e:
        progress_bar.empty()
        st.error(f"Parsing error: {e}")
        st.stop()
    progress_bar.empty()

    pi, children, summary = result["provider_info"], result["children"], result["summary"]

    # Provider info cards
    st.subheader("Provider Information")
    cols = st.columns(4)
    for col, (label, val) in zip(cols, [
        ("Provider",       pi.get("name", "—")),
        ("Report Period",  pi.get("period", "—")),
        ("DHS Provider ID", pi.get("dhs_id", "—")),
        ("Date of Issue",  pi.get("date_of_issue", "—")),
    ]):
        col.markdown(f'<div class="info-card"><h4>{label}</h4><p>{val}</p></div>',
                     unsafe_allow_html=True)

    st.divider()

    if children:
        st.subheader(f"Extracted Records ({len(children)} children)")
        st.caption("Review and fix any OCR errors directly in the table before downloading.")
        df = pd.DataFrame(children)
        edited_df = st.data_editor(df, use_container_width=True, hide_index=True, column_config={
            "Rate":      st.column_config.NumberColumn("Rate ($)",    format="$%.2f"),
            "Co-Pay":    st.column_config.NumberColumn("Co-Pay ($)",  format="$%.2f"),
            "Elig Days": st.column_config.NumberColumn("Elig Days",   step=1),
            "Attd Days": st.column_config.NumberColumn("Attd Days",   step=1),
            "Adj Days":  st.column_config.NumberColumn("Adj Days",    step=1),
        })
        edited_df["Payment"] = edited_df.apply(compute_payment, axis=1)
        children_out = edited_df.to_dict('records')

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Elig Days", int(edited_df["Elig Days"].fillna(0).sum()))
        c2.metric("Total Attd Days", int(edited_df["Attd Days"].fillna(0).sum()))
        c3.metric("Total Payment",   f"${edited_df['Payment'].sum():,.2f}")
    else:
        st.warning("No records extracted — scan quality may be too low. "
                   "You can add rows manually after downloading.")
        children_out = []

    st.divider()
    st.subheader("Download Excel")

    raw_name   = pi.get("name", "ChildCare")
    raw_period = pi.get("period", "")
    name_clean = re.sub(r'\b(DAYCARE|DAY CARE|LLC|INC|CORP|CENTER)\b', '',
                        raw_name, flags=re.IGNORECASE).strip()
    name_clean = " ".join(name_clean.split()).title()
    parts = raw_period.split()
    period_short = f"{parts[0][:3].title()} {parts[1]}" if len(parts) == 2 else raw_period.title()
    filename = f"{name_clean} {period_short} Certs.xlsx"

    st.download_button(
        label="⬇️  Download Excel File",
        data=build_excel(pi, children_out, summary),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.info("👆 Upload a PDF to get started.")
