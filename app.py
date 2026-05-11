import os
import io
import json
import base64
import re
import anthropic
from flask import Flask, request, jsonify, send_file, render_template
from pdf2image import convert_from_bytes
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

ALLOWED_EXTENSIONS = {'.pdf', '.jpg', '.jpeg', '.png'}

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

# ── Claude vision parser ───────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are a data extraction assistant for Illinois DHS Child Care Certificate Reports.
When given an image of a certificate page, extract all child records and return ONLY valid JSON.
No preamble, no explanation, no markdown fences — raw JSON only.

Return this exact structure:
{
  "provider": {
    "name": "",
    "dhs_id": "",
    "period": "",
    "date_of_issue": "",
    "days_open": null
  },
  "children": [
    {
      "parent_client": "",
      "child_name": "",
      "dob": "",
      "rate": null,
      "copay": null,
      "elig_days": null,
      "attd_days": null,
      "adj_days": null
    }
  ],
  "summary": {
    "total_attended": null,
    "total_eligible": null,
    "attendance_pct": null,
    "signed_by": "",
    "sig_date": ""
  },
  "is_data_page": true
}

Rules:
- is_data_page: false for blank pages, signature-only pages, or supplemental cert pages with no children
- provider: fill from header if visible; leave blank if not on this page
- children: only records with actual data (name + rate). Skip header rows.
- rate: use the full-day (F row) rate — the row that has elig_days filled in
- copay: parent co-pay from the F row
- elig_days/attd_days/adj_days: integers or null
- For children marked "no longer attends" (C code): adj_days = 0, attd_days = 0
- For supplemental pages: include children normally
- summary: only fill on the final signature page; leave null elsewhere
- Do not invent data. If a field is blank/illegible, use null or ""
"""

def image_to_b64(pil_image):
    buf = io.BytesIO()
    pil_image.save(buf, format="JPEG", quality=85)
    return base64.standard_b64encode(buf.getvalue()).decode()

def parse_page_with_claude(pil_image):
    """Send one page image to Claude and get structured JSON back."""
    b64 = image_to_b64(pil_image)
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
                },
                {
                    "type": "text",
                    "text": "Extract all data from this DHS Child Care Certificate page. Return only JSON."
                }
            ]
        }]
    )
    raw = message.content[0].text.strip()
    raw = re.sub(r'^```json\s*', '', raw)
    raw = re.sub(r'^```\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)
    return json.loads(raw)

def get_pages(file_bytes, filename):
    """
    Return a list of PIL Images from the uploaded file.
    Supports PDF (multi-page) and JPG/PNG (single image).
    """
    ext = os.path.splitext(filename.lower())[1]
    if ext == '.pdf':
        return convert_from_bytes(file_bytes, dpi=150)
    elif ext in ('.jpg', '.jpeg', '.png'):
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        return [img]
    else:
        raise ValueError(f"Unsupported file type: {ext}")

def parse_upload(file_bytes, filename):
    """Convert file to pages, parse each with Claude, merge results."""
    pages = get_pages(file_bytes, filename)

    provider_info = {}
    all_children = []
    summary = {}

    for i, page in enumerate(pages):
        try:
            result = parse_page_with_claude(page)
        except Exception as e:
            print(f"Page {i+1} parse error: {e}")
            continue

        if not result.get("is_data_page", True):
            continue

        # Merge provider info (first non-empty value wins)
        p = result.get("provider", {})
        for key in ("name", "dhs_id", "period", "date_of_issue", "days_open"):
            if not provider_info.get(key) and p.get(key):
                provider_info[key] = p[key]

        # Collect children
        for child in result.get("children", []):
            if child.get("child_name"):
                all_children.append(child)

        # Merge summary (last page with data wins)
        s = result.get("summary", {})
        for key in ("total_attended", "total_eligible", "attendance_pct", "signed_by", "sig_date"):
            if s.get(key):
                summary[key] = s[key]

    return provider_info, all_children, summary

# ── Excel builder ──────────────────────────────────────────────────────────────

def build_excel(provider_info, children, summary):
    wb = Workbook()
    ws = wb.active
    ws.title = "Children Detail"

    hf  = PatternFill("solid", start_color="1F4E79")
    shf = PatternFill("solid", start_color="2E75B6")
    af  = PatternFill("solid", start_color="D6E4F0")
    wf  = PatternFill("solid", start_color="FFFFFF")
    thin = Side(style="thin", color="AAAAAA")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    period = provider_info.get("period", "")

    ws.merge_cells("A1:I1")
    ws["A1"] = f"Child Care Certificate Report – {period}"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = hf
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for i, (lbl, val) in enumerate([
        ("Provider:",        provider_info.get("name", "")),
        ("DHS Provider ID#:", provider_info.get("dhs_id", "")),
        ("Report Period:",   period),
        ("Days Open:",       provider_info.get("days_open", "")),
        ("Date of Issue:",   provider_info.get("date_of_issue", "")),
    ], start=2):
        ws[f"A{i}"] = lbl; ws[f"A{i}"].font = Font(name="Arial", bold=True, size=10)
        ws[f"B{i}"] = val; ws[f"B{i}"].font = Font(name="Arial", size=10)

    hr = 8
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
        vals = [
            child.get("parent_client", ""),
            child.get("child_name", ""),
            child.get("dob", ""),
            child.get("rate"),
            child.get("copay"),
            child.get("elig_days"),
            child.get("attd_days"),
            child.get("adj_days"),
        ]
        for ci, val in enumerate(vals, 1):
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
        ("Signature Date", summary.get("sig_date", "")),
    ], 2):
        ws2[f"A{r}"] = lbl; ws2[f"A{r}"].font = Font(name="Arial", bold=True, size=11)
        ws2[f"B{r}"] = val; ws2[f"B{r}"].font = Font(name="Arial", size=11)
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Filename helper ────────────────────────────────────────────────────────────

def make_filename(provider_info):
    raw_name   = provider_info.get("name", "ChildCare")
    raw_period = provider_info.get("period", "")
    name_clean = re.sub(r'\b(DAYCARE|DAY CARE|LLC|INC|CORP|CENTER)\b', '',
                        raw_name, flags=re.IGNORECASE).strip()
    name_clean = " ".join(name_clean.split()).title()
    parts = raw_period.split()
    period_short = f"{parts[0][:3].title()} {parts[1]}" if len(parts) == 2 else raw_period.title()
    return f"{name_clean} {period_short} Certs.xlsx"

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/parse", methods=["POST"])
def parse():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    ext = os.path.splitext(f.filename.lower())[1]
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": f"Unsupported file type. Please upload a PDF, JPG, or PNG."}), 400

    file_bytes = f.read()
    try:
        provider_info, children, summary = parse_upload(file_bytes, f.filename)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify({
        "provider": provider_info,
        "children": children,
        "summary": summary,
        "count": len(children)
    })

@app.route("/download", methods=["POST"])
def download():
    data = request.get_json()
    provider_info = data.get("provider", {})
    children      = data.get("children", [])
    summary       = data.get("summary", {})

    buf = build_excel(provider_info, children, summary)
    filename = make_filename(provider_info)

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/health")
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
