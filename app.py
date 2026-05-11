import os
import io
import json
import base64
import re
import traceback
import anthropic
from flask import Flask, request, jsonify, send_file
from pdf2image import convert_from_bytes
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

ALLOWED_EXTENSIONS = {'.pdf', '.jpg', '.jpeg', '.png'}

api_key = os.environ.get("ANTHROPIC_API_KEY")
if not api_key:
    raise RuntimeError("ANTHROPIC_API_KEY environment variable is not set")
client = anthropic.Anthropic(api_key=api_key)

# ── Inline HTML (no templates folder needed) ───────────────────────────────────

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DHS Child Care Certificate Parser</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'DM Sans', sans-serif; background: #f4f6fa; min-height: 100vh; color: #1a2a3a; }
  .header { background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%); padding: 2rem 2.5rem; color: white; }
  .header h1 { font-family: 'DM Serif Display', serif; font-size: 1.9rem; margin-bottom: 0.3rem; }
  .header p { color: #b8d4f0; font-size: 0.95rem; }
  .container { max-width: 900px; margin: 2rem auto; padding: 0 1.5rem; }
  .card { background: white; border-radius: 12px; padding: 2rem; box-shadow: 0 2px 12px rgba(0,0,0,0.07); margin-bottom: 1.5rem; }
  .upload-zone { border: 2px dashed #adc8e8; border-radius: 10px; padding: 3rem 2rem; text-align: center; cursor: pointer; transition: all 0.2s; background: #f8fbff; }
  .upload-zone:hover, .upload-zone.drag-over { border-color: #2e6da4; background: #eef5ff; }
  .upload-zone .icon { font-size: 3rem; margin-bottom: 0.75rem; }
  .upload-zone p { color: #4a6a8a; font-size: 0.95rem; }
  .upload-zone strong { color: #1a3a5c; }
  #file-input { display: none; }
  .btn { display: inline-flex; align-items: center; gap: 0.5rem; padding: 0.75rem 1.5rem; border: none; border-radius: 8px; font-family: inherit; font-size: 0.95rem; font-weight: 600; cursor: pointer; transition: all 0.2s; }
  .btn-primary { background: #2e6da4; color: white; }
  .btn-primary:hover { background: #1a3a5c; }
  .btn-primary:disabled { background: #9ab8d4; cursor: not-allowed; }
  .btn-success { background: #1a7a4a; color: white; }
  .btn-success:hover { background: #145c38; }
  .btn-lg { padding: 1rem 2rem; font-size: 1.05rem; width: 100%; justify-content: center; }
  .progress-wrap { margin-top: 1.5rem; display: none; }
  .progress-bar-outer { height: 8px; background: #dde8f5; border-radius: 99px; overflow: hidden; margin-bottom: 0.5rem; }
  .progress-bar-inner { height: 100%; background: #2e6da4; border-radius: 99px; width: 0%; transition: width 0.4s ease; }
  .progress-text { font-size: 0.85rem; color: #4a6a8a; }
  .info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1rem; margin-bottom: 1.5rem; }
  .info-card { background: #f0f6ff; border-left: 4px solid #2e6da4; border-radius: 0 8px 8px 0; padding: 0.85rem 1rem; }
  .info-card h4 { font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.05em; color: #4a6a8a; margin-bottom: 0.2rem; }
  .info-card p { font-size: 0.95rem; font-weight: 600; color: #1a3a5c; }
  .metrics { display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; margin-bottom: 1.5rem; }
  .metric { background: #1a3a5c; color: white; border-radius: 10px; padding: 1rem; text-align: center; }
  .metric-value { font-size: 1.8rem; font-weight: 700; }
  .metric-label { font-size: 0.78rem; opacity: 0.75; margin-top: 0.2rem; }
  .table-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
  th { background: #2e6da4; color: white; padding: 0.6rem 0.75rem; text-align: left; font-weight: 600; white-space: nowrap; }
  td { padding: 0.55rem 0.75rem; border-bottom: 1px solid #e8eef5; }
  tr:nth-child(even) td { background: #f0f6ff; }
  tr:hover td { background: #ddeeff; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  .editable:focus { background: #fffde7 !important; outline: none; }
  .error-box { background: #fff0f0; border: 1px solid #ffcccc; border-radius: 8px; padding: 1rem 1.25rem; color: #c00; font-size: 0.9rem; margin-top: 1rem; display: none; }
  #results-section { display: none; }
  .section-title { font-family: 'DM Serif Display', serif; font-size: 1.25rem; margin-bottom: 1rem; color: #1a3a5c; }
  .hint { font-size: 0.8rem; color: #7a9ab8; margin-top: 0.4rem; }
</style>
</head>
<body>
<div class="header">
  <h1>🦉 DHS Child Care Certificate Parser</h1>
  <p>Upload a DHS Child Care Certificate Report — Claude reads it and generates a formatted Excel file.</p>
</div>
<div class="container">
  <div class="card">
    <div class="upload-zone" id="upload-zone">
      <div class="icon">📄</div>
      <p><strong>Drop your file here</strong> or click to browse</p>
      <p style="margin-top:0.4rem;font-size:0.82rem;">PDF, JPG, or PNG &middot; IL444-3492A form &middot; up to 50 MB</p>
    </div>
    <input type="file" id="file-input" accept=".pdf,.jpg,.jpeg,.png">
    <div class="progress-wrap" id="progress-wrap">
      <div class="progress-bar-outer"><div class="progress-bar-inner" id="progress-bar"></div></div>
      <div class="progress-text" id="progress-text">Starting&hellip;</div>
    </div>
    <div class="error-box" id="error-box"></div>
    <div style="margin-top:1.25rem;">
      <button class="btn btn-primary btn-lg" id="parse-btn" disabled>&#10024; Parse Certificate</button>
      <p class="hint" style="text-align:center;margin-top:0.5rem;">Claude reads each page using vision AI &mdash; typically 15&ndash;45 seconds depending on page count.</p>
    </div>
  </div>

  <div id="results-section">
    <div class="card">
      <div class="section-title">Provider Information</div>
      <div class="info-grid" id="info-grid"></div>
      <div class="metrics" id="metrics"></div>
    </div>
    <div class="card">
      <div class="section-title" id="children-title">Children</div>
      <p class="hint" style="margin-bottom:0.75rem;">Click any cell to edit before downloading.</p>
      <div class="table-wrap">
        <table id="data-table">
          <thead><tr>
            <th>Parent/Client</th><th>Child's Name</th><th>Date of Birth</th>
            <th>Rate ($)</th><th>Co-Pay ($)</th><th>Elig Days</th>
            <th>Attd Days</th><th>Adj Days</th><th>Payment ($)</th>
          </tr></thead>
          <tbody id="table-body"></tbody>
        </table>
      </div>
      <div style="margin-top:1.5rem;">
        <button class="btn btn-success btn-lg" id="download-btn">&#11015;&#65039; Download Excel</button>
      </div>
    </div>
  </div>
</div>

<script>
let parsedData = null;
const uploadZone  = document.getElementById('upload-zone');
const fileInput   = document.getElementById('file-input');
const parseBtn    = document.getElementById('parse-btn');
const progressWrap = document.getElementById('progress-wrap');
const progressBar  = document.getElementById('progress-bar');
const progressText = document.getElementById('progress-text');
const errorBox    = document.getElementById('error-box');
const resultsSection = document.getElementById('results-section');
const downloadBtn = document.getElementById('download-btn');

uploadZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) {
    uploadZone.querySelector('p').textContent = '\uD83D\uDCCE ' + fileInput.files[0].name;
    parseBtn.disabled = false;
    errorBox.style.display = 'none';
  }
});
uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', e => {
  e.preventDefault();
  uploadZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  const ok = ['.pdf','.jpg','.jpeg','.png'].some(ext => file.name.toLowerCase().endsWith(ext));
  if (file && ok) {
    fileInput.files = e.dataTransfer.files;
    uploadZone.querySelector('p').textContent = '\uD83D\uDCCE ' + file.name;
    parseBtn.disabled = false;
  }
});

parseBtn.addEventListener('click', async () => {
  const file = fileInput.files[0];
  if (!file) return;
  parseBtn.disabled = true;
  errorBox.style.display = 'none';
  resultsSection.style.display = 'none';
  progressWrap.style.display = 'block';

  let pct = 5;
  progressBar.style.width = pct + '%';
  progressText.textContent = 'Uploading\u2026';
  const ticker = setInterval(() => {
    if (pct < 88) {
      pct += Math.random() * 2.5;
      progressBar.style.width = Math.min(pct, 88) + '%';
      if (pct < 20) progressText.textContent = 'Uploading\u2026';
      else if (pct < 50) progressText.textContent = 'Claude is reading the pages\u2026';
      else if (pct < 75) progressText.textContent = 'Extracting child records\u2026';
      else progressText.textContent = 'Almost done\u2026';
    }
  }, 900);

  try {
    const form = new FormData();
    form.append('file', file);
    const resp = await fetch('/parse', { method: 'POST', body: form });
    const data = await resp.json();
    clearInterval(ticker);
    if (!resp.ok || data.error) throw new Error(data.error || 'Unknown error');
    progressBar.style.width = '100%';
    progressText.textContent = 'Done \u2014 ' + data.count + ' children extracted.';
    parsedData = data;
    renderResults(data);
    resultsSection.style.display = 'block';
    // Show page errors if any and count is 0
    if (data.count === 0 && data.page_errors && data.page_errors.length > 0) {
      errorBox.innerHTML = '\u274C Parser errors:<br>' + data.page_errors.slice(0,5).join('<br>');
      errorBox.style.display = 'block';
    } else if (data.count === 0) {
      errorBox.textContent = '\u26A0\uFE0F No records found. Check Railway logs for details, or try the /debug endpoint.';
      errorBox.style.display = 'block';
    }
  } catch (err) {
    clearInterval(ticker);
    progressWrap.style.display = 'none';
    errorBox.textContent = '\u274C ' + err.message;
    errorBox.style.display = 'block';
    parseBtn.disabled = false;
  }
});

function calcPayment(rate, copay, elig, attd, adj) {
  rate  = parseFloat(rate)  || 0;
  copay = parseFloat(copay) || 0;
  const days = (adj !== null && adj !== '' && adj !== undefined) ? parseFloat(adj) : (parseFloat(elig) || 0);
  return (rate * days - copay).toFixed(2);
}

function renderResults(data) {
  const pi = data.provider || {};
  const children = data.children || [];
  const summary  = data.summary  || {};

  document.getElementById('info-grid').innerHTML = [
    ['Provider', pi.name || '\u2014'],
    ['Report Period', pi.period || '\u2014'],
    ['DHS Provider ID', pi.dhs_id || '\u2014'],
    ['Date of Issue', pi.date_of_issue || '\u2014'],
  ].map(([l,v]) => `<div class="info-card"><h4>${l}</h4><p>${v}</p></div>`).join('');

  const totalAttd = children.reduce((s,c) => s + (parseInt(c.attd_days)||0), 0);
  const totalPay  = children.reduce((s,c) => s + parseFloat(calcPayment(c.rate, c.copay, c.elig_days, c.attd_days, c.adj_days)), 0);
  document.getElementById('metrics').innerHTML = `
    <div class="metric"><div class="metric-value">${children.length}</div><div class="metric-label">Children</div></div>
    <div class="metric"><div class="metric-value">${totalAttd}</div><div class="metric-label">Attended Days</div></div>
    <div class="metric"><div class="metric-value">$${totalPay.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}</div><div class="metric-label">Total Payment</div></div>`;

  document.getElementById('children-title').textContent = 'Children (' + children.length + ' records)';

  const tbody = document.getElementById('table-body');
  tbody.innerHTML = '';
  children.forEach((c, idx) => {
    const pay = calcPayment(c.rate, c.copay, c.elig_days, c.attd_days, c.adj_days);
    const tr = document.createElement('tr');
    tr.dataset.idx = idx;
    tr.innerHTML = `
      <td contenteditable class="editable" data-field="parent_client">${c.parent_client||''}</td>
      <td contenteditable class="editable" data-field="child_name">${c.child_name||''}</td>
      <td contenteditable class="editable" data-field="dob">${c.dob||''}</td>
      <td contenteditable class="editable num" data-field="rate">${c.rate!=null?parseFloat(c.rate).toFixed(2):''}</td>
      <td contenteditable class="editable num" data-field="copay">${c.copay!=null?parseFloat(c.copay).toFixed(2):''}</td>
      <td contenteditable class="editable num" data-field="elig_days">${c.elig_days??''}</td>
      <td contenteditable class="editable num" data-field="attd_days">${c.attd_days??''}</td>
      <td contenteditable class="editable num" data-field="adj_days">${c.adj_days??''}</td>
      <td class="num pay-cell">$${pay}</td>`;
    tr.querySelectorAll('[contenteditable]').forEach(cell => {
      cell.addEventListener('blur', () => {
        const field = cell.dataset.field;
        let val = cell.textContent.trim();
        const numFields = ['rate','copay','elig_days','attd_days','adj_days'];
        if (numFields.includes(field)) val = val === '' ? null : parseFloat(val);
        parsedData.children[idx][field] = val;
        const ch = parsedData.children[idx];
        tr.querySelector('.pay-cell').textContent = '$' + calcPayment(ch.rate, ch.copay, ch.elig_days, ch.attd_days, ch.adj_days);
      });
    });
    tbody.appendChild(tr);
  });
}

downloadBtn.addEventListener('click', async () => {
  if (!parsedData) return;
  downloadBtn.disabled = true;
  downloadBtn.textContent = '\u23F3 Building Excel\u2026';
  try {
    const resp = await fetch('/download', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsedData)
    });
    if (!resp.ok) throw new Error('Download failed');
    const blob = await resp.blob();
    const cd = resp.headers.get('Content-Disposition') || '';
    const match = cd.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : 'Certs.xlsx';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    alert('Download error: ' + err.message);
  } finally {
    downloadBtn.disabled = false;
    downloadBtn.textContent = '\u2B07\uFE0F Download Excel';
  }
});
</script>
</body>
</html>"""

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
- is_data_page: false for blank pages, signature-only pages, or pages with no child data
- provider: fill from header if visible; leave fields as empty string if not present
- children: only records with actual data (name + rate). Skip column header rows.
- rate: use the full-day (F row) rate — the row that has elig_days filled in
- copay: parent co-pay from the F row; use 0.0 if blank
- elig_days/attd_days/adj_days: integers or null. Never strings.
- For children marked "no longer attends" (C code): adj_days = 0, attd_days = 0
- summary: only fill on the final signature/summary page; leave null elsewhere
- Do not invent data. If a field is blank or illegible, use null or empty string.
"""

def image_to_b64(pil_image):
    buf = io.BytesIO()
    pil_image.save(buf, format="JPEG", quality=85)
    return base64.standard_b64encode(buf.getvalue()).decode()

def parse_page_with_claude(pil_image, page_num=1):
    b64 = image_to_b64(pil_image)
    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text", "text": "Extract all data from this DHS Child Care Certificate page. Return only JSON."}
            ]
        }]
    )
    raw = message.content[0].text.strip()
    raw = re.sub(r'^```json\s*', '', raw, flags=re.MULTILINE)
    raw = re.sub(r'^```\s*', '', raw, flags=re.MULTILINE)
    raw = re.sub(r'```\s*$', '', raw, flags=re.MULTILINE)
    raw = raw.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        app.logger.error(f"Page {page_num} JSON error: {e}\nRaw: {raw[:500]}")
        raise ValueError(f"Claude returned invalid JSON on page {page_num}: {e}")

def get_pages(file_bytes, filename):
    ext = os.path.splitext(filename.lower())[1]
    if ext == '.pdf':
        pages = convert_from_bytes(file_bytes, dpi=150)
        app.logger.info(f"PDF: {len(pages)} pages")
        return pages
    elif ext in ('.jpg', '.jpeg', '.png'):
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        app.logger.info("Image: single page")
        return [img]
    else:
        raise ValueError(f"Unsupported file type: {ext}")

def parse_upload(file_bytes, filename):
    pages = get_pages(file_bytes, filename)
    provider_info = {}
    all_children  = []
    summary       = {}
    page_errors   = []

    for i, page in enumerate(pages):
        app.logger.info(f"Parsing page {i+1}/{len(pages)}")
        try:
            result = parse_page_with_claude(page, page_num=i+1)
        except Exception as e:
            err_msg = f"Page {i+1}: {e}"
            app.logger.error(err_msg)
            page_errors.append(err_msg)
            continue

        if not result.get("is_data_page", True):
            app.logger.info(f"Page {i+1}: skipped (not a data page)")
            continue

        p = result.get("provider") or {}
        for key in ("name", "dhs_id", "period", "date_of_issue", "days_open"):
            if not provider_info.get(key) and p.get(key):
                provider_info[key] = p[key]

        page_kids = result.get("children") or []
        for child in page_kids:
            if child.get("child_name"):
                all_children.append(child)
        app.logger.info(f"Page {i+1}: {len(page_kids)} children, is_data_page={result.get('is_data_page')}")

        s = result.get("summary") or {}
        for key in ("total_attended", "total_eligible", "attendance_pct", "signed_by", "sig_date"):
            if s.get(key) not in (None, "", 0):
                summary[key] = s[key]

    app.logger.info(f"Done: {len(all_children)} children, provider={provider_info.get('name')}, errors={len(page_errors)}")
    return provider_info, all_children, summary, page_errors

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
        ("Provider:", provider_info.get("name", "")),
        ("DHS Provider ID#:", provider_info.get("dhs_id", "")),
        ("Report Period:", period),
        ("Days Open:", provider_info.get("days_open", "")),
        ("Date of Issue:", provider_info.get("date_of_issue", "")),
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
        for ci, val in enumerate([
            child.get("parent_client", ""), child.get("child_name", ""),
            child.get("dob", ""), child.get("rate"), child.get("copay"),
            child.get("elig_days"), child.get("attd_days"), child.get("adj_days"),
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
        cell.border = bdr; cell.fill = PatternFill("solid", start_color="BDD7EE")
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
        ("Total Attended Days", summary.get("total_attended", "")),
        ("Total Adjusted Eligible Days", summary.get("total_eligible", "")),
        ("Attendance Percentage", pct), ("", ""),
        ("Signed By", summary.get("signed_by", "")),
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

def make_filename(provider_info):
    raw_name = provider_info.get("name", "ChildCare")
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
    return INDEX_HTML, 200, {"Content-Type": "text/html"}

@app.route("/parse", methods=["POST"])
def parse():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400
    ext = os.path.splitext(f.filename.lower())[1]
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": f"Unsupported file type '{ext}'. Upload a PDF, JPG, or PNG."}), 400

    file_bytes = f.read()
    app.logger.info(f"File: {f.filename} ({len(file_bytes)} bytes)")
    try:
        provider_info, children, summary, page_errors = parse_upload(file_bytes, f.filename)
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

    return jsonify({"provider": provider_info, "children": children,
                    "summary": summary, "count": len(children),
                    "page_errors": page_errors})

@app.route("/download", methods=["POST"])
def download():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON body"}), 400
        buf = build_excel(data.get("provider", {}), data.get("children", []), data.get("summary", {}))
        return send_file(buf, as_attachment=True,
                         download_name=make_filename(data.get("provider", {})),
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route("/health")
def health():
    return jsonify({"status": "ok", "api_key_set": bool(api_key)})

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Maximum 50MB."}), 413

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)


@app.route("/debug", methods=["POST"])
def debug():
    """Parse just the first page and return raw Claude output for debugging."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    file_bytes = f.read()
    try:
        pages = get_pages(file_bytes, f.filename)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    results = []
    # Parse first 3 pages so we can see what Claude returns
    for i, page in enumerate(pages[:3]):
        try:
            b64 = image_to_b64(page)
            message = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=2000,
                system=SYSTEM_PROMPT,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                        {"type": "text", "text": "Extract all data from this DHS Child Care Certificate page. Return only JSON."}
                    ]
                }]
            )
            raw = message.content[0].text.strip()
            # Try to parse, but return raw too
            try:
                parsed = json.loads(re.sub(r'```[\w]*\n?|```', '', raw).strip())
            except Exception:
                parsed = None
            results.append({
                "page": i + 1,
                "raw_response": raw[:2000],
                "parsed": parsed,
                "error": None
            })
        except Exception as e:
            results.append({"page": i + 1, "raw_response": None, "parsed": None, "error": str(e)})

    return jsonify({"total_pages": len(pages), "pages_sampled": len(results), "results": results})
