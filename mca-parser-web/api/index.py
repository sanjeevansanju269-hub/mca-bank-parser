import os
import io
import re
from datetime import datetime

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, HTMLResponse

app = FastAPI(
    title="XGen Automations — MCA Bank Statement Parser",
    description="All-in-one MCA Bank Statement Underwriting Engine",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

HTML_CONTENT = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>MCA Bank Statement Parser | XGen Automations</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: #0a0f1e;
      color: #e2e8f0;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .card {
      background: #111827;
      border: 1px solid #1e40af;
      border-radius: 16px;
      padding: 44px 40px;
      max-width: 580px;
      width: 100%;
      box-shadow: 0 0 60px rgba(30, 64, 175, 0.2);
    }
    .badge {
      display: inline-block;
      background: rgba(30, 64, 175, 0.25);
      color: #60a5fa;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 1.5px;
      text-transform: uppercase;
      padding: 4px 12px;
      border-radius: 100px;
      border: 1px solid rgba(96, 165, 250, 0.3);
      margin-bottom: 16px;
    }
    h1 {
      font-size: 26px;
      font-weight: 700;
      color: #f1f5f9;
      line-height: 1.3;
      margin-bottom: 8px;
    }
    p.sub {
      color: #94a3b8;
      font-size: 14px;
      margin-bottom: 28px;
      line-height: 1.6;
    }
    .upload-zone {
      border: 2px dashed #1e40af;
      border-radius: 12px;
      padding: 32px 20px;
      text-align: center;
      cursor: pointer;
      transition: all 0.2s;
      background: rgba(30, 64, 175, 0.04);
      position: relative;
    }
    .upload-zone:hover, .upload-zone.drag {
      border-color: #3b82f6;
      background: rgba(59, 130, 246, 0.08);
    }
    .upload-zone input[type="file"] {
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
      width: 100%;
      height: 100%;
    }
    .upload-icon { font-size: 36px; margin-bottom: 10px; }
    .upload-zone p { color: #64748b; font-size: 14px; }
    .upload-zone strong { color: #93c5fd; }
    .filename-display {
      margin-top: 12px;
      color: #34d399;
      font-size: 13px;
      font-weight: 500;
      display: none;
    }
    .btn {
      display: block;
      width: 100%;
      margin-top: 20px;
      padding: 14px;
      background: #1d4ed8;
      color: white;
      font-size: 15px;
      font-weight: 600;
      border: none;
      border-radius: 10px;
      cursor: pointer;
      transition: background 0.2s;
      letter-spacing: 0.3px;
    }
    .btn:hover:not(:disabled) { background: #2563eb; }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .status {
      margin-top: 18px;
      padding: 14px 16px;
      border-radius: 10px;
      font-size: 13px;
      font-weight: 500;
      display: none;
    }
    .status.loading { background: rgba(251, 191, 36, 0.1); color: #fbbf24; border: 1px solid rgba(251, 191, 36, 0.3); }
    .status.success { background: rgba(52, 211, 153, 0.1); color: #34d399; border: 1px solid rgba(52, 211, 153, 0.3); }
    .status.error { background: rgba(239, 68, 68, 0.1); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }
    .metrics-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 12px;
      margin-top: 20px;
      display: none;
    }
    .metric-card {
      background: #0f172a;
      border: 1px solid #1e293b;
      border-radius: 10px;
      padding: 14px;
    }
    .metric-card .label { color: #64748b; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px; }
    .metric-card .value { color: #f1f5f9; font-size: 17px; font-weight: 700; }
    .metric-card .value.red { color: #f87171; }
    .metric-card .value.green { color: #34d399; }
    .divider { border: none; border-top: 1px solid #1e293b; margin: 28px 0; }
    .features { display: flex; gap: 12px; flex-wrap: wrap; margin-top: 20px; }
    .feature { display: flex; align-items: center; gap: 6px; font-size: 12px; color: #64748b; }
    .feature span { color: #3b82f6; }
    footer { margin-top: 28px; color: #475569; font-size: 12px; text-align: center; }
    footer a { color: #3b82f6; text-decoration: none; }
  </style>
</head>
<body>

<div class="card">
  <div class="badge">XGen Automations</div>
  <h1>MCA Bank Statement Parser</h1>
  <p class="sub">Upload any bank statement PDF. We extract every transaction, calculate Average Daily Balance (ADB), NSF count, and generate an Excel underwriting schedule in seconds.</p>

  <div class="upload-zone" id="uploadZone">
    <input type="file" id="fileInput" accept=".pdf" />
    <div class="upload-icon">📄</div>
    <p><strong>Click to upload</strong> or drag and drop</p>
    <p style="margin-top:4px; font-size:12px;">PDF bank statements only · Max 20MB</p>
    <div class="filename-display" id="filenameDisplay"></div>
  </div>

  <button class="btn" id="parseBtn" disabled onclick="parsePDF()">
    ⚡ Parse & Download Excel
  </button>

  <div class="status" id="statusBox"></div>

  <div class="metrics-grid" id="metricsGrid">
    <div class="metric-card">
      <div class="label">Total Deposits</div>
      <div class="value green" id="m_deposits">—</div>
    </div>
    <div class="metric-card">
      <div class="label">Total Withdrawals</div>
      <div class="value red" id="m_withdrawals">—</div>
    </div>
    <div class="metric-card">
      <div class="label">Avg Daily Balance</div>
      <div class="value green" id="m_adb">—</div>
    </div>
    <div class="metric-card">
      <div class="label">NSF Count</div>
      <div class="value" id="m_nsf">—</div>
    </div>
    <div class="metric-card">
      <div class="label">Net Cash Flow</div>
      <div class="value" id="m_net">—</div>
    </div>
    <div class="metric-card">
      <div class="label">Transactions Found</div>
      <div class="value" id="m_txn">—</div>
    </div>
  </div>

  <hr class="divider"/>

  <div class="features">
    <div class="feature"><span>✓</span> Scanned PDFs supported</div>
    <div class="feature"><span>✓</span> Auto ADB calculation</div>
    <div class="feature"><span>✓</span> NSF & Overdraft detection</div>
    <div class="feature"><span>✓</span> Styled Excel output</div>
  </div>
</div>

<footer>
  Built by <a href="https://xgenautomations.com" target="_blank">XGen Automations</a> · For MCA brokers & commercial lenders
</footer>

<script>
  const fileInput = document.getElementById("fileInput");
  const parseBtn = document.getElementById("parseBtn");
  const filenameDisplay = document.getElementById("filenameDisplay");
  const statusBox = document.getElementById("statusBox");
  const metricsGrid = document.getElementById("metricsGrid");
  const uploadZone = document.getElementById("uploadZone");

  let selectedFile = null;

  fileInput.addEventListener("change", () => {
    selectedFile = fileInput.files[0];
    if (selectedFile) {
      filenameDisplay.style.display = "block";
      filenameDisplay.textContent = "✓ " + selectedFile.name;
      parseBtn.disabled = false;
    }
  });

  uploadZone.addEventListener("dragover", (e) => { e.preventDefault(); uploadZone.classList.add("drag"); });
  uploadZone.addEventListener("dragleave", () => uploadZone.classList.remove("drag"));
  uploadZone.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadZone.classList.remove("drag");
    if (e.dataTransfer.files[0]) {
      fileInput.files = e.dataTransfer.files;
      fileInput.dispatchEvent(new Event("change"));
    }
  });

  function showStatus(type, msg) {
    statusBox.className = "status " + type;
    statusBox.style.display = "block";
    statusBox.textContent = msg;
  }

  function fmt(val) { return "$" + parseFloat(val).toLocaleString("en-US", {minimumFractionDigits: 2}); }

  async function parsePDF() {
    if (!selectedFile) return;

    parseBtn.disabled = true;
    parseBtn.textContent = "⏳ Processing...";
    showStatus("loading", "Extracting transactions and calculating underwriting metrics...");
    metricsGrid.style.display = "none";

    try {
      const previewForm = new FormData();
      previewForm.append("file", selectedFile);
      const previewRes = await fetch("/api/parse-pdf-preview", { method: "POST", body: previewForm });

      if (previewRes.ok) {
        const preview = await previewRes.json();
        const m = preview.metrics;
        document.getElementById("m_deposits").textContent = fmt(m.total_deposits);
        document.getElementById("m_withdrawals").textContent = fmt(m.total_withdrawals);
        document.getElementById("m_adb").textContent = fmt(m.average_daily_balance);
        document.getElementById("m_nsf").textContent = m.nsf_count;
        document.getElementById("m_net").textContent = fmt(m.net_cash_flow);
        document.getElementById("m_txn").textContent = m.transaction_count + " txns";
        document.getElementById("m_net").className = "value " + (m.net_cash_flow >= 0 ? "green" : "red");
        metricsGrid.style.display = "grid";
      }

      const downloadForm = new FormData();
      downloadForm.append("file", selectedFile);
      const downloadRes = await fetch("/api/parse-pdf", { method: "POST", body: downloadForm });

      if (!downloadRes.ok) {
        const err = await downloadRes.json();
        throw new Error(err.detail || "Processing failed.");
      }

      const blob = await downloadRes.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = selectedFile.name.replace(".pdf", "_Underwriting_Schedule.xlsx");
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      showStatus("success", "✓ Underwriting Excel generated and downloaded!");
    } catch (err) {
      showStatus("error", "Error: " + err.message);
    }

    parseBtn.disabled = false;
    parseBtn.textContent = "⚡ Parse & Download Excel";
  }
</script>

</body>
</html>"""

def clean_amount(val):
    if not val:
        return 0.0
    val_str = str(val).replace('$', '').replace(',', '').strip()
    if val_str.startswith('(') and val_str.endswith(')'):
        val_str = '-' + val_str[1:-1]
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def extract_transactions(pdf_bytes: bytes):
    transactions = []
    text_lines = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_pages = len(pdf.pages)
        for page in pdf.pages:
            text = page.extract_text() or ""
            text_lines.append(text)
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                for row in table:
                    if not row or len(row) < 3:
                        continue
                    row_str = [str(c).strip() if c else "" for c in row]
                    date_match = re.match(r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', row_str[0])
                    if date_match:
                        date_str = date_match.group(1)
                        desc = row_str[1] if len(row_str) > 1 else "—"
                        amount_raw = row_str[2] if len(row_str) > 2 else "0"
                        balance_raw = row_str[-1] if len(row_str) > 3 else "0"
                        amount = clean_amount(amount_raw)
                        balance = clean_amount(balance_raw)
                        txn_type = "Deposit" if amount > 0 else "Withdrawal"
                        transactions.append({
                            "Date": date_str,
                            "Description": desc[:60],
                            "Amount ($)": round(amount, 2),
                            "Balance ($)": round(balance, 2),
                            "Type": txn_type
                        })

    if not transactions:
        full_text = "\n".join(text_lines)
        pattern = re.compile(
            r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s+'
            r'(.{5,50}?)\s+'
            r'([+-]?\$?[\d,]+\.\d{2})'
        )
        for match in pattern.finditer(full_text):
            date_str, desc, amt_str = match.groups()
            amount = clean_amount(amt_str)
            transactions.append({
                "Date": date_str,
                "Description": desc.strip()[:60],
                "Amount ($)": round(amount, 2),
                "Balance ($)": 0.0,
                "Type": "Deposit" if amount > 0 else "Withdrawal"
            })

    return transactions, total_pages

def calculate_metrics(transactions):
    if not transactions:
        return {
            "total_deposits": 0.0,
            "total_withdrawals": 0.0,
            "net_cash_flow": 0.0,
            "average_daily_balance": 0.0,
            "nsf_count": 0,
            "negative_days": 0,
            "transaction_count": 0
        }

    df = pd.DataFrame(transactions)
    deposits = df[df["Amount ($)"] > 0]["Amount ($)"].sum()
    withdrawals = abs(df[df["Amount ($)"] < 0]["Amount ($)"].sum())
    net = deposits - withdrawals

    balances = df["Balance ($)"].values
    valid_balances = [b for b in balances if b != 0.0]
    adb = sum(valid_balances) / len(valid_balances) if valid_balances else 0.0

    nsf_keywords = ["nsf", "overdraft", "returned", "insufficient", "od fee"]
    nsf_count = sum(
        1 for desc in df["Description"].str.lower()
        if any(kw in desc for kw in nsf_keywords)
    )

    negative_days = len([b for b in balances if b < 0])

    return {
        "total_deposits": round(deposits, 2),
        "total_withdrawals": round(withdrawals, 2),
        "net_cash_flow": round(net, 2),
        "average_daily_balance": round(adb, 2),
        "nsf_count": nsf_count,
        "negative_days": negative_days,
        "transaction_count": len(transactions)
    }

def build_excel(transactions, metrics, filename="statement") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Underwriting Schedule"

    navy = "003366"
    light_blue = "D6E4F0"
    green = "1A7A4A"
    red = "C0392B"
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    title_font = Font(name="Calibri", bold=True, size=14, color=navy)
    data_font = Font(name="Calibri", size=10)
    header_fill = PatternFill("solid", fgColor=navy)
    alt_fill = PatternFill("solid", fgColor=light_blue)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    def border():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    ws.merge_cells("A1:F1")
    ws["A1"] = "XGen Automations — MCA Bank Statement Underwriting Schedule"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    ws.merge_cells("A2:F2")
    ws["A2"] = f"Prepared: {datetime.now().strftime('%B %d, %Y %I:%M %p')}   |   Source: {filename}"
    ws["A2"].font = Font(name="Calibri", italic=True, size=9, color="666666")
    ws["A2"].alignment = center

    ws.append([])

    summary_headers = ["Total Deposits", "Total Withdrawals", "Net Cash Flow",
                        "Avg Daily Balance", "NSF Count", "Negative Days"]
    summary_values = [
        f"${metrics['total_deposits']:,.2f}",
        f"${metrics['total_withdrawals']:,.2f}",
        f"${metrics['net_cash_flow']:,.2f}",
        f"${metrics['average_daily_balance']:,.2f}",
        str(metrics['nsf_count']),
        str(metrics['negative_days'])
    ]

    for col_idx, (h, v) in enumerate(zip(summary_headers, summary_values), start=1):
        h_cell = ws.cell(row=4, column=col_idx, value=h)
        h_cell.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
        h_cell.fill = PatternFill("solid", fgColor="1A5276")
        h_cell.alignment = center
        h_cell.border = border()

        v_cell = ws.cell(row=5, column=col_idx, value=v)
        is_neg = (col_idx == 3 and metrics['net_cash_flow'] < 0)
        v_cell.font = Font(name="Calibri", bold=True, size=11,
                            color=red if is_neg else green)
        v_cell.alignment = center
        v_cell.border = border()

    ws.append([])

    col_headers = ["Date", "Description", "Amount ($)", "Balance ($)", "Type"]
    for col_idx, h in enumerate(col_headers, start=1):
        cell = ws.cell(row=7, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border()

    for row_idx, txn in enumerate(transactions, start=8):
        fill = alt_fill if row_idx % 2 == 0 else None
        values = [txn["Date"], txn["Description"],
                  txn["Amount ($)"], txn["Balance ($)"], txn["Type"]]
        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.alignment = left
            if fill:
                cell.fill = fill
            cell.border = border()
            if col_idx == 3 and isinstance(val, float) and val < 0:
                cell.font = Font(name="Calibri", size=10, color=red)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 14

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 22

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

# ── ROUTES ──

@app.get("/", response_class=HTMLResponse)
@app.get("/api", response_class=HTMLResponse)
def serve_ui():
    return HTML_CONTENT

@app.get("/health")
@app.get("/api/health")
def health():
    return {"status": "ok", "service": "XGen Automations MCA Parser"}

@app.post("/parse-pdf")
@app.post("/api/parse-pdf")
async def parse_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    try:
        pdf_bytes = await file.read()
        transactions, total_pages = extract_transactions(pdf_bytes)
        metrics = calculate_metrics(transactions)
        excel_bytes = build_excel(transactions, metrics, file.filename)

        output_filename = file.filename.replace(".pdf", "_Underwriting_Schedule.xlsx")

        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{output_filename}"'}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Parsing error: {str(e)}")

@app.post("/parse-pdf-preview")
@app.post("/api/parse-pdf-preview")
async def parse_pdf_preview(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    try:
        pdf_bytes = await file.read()
        transactions, total_pages = extract_transactions(pdf_bytes)
        metrics = calculate_metrics(transactions)

        return {
            "success": True,
            "filename": file.filename,
            "pages_processed": total_pages,
            "metrics": metrics,
            "transactions_preview": transactions[:10]
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Preview error: {str(e)}")
