"""
MCA Bank Statement Parser — FastAPI Backend
Accepts PDF upload, runs pdfplumber + analysis, returns Excel underwriting schedule.
Deploy on Render.com (free tier) or Railway.app.
"""

import os
import io
import re
import uuid
import tempfile
from datetime import datetime
from typing import Optional

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse

app = FastAPI(
    title="XGen Automations — MCA Bank Statement Parser",
    description="Instantly parse bank statement PDFs into Excel underwriting schedules. Built by XGen Automations.",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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
    """Extract all transactions from PDF bytes using pdfplumber."""
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

    # Fallback: regex parse raw text if table extraction got nothing
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
    """Calculate ADB, Deposits, Withdrawals, NSF count, Negative Days."""
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
    """Build a professionally styled Excel underwriting schedule."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Underwriting Schedule"

    # Styles
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

    # ── Title Block ──
    ws.merge_cells("A1:F1")
    ws["A1"] = "XGen Automations — MCA Bank Statement Underwriting Schedule"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    ws.merge_cells("A2:F2")
    ws["A2"] = f"Prepared: {datetime.now().strftime('%B %d, %Y %I:%M %p')}   |   Source: {filename}"
    ws["A2"].font = Font(name="Calibri", italic=True, size=9, color="666666")
    ws["A2"].alignment = center

    ws.append([])  # blank row 3

    # ── Summary Block ──
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

    ws.append([])  # blank row 6

    # ── Transaction Table ──
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

    # Column widths
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


# ─────────────────────────────────────────────────────────
#  API ROUTES
# ─────────────────────────────────────────────────────────

@app.get("/")
def root():
    return {
        "service": "XGen Automations — MCA Bank Statement Parser",
        "status": "online",
        "usage": "POST /parse-pdf with a PDF file attached (field: file)",
        "version": "1.0.0"
    }

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/parse-pdf")
async def parse_pdf(file: UploadFile = File(...)):
    """
    Upload a bank statement PDF.
    Returns a professionally styled Excel underwriting schedule as a download.
    """
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    if file.size and file.size > 20 * 1024 * 1024:  # 20MB limit
        raise HTTPException(status_code=400, detail="File too large. Max 20MB.")

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
async def parse_pdf_preview(file: UploadFile = File(...)):
    """
    Same as /parse-pdf but returns JSON preview of metrics (for live dashboard display).
    """
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
            "transactions_preview": transactions[:10],
            "message": "Parsing complete. Download /parse-pdf for full Excel."
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Preview error: {str(e)}")
