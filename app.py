from flask import Flask, request, jsonify, send_file
from pypdf import PdfReader, PdfWriter
import pdfplumber
from pdf2image import convert_from_bytes
import pytesseract
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re, uuid, zipfile
from datetime import datetime
from collections import Counter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024

SESSIONS = {}

# ─── DATE / AMOUNT HELPERS ───────────────────────────────────────────────────
MONTHS = {
    'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
    'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12
}

def parse_date(date_str):
    date_str = date_str.strip()
    m = re.match(r'(\d{1,2})[\s\-/]([A-Za-z]{3})[\s\-/](\d{2,4})', date_str)
    if m:
        d, mo, y = int(m.group(1)), MONTHS.get(m.group(2).lower(), 0), int(m.group(3))
        if y < 100: y += 2000
        if mo: return datetime(y, mo, d)
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', date_str)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return datetime(y, mo, d)
    return None

def parse_amount(amt_str):
    if not amt_str: return None
    s = str(amt_str).strip().replace(',', '').replace(' ', '')
    negative = s.startswith('(') or s.startswith('-')
    s = s.strip('()-').strip()
    try:
        val = float(s)
        return -val if negative else val
    except:
        return None

DATE_RE   = re.compile(r'\b(\d{1,2}[\s\-/][A-Za-z]{3}[\s\-/]\d{2,4}|\d{1,2}/\d{1,2}/\d{4})\b')
AMOUNT_RE = re.compile(r'[\d,]+\.\d{2}')

# ─── PDF PARSING ─────────────────────────────────────────────────────────────
def extract_text_from_pdf(pdf_bytes):
    text_pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text_pages.append(page.extract_text() or '')
    full_text = '\n'.join(text_pages)
    if len(full_text.strip()) < 100 * max(len(text_pages), 1):
        images = convert_from_bytes(pdf_bytes, dpi=200)
        full_text = '\n'.join(pytesseract.image_to_string(img) for img in images)
    return full_text

def extract_tables_from_pdf(pdf_bytes):
    all_tables = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            all_tables.extend(page.extract_tables())
    return all_tables

def detect_columns(header_row):
    mapping = {}
    for i, cell in enumerate(header_row):
        if not cell: continue
        h = str(cell).lower().strip()
        if any(w in h for w in ['date','txn date','trans date','transaction date','post date']):
            mapping.setdefault('date', i)
        elif any(w in h for w in ['desc','narration','particular','detail','merchant','transaction detail']):
            mapping.setdefault('description', i)
        elif 'debit' in h and 'credit' not in h:
            mapping.setdefault('debit', i)
        elif 'credit' in h and 'debit' not in h:
            mapping.setdefault('credit', i)
        elif any(w in h for w in ['amount','amt','inr']):
            mapping.setdefault('amount', i)
        elif 'type' in h or 'dr/cr' in h or 'cr/dr' in h:
            mapping.setdefault('type', i)
    return mapping

def parse_transactions_from_tables(tables):
    transactions = []
    for table in tables:
        if not table or len(table) < 2: continue
        header_idx = None
        for i, row in enumerate(table):
            if not row: continue
            row_text = ' '.join(str(c).lower() for c in row if c)
            if any(w in row_text for w in ['date','description','amount','debit','narration']):
                header_idx = i; break
        if header_idx is None: continue
        col_map = detect_columns(table[header_idx])
        if 'date' not in col_map: continue
        for row in table[header_idx + 1:]:
            if not row or all(not c for c in row): continue
            parsed_date = parse_date(str(row[col_map['date']] or '').strip())
            if not parsed_date: continue
            desc = str(row[col_map.get('description', 1)] or '').strip() if col_map.get('description') is not None else ''
            amount, txn_type = None, ''
            if 'debit' in col_map and 'credit' in col_map:
                dv = parse_amount(row[col_map['debit']])
                cv = parse_amount(row[col_map['credit']])
                if dv: amount, txn_type = dv, 'Debit'
                elif cv: amount, txn_type = cv, 'Credit'
            elif 'amount' in col_map:
                amount = parse_amount(row[col_map['amount']])
                txn_type = str(row[col_map['type']] or '').strip() if 'type' in col_map else ('Credit' if (amount and amount < 0) else 'Debit')
                if amount: amount = abs(amount)
            if amount is None: continue
            transactions.append({'date': parsed_date, 'description': desc, 'amount': abs(amount), 'type': txn_type or 'Debit'})
    return transactions

def parse_transactions_from_text(text):
    transactions = []
    for line in text.split('\n'):
        line = line.strip()
        dates = DATE_RE.findall(line)
        if not dates: continue
        amounts = AMOUNT_RE.findall(line)
        if not amounts: continue
        parsed_date = parse_date(dates[0])
        if not parsed_date: continue
        desc = line
        for d in dates: desc = desc.replace(d, '')
        for a in amounts: desc = desc.replace(a, '').replace(',', '')
        txn_type = 'Credit' if re.search(r'\b(cr|credit|payment|reversal)\b', desc, re.I) else 'Debit'
        desc = re.sub(r'\b(debit|credit|dr|cr)\b', '', desc, flags=re.I)
        desc = re.sub(r'\s{2,}', ' ', desc).strip(' -|/')
        amount = parse_amount(amounts[-1])
        if not amount: continue
        transactions.append({'date': parsed_date, 'description': desc, 'amount': abs(amount), 'type': txn_type})
    return transactions

def extract_account_info(text):
    info = {}
    m = re.search(r'(?:account|card)[\s\w]*?:?\s*([Xx*\d]{4}[\s\-]?[Xx*\d]{4}[\s\-]?[Xx*\d]{4}[\s\-]?[\dXx*]{4})', text, re.I)
    if m: info['card_number'] = m.group(1).strip()
    m = re.search(r'(?:period|from|statement date)[:\s]+([A-Za-z\d\s]+?)\s+to\s+([A-Za-z\d\s]+?)(?:\n|$)', text, re.I)
    if m: info['period_from'] = m.group(1).strip(); info['period_to'] = m.group(2).strip()
    # Detect bank name from text
    for bank in ['HDFC','ICICI','SBI','AXIS','KOTAK','CITI','AMEX','IDFC','YES BANK','INDUSIND']:
        if bank in text.upper():
            info['bank'] = bank; break
    return info

def parse_pdf(pdf_bytes, source_name=''):
    text = extract_text_from_pdf(pdf_bytes)
    tables = extract_tables_from_pdf(pdf_bytes)
    account_info = extract_account_info(text)
    account_info['source'] = source_name
    transactions = parse_transactions_from_tables(tables)
    if len(transactions) < 2:
        transactions = parse_transactions_from_text(text)
    for t in transactions:
        t['source'] = source_name
        t['bank'] = account_info.get('bank', 'Unknown')
    return account_info, transactions

# ─── EXCEL BUILDER ───────────────────────────────────────────────────────────
def build_excel(all_transactions, account_infos):
    wb = openpyxl.Workbook()
    HDR_FILL  = PatternFill('solid', start_color='1A1A2E')
    ALT_FILL  = PatternFill('solid', start_color='EBF0FA')
    DEB_FILL  = PatternFill('solid', start_color='FFF0F0')
    CRD_FILL  = PatternFill('solid', start_color='F0FFF4')
    SUM_FILL  = PatternFill('solid', start_color='FFF8E1')
    HDR_FONT  = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    BODY_FONT = Font(name='Arial', size=10)
    BOLD_FONT = Font(name='Arial', bold=True, size=10)
    thin = Side(style='thin', color='D0D0D0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(ws, row, col, val, width=None):
        c = ws.cell(row=row, column=col, value=val)
        c.font = HDR_FONT; c.fill = HDR_FILL; c.border = border
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        if width: ws.column_dimensions[get_column_letter(col)].width = width

    # ── Sheet 1: All Transactions ────────────────────────────────────────────
    ws = wb.active
    ws.title = 'All Transactions'
    ws.freeze_panes = 'A3'
    ws.merge_cells('A1:H1')
    t = ws['A1']
    t.value = 'Credit Card — Consolidated Statement'
    t.font = Font(name='Arial', bold=True, size=14, color='1A1A2E')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 32

    for col, (h, w) in enumerate(zip(
        ['#', 'Date', 'Description', 'Bank', 'Type', 'Debit (Rs)', 'Credit (Rs)', 'Source File'],
        [5,   14,    44,             14,      12,     16,           16,             30]
    ), 1):
        hdr(ws, 2, col, h, w)
    ws.row_dimensions[2].height = 22

    sorted_txns = sorted(all_transactions, key=lambda x: x['date'])
    for i, txn in enumerate(sorted_txns, 1):
        row = i + 2
        is_credit = txn['type'].lower() in ('credit', 'cr', 'payment', 'reversal')
        fill = CRD_FILL if is_credit else (ALT_FILL if i % 2 == 0 else PatternFill())
        vals = [i, txn['date'].strftime('%d %b %Y'), txn['description'], txn.get('bank',''),
                txn['type'],
                None if is_credit else txn['amount'],
                txn['amount'] if is_credit else None,
                txn['source']]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = BODY_FONT; c.fill = fill; c.border = border
            if col == 1: c.alignment = Alignment(horizontal='center')
            elif col in (6, 7) and val is not None:
                c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal='right')

    total_row = len(sorted_txns) + 3
    for col in range(1, 9):
        c = ws.cell(row=total_row, column=col)
        c.fill = SUM_FILL; c.border = border
    ws.cell(row=total_row, column=3, value='TOTAL').font = BOLD_FONT
    ws.cell(row=total_row, column=3).fill = SUM_FILL
    ws.cell(row=total_row, column=3).border = border
    for col, formula in [(6, f'=SUM(F3:F{total_row-1})'), (7, f'=SUM(G3:G{total_row-1})')]:
        c = ws.cell(row=total_row, column=col, value=formula)
        c.font = BOLD_FONT; c.fill = SUM_FILL; c.border = border
        c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal='right')

    # ── Sheet 2: Summary ────────────────────────────────────────────────────
    ws2 = wb.create_sheet('Summary')
    ws2.column_dimensions['A'].width = 32
    ws2.column_dimensions['B'].width = 22
    ws2.merge_cells('A1:B1')
    ws2['A1'].value = 'Consolidated Summary'
    ws2['A1'].font = Font(name='Arial', bold=True, size=13, color='1A1A2E')
    ws2['A1'].alignment = Alignment(horizontal='center')
    ws2.row_dimensions[1].height = 28

    def srow(ws, row, label, val, fmt='#,##0.00', bold=False, fill=None):
        c1 = ws.cell(row=row, column=1, value=label)
        c2 = ws.cell(row=row, column=2, value=val)
        f = BOLD_FONT if bold else BODY_FONT
        c1.font = f; c2.font = f
        c1.border = border; c2.border = border
        c2.number_format = fmt; c2.alignment = Alignment(horizontal='right')
        if fill: c1.fill = fill; c2.fill = fill

    hdr(ws2, 2, 1, 'Metric'); hdr(ws2, 2, 2, 'Value')
    r = 3
    srow(ws2, r, 'Total Transactions', len(sorted_txns), '0'); r+=1
    srow(ws2, r, 'Date From', sorted_txns[0]['date'].strftime('%d %b %Y') if sorted_txns else '-', '@'); r+=1
    srow(ws2, r, 'Date To', sorted_txns[-1]['date'].strftime('%d %b %Y') if sorted_txns else '-', '@'); r+=1
    srow(ws2, r, 'Cards / Files', len(account_infos), '0'); r+=2
    srow(ws2, r, 'Total Debits (Rs)', f"='All Transactions'!F{total_row}", bold=True, fill=DEB_FILL); r+=1
    srow(ws2, r, 'Total Credits (Rs)', f"='All Transactions'!G{total_row}", bold=True, fill=CRD_FILL); r+=1
    srow(ws2, r, 'Net Spend (Rs)', f"='All Transactions'!F{total_row}-'All Transactions'!G{total_row}", bold=True, fill=SUM_FILL); r+=2

    # Per-bank breakdown
    hdr(ws2, r, 1, 'Bank'); hdr(ws2, r, 2, 'Transactions'); r+=1
    for bank, cnt in Counter(t.get('bank','Unknown') for t in sorted_txns).items():
        srow(ws2, r, bank, cnt, '0'); r+=1

    r += 1
    hdr(ws2, r, 1, 'Source File'); hdr(ws2, r, 2, 'Transactions'); r+=1
    for src, cnt in Counter(t['source'] for t in sorted_txns).items():
        srow(ws2, r, src, cnt, '0'); r+=1

    # ── Sheet 3: Card Details ────────────────────────────────────────────────
    ws3 = wb.create_sheet('Card Details')
    for col, w in zip(range(1,6), [35, 14, 22, 22, 22]):
        ws3.column_dimensions[get_column_letter(col)].width = w
    ws3.merge_cells('A1:E1')
    ws3['A1'].value = 'Source Card / File Details'
    ws3['A1'].font = Font(name='Arial', bold=True, size=13, color='1A1A2E')
    ws3['A1'].alignment = Alignment(horizontal='center')
    for col, h in enumerate(['File Name','Bank','Card Number','Period From','Period To'], 1):
        hdr(ws3, 2, col, h)
    for i, info in enumerate(account_infos, 3):
        for col, key in enumerate(['source','bank','card_number','period_from','period_to'], 1):
            c = ws3.cell(row=i, column=col, value=info.get(key, '-'))
            c.font = BODY_FONT; c.border = border
            c.fill = ALT_FILL if i % 2 == 0 else PatternFill()

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─── EMBEDDED HTML ───────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Credit Card PDF Suite</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@700;800&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{--bg:#08090f;--surface:#0f1018;--card:#12141f;--border:#1c1f30;--accent:#6366f1;--acc2:#818cf8;--green:#22d3a5;--red:#f87171;--yellow:#fbbf24;--text:#e2e8f0;--muted:#64748b}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;min-height:100vh;overflow-x:hidden}
body::before{content:'';position:fixed;inset:0;background-image:linear-gradient(var(--border) 1px,transparent 1px),linear-gradient(90deg,var(--border) 1px,transparent 1px);background-size:48px 48px;opacity:.3;pointer-events:none}
.glow{position:fixed;top:-300px;left:50%;transform:translateX(-50%);width:700px;height:700px;background:radial-gradient(circle,rgba(99,102,241,.1) 0%,transparent 65%);pointer-events:none}
.wrap{max-width:880px;margin:0 auto;padding:52px 20px 80px;position:relative}
header{text-align:center;margin-bottom:48px}
.logo-box{width:54px;height:54px;background:linear-gradient(135deg,var(--accent),#4f46e5);border-radius:14px;display:inline-flex;align-items:center;justify-content:center;font-size:26px;box-shadow:0 0 32px rgba(99,102,241,.4);margin-bottom:16px}
h1{font-family:'Syne',sans-serif;font-size:clamp(1.8rem,5vw,2.9rem);font-weight:800;letter-spacing:-1px;background:linear-gradient(135deg,#fff 35%,var(--acc2));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.tagline{color:var(--muted);font-family:'DM Mono',monospace;font-size:.8rem;margin-top:8px}
.bank-chips{display:flex;flex-wrap:wrap;justify-content:center;gap:7px;margin-top:14px}
.chip{background:var(--surface);border:1px solid var(--border);border-radius:999px;padding:4px 13px;font-family:'DM Mono',monospace;font-size:.73rem;color:var(--muted)}

.tabs{display:flex;background:var(--card);border:1px solid var(--border);border-radius:14px;padding:5px;margin-bottom:28px;gap:6px}
.tab{flex:1;text-align:center;padding:13px 10px;border-radius:10px;font-family:'Syne',sans-serif;font-weight:700;font-size:.88rem;cursor:pointer;transition:all .2s;color:var(--muted);border:none;background:none}
.tab.active{background:var(--accent);color:#fff;box-shadow:0 0 20px rgba(99,102,241,.35)}
.tab-badge{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;background:rgba(255,255,255,.15);font-size:.7rem;margin-right:6px;vertical-align:middle}
.tab.done .tab-badge{background:var(--green);color:#000}

.panel{display:none}
.panel.active{display:block;animation:fadeUp .25s ease}
@keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}

.dropzone{border:2px dashed var(--border);border-radius:18px;padding:44px 24px;text-align:center;cursor:pointer;transition:all .25s;background:var(--card)}
.dropzone:hover,.dropzone.over{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent),0 0 36px rgba(99,102,241,.12)}
.drop-icon{font-size:2.6rem;margin-bottom:12px}
.drop-title{font-family:'Syne',sans-serif;font-size:1.2rem;font-weight:700;margin-bottom:5px}
.drop-sub{color:var(--muted);font-family:'DM Mono',monospace;font-size:.78rem}
input[type=file]{display:none}

.pill-wrap{margin-top:16px;display:flex;flex-wrap:wrap;gap:7px}
.pill{display:inline-flex;align-items:center;gap:6px;background:var(--surface);border:1px solid var(--border);border-radius:999px;padding:4px 12px;font-size:.75rem;font-family:'DM Mono',monospace;animation:pop .15s ease}
.pill .x{cursor:pointer;color:var(--muted);margin-left:2px;transition:color .15s}
.pill .x:hover{color:var(--red)}
@keyframes pop{from{transform:scale(.8);opacity:0}to{transform:scale(1);opacity:1}}

.pw-row{display:flex;gap:10px;margin-top:22px;align-items:stretch}
.pw-wrap{flex:1;position:relative}
.pw-wrap input{width:100%;background:var(--card);border:1px solid var(--border);border-radius:12px;padding:13px 46px 13px 17px;color:var(--text);font-family:'DM Mono',monospace;font-size:.95rem;outline:none;transition:border-color .2s,box-shadow .2s}
.pw-wrap input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(99,102,241,.18)}
.pw-wrap input::placeholder{color:var(--muted)}
.eye{position:absolute;right:13px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--muted);font-size:1rem;padding:4px;transition:color .15s}
.eye:hover{color:var(--accent)}

.btn{border:none;border-radius:12px;font-family:'Syne',sans-serif;font-weight:700;font-size:.95rem;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:all .2s;white-space:nowrap}
.btn-primary{background:linear-gradient(135deg,var(--accent),#4f46e5);color:#fff;padding:13px 26px;box-shadow:0 4px 20px rgba(99,102,241,.3)}
.btn-primary:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 6px 28px rgba(99,102,241,.45)}
.btn-primary:disabled{opacity:.45;cursor:not-allowed}
.btn-green{background:rgba(34,211,165,.12);border:1px solid rgba(34,211,165,.3);color:var(--green);padding:14px 22px}
.btn-green:hover{background:rgba(34,211,165,.22);border-color:var(--green);transform:translateY(-1px)}
.btn-outline{background:none;border:1px solid var(--border);color:var(--muted);padding:10px 20px;font-family:'DM Mono',monospace;font-size:.82rem}
.btn-outline:hover{border-color:var(--accent);color:var(--acc2)}
.btn-full{width:100%;margin-top:14px}

.prog-bar{display:none;margin-top:16px;background:var(--surface);border:1px solid var(--border);border-radius:999px;height:7px;overflow:hidden}
.prog-fill{height:100%;width:0%;background:linear-gradient(90deg,var(--accent),var(--acc2));border-radius:999px;transition:width .4s ease}

.result-card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:22px;margin-bottom:14px}
.card-hdr{display:flex;align-items:center;gap:9px;margin-bottom:14px;font-family:'Syne',sans-serif;font-size:.95rem;font-weight:700}
.badge{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:999px;font-size:.72rem;font-family:'DM Mono',monospace}
.bg{background:rgba(34,211,165,.1);color:var(--green);border:1px solid rgba(34,211,165,.22)}
.br{background:rgba(248,113,113,.1);color:var(--red);border:1px solid rgba(248,113,113,.22)}

.flist{display:flex;flex-direction:column;gap:5px}
.fitem{display:flex;align-items:center;gap:9px;padding:7px 11px;border-radius:8px;font-family:'DM Mono',monospace;font-size:.78rem;background:var(--surface)}
.dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.dg{background:var(--green);box-shadow:0 0 6px var(--green)}
.dr{background:var(--red);box-shadow:0 0 6px var(--red)}
.fname{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}

.stats{display:flex;gap:10px;margin-bottom:22px}
.stat{flex:1;background:var(--card);border:1px solid var(--border);border-radius:13px;padding:16px;text-align:center}
.stat-n{font-family:'Syne',sans-serif;font-size:2rem;font-weight:800}
.stat-l{font-size:.72rem;color:var(--muted);font-family:'DM Mono',monospace;margin-top:2px}
.s-acc .stat-n{color:var(--accent)}
.s-g .stat-n{color:var(--green)}
.s-r .stat-n{color:var(--red)}

.retry-banner{background:rgba(248,113,113,.07);border:1px solid rgba(248,113,113,.18);border-radius:12px;padding:15px 18px;display:flex;align-items:center;gap:11px;margin-bottom:14px}
.rb-title{font-family:'Syne',sans-serif;font-weight:700;font-size:.92rem;color:var(--red)}
.rb-sub{font-size:.78rem;color:var(--muted);font-family:'DM Mono',monospace;margin-top:2px}

.info-box{background:rgba(99,102,241,.07);border:1px solid rgba(99,102,241,.2);border-radius:12px;padding:16px 20px;margin-bottom:20px}
.info-box strong{font-family:'Syne',sans-serif;display:block;font-size:.95rem;margin-bottom:4px;color:#fff}
.info-box span{font-family:'DM Mono',monospace;font-size:.82rem;color:var(--acc2)}

.excel-result{background:rgba(34,211,165,.06);border:1px solid rgba(34,211,165,.2);border-radius:14px;padding:28px 24px;text-align:center;margin-top:20px}
.excel-title{font-family:'Syne',sans-serif;font-size:1.15rem;font-weight:700;color:var(--green);margin-bottom:4px}
.excel-sub{font-size:.8rem;color:var(--muted);font-family:'DM Mono',monospace;margin-top:6px}

.spin{display:inline-block;width:15px;height:15px;border:2px solid rgba(255,255,255,.25);border-top-color:#fff;border-radius:50%;animation:rot .7s linear infinite;vertical-align:middle}
@keyframes rot{to{transform:rotate(360deg)}}
.warn-box{background:rgba(251,191,36,.07);border:1px solid rgba(251,191,36,.2);border-radius:10px;padding:12px 16px;font-size:.8rem;font-family:'DM Mono',monospace;color:var(--yellow);margin-top:14px}
.action-row{display:flex;gap:10px;margin-top:16px;flex-wrap:wrap;align-items:center}
</style>
</head>
<body>
<div class="glow"></div>
<div class="wrap">
  <header>
    <div class="logo-box">&#128179;</div>
    <h1>Credit Card PDF Suite</h1>
    <p class="tagline">// unlock passwords &middot; extract transactions &middot; consolidate all cards to excel</p>
    <div class="bank-chips">
      <span class="chip">HDFC</span>
      <span class="chip">SBI Card</span>
      <span class="chip">ICICI</span>
      <span class="chip">Axis</span>
      <span class="chip">Kotak</span>
      <span class="chip">Citi</span>
      <span class="chip">IDFC</span>
      <span class="chip">AmEx</span>
      <span class="chip">IndusInd</span>
      <span class="chip">+ any bank</span>
    </div>
  </header>

  <div class="tabs">
    <button type="button" class="tab active" id="tab1" onclick="switchTab(1)">
      <span class="tab-badge" id="tbadge1">1</span>Unlock PDFs
    </button>
    <button type="button" class="tab" id="tab2" onclick="switchTab(2)">
      <span class="tab-badge" id="tbadge2">2</span>Consolidate to Excel
    </button>
  </div>

  <!-- PANEL 1: Unlock -->
  <div class="panel active" id="panel1">
    <div class="dropzone" id="dz1" onclick="document.getElementById('fi1').click()">
      <div class="drop-icon">&#128274;</div>
      <div class="drop-title">Upload Password-Protected PDFs</div>
      <div class="drop-sub">Drag &amp; drop or click &mdash; any bank, multiple files at once</div>
      <input type="file" id="fi1" accept=".pdf" multiple>
    </div>
    <div class="pill-wrap" id="pills1"></div>

    <div class="pw-row">
      <div class="pw-wrap">
        <input type="password" id="pw1" placeholder="Enter PDF password (e.g. DOB: 01011990)">
        <button type="button" class="eye" onclick="toggleEye('pw1')">&#128065;</button>
      </div>
      <button type="button" class="btn btn-primary" id="unlockBtn" onclick="doUnlock()">&#128275; Unlock All</button>
    </div>
    <div class="prog-bar" id="prog1"><div class="prog-fill" id="pfill1"></div></div>

    <div id="unlockResults" style="display:none;margin-top:26px">
      <div class="stats">
        <div class="stat s-acc"><div class="stat-n" id="sTotal">0</div><div class="stat-l">total</div></div>
        <div class="stat s-g"><div class="stat-n" id="sOk">0</div><div class="stat-l">unlocked</div></div>
        <div class="stat s-r"><div class="stat-n" id="sFail">0</div><div class="stat-l">failed</div></div>
      </div>

      <div class="result-card" id="okCard" style="display:none">
        <div class="card-hdr"><span class="badge bg">&#10003; Unlocked</span><span id="okCount"></span></div>
        <div class="flist" id="okList"></div>
        <button type="button" class="btn btn-green btn-full" onclick="downloadUnlocked()">&#11015; Download Unlocked PDFs</button>
      </div>

      <div id="retryWrap" style="display:none">
        <div class="retry-banner">
          <span style="font-size:1.4rem">&#128272;</span>
          <div>
            <div class="rb-title">Some files need a different password</div>
            <div class="rb-sub" id="retrySubText"></div>
          </div>
        </div>
        <div class="result-card">
          <div class="card-hdr"><span class="badge br">&#10007; Wrong Password</span><span id="failCount"></span></div>
          <div class="flist" id="failList"></div>
          <div class="pw-row" style="margin-top:16px">
            <div class="pw-wrap">
              <input type="password" id="pw2" placeholder="Try a different password...">
              <button type="button" class="eye" onclick="toggleEye('pw2')">&#128065;</button>
            </div>
            <button type="button" class="btn btn-primary" onclick="doRetry()">&#128260; Retry</button>
          </div>
        </div>
      </div>

      <div class="action-row">
        <button type="button" class="btn btn-primary" onclick="goConsolidate()" id="goConsBtn" style="display:none">
          &#128202; Consolidate to Excel &#8594;
        </button>
        <button type="button" class="btn btn-outline" onclick="resetAll()">&#8635; Start Over</button>
      </div>
    </div>
  </div>

  <!-- PANEL 2: Consolidate -->
  <div class="panel" id="panel2">
    <div id="step2FromSession" style="display:none">
      <div class="info-box">
        <strong>&#10003; Using your unlocked PDFs from Step 1</strong>
        <span id="step2SessionInfo">Ready to consolidate</span>
      </div>
    </div>

    <div id="step2Fresh">
      <p style="color:var(--muted);font-size:.88rem;margin-bottom:14px;font-family:'DM Mono',monospace">
        Already have unlocked PDFs? Upload them directly. Or complete Step 1 first to remove passwords.
      </p>
      <div class="dropzone" id="dz2" onclick="document.getElementById('fi2').click()">
        <div class="drop-icon">&#128194;</div>
        <div class="drop-title">Upload Credit Card PDFs</div>
        <div class="drop-sub">Mix of any banks &mdash; HDFC, SBI, ICICI, Axis and more</div>
        <input type="file" id="fi2" accept=".pdf" multiple>
      </div>
      <div class="pill-wrap" id="pills2"></div>
    </div>

    <button type="button" class="btn btn-primary btn-full" id="consolidateBtn" onclick="doConsolidate()" style="margin-top:24px;padding:16px;font-size:1rem">
      &#128202; Generate Consolidated Excel
    </button>
    <div class="prog-bar" id="prog2"><div class="prog-fill" id="pfill2"></div></div>

    <div class="warn-box">
      &#9888; Works best with digital (text-based) PDFs. Scanned image PDFs will use OCR &mdash; slower and may vary in accuracy.
    </div>

    <div id="consolidateResult" style="display:none;margin-top:22px">
      <div class="excel-result">
        <div style="font-size:3rem;margin-bottom:10px">&#128994;</div>
        <div class="excel-title">Excel Ready!</div>
        <div class="excel-sub" id="xlsxMeta"></div>
        <button type="button" class="btn btn-green btn-full" style="max-width:340px;margin:18px auto 0" onclick="downloadExcel()">
          &#11015; Download Consolidated Excel
        </button>
      </div>
      <div id="parseErrors" style="display:none" class="warn-box"></div>
    </div>
  </div>

</div>

<script>
var sessionId=null, downloadId=null, files1={}, freshFiles2={};

function switchTab(n){
  var tabs=['tab1','tab2'];
  var panels=['panel1','panel2'];
  tabs.forEach(function(id,i){
    var el=document.getElementById(id);
    if(el) el.classList.toggle('active',i===n-1);
  });
  panels.forEach(function(id,i){
    var el=document.getElementById(id);
    if(el) el.classList.toggle('active',i===n-1);
  });
}
function goConsolidate(){
  var total=parseInt(document.getElementById('sOk').textContent)||0;
  document.getElementById('step2SessionInfo').textContent=total+' file'+(total!==1?'s':'')+' ready to consolidate';
  document.getElementById('step2FromSession').style.display='block';
  document.getElementById('step2Fresh').style.display='none';
  switchTab(2);
  setTimeout(function(){
    var btn=document.getElementById('consolidateBtn');
    if(btn) btn.scrollIntoView({behavior:'smooth',block:'center'});
  },100);
}
function setupDrop(dzId,inputId,pillsId,store){
  var dz=document.getElementById(dzId), inp=document.getElementById(inputId);
  dz.addEventListener('dragover',function(e){e.preventDefault();dz.classList.add('over');});
  dz.addEventListener('dragleave',function(){dz.classList.remove('over');});
  dz.addEventListener('drop',function(e){e.preventDefault();dz.classList.remove('over');addFiles(e.dataTransfer.files,pillsId,store);});
  inp.addEventListener('change',function(){addFiles(inp.files,pillsId,store);});
}
function addFiles(list,pillsId,store){
  for(var i=0;i<list.length;i++){if(list[i].name.toLowerCase().endsWith('.pdf'))store[list[i].name]=list[i];}
  renderPills(pillsId,store);
}
function renderPills(pillsId,store){
  var wrap=document.getElementById(pillsId); wrap.innerHTML='';
  Object.keys(store).forEach(function(name){
    var p=document.createElement('div'); p.className='pill';
    var s=name.replace(/\\/g,'\\\\').replace(/'/g,"\\'");
    p.innerHTML='&#128196; <span style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+name+'">'+name+'</span><span class="x" onclick="removeFile(\''+s+'\',\''+pillsId+'\')">&#215;</span>';
    wrap.appendChild(p);
  });
}
function removeFile(name,pillsId){var store=pillsId==='pills1'?files1:freshFiles2;delete store[name];renderPills(pillsId,store);}
function toggleEye(id){var inp=document.getElementById(id);inp.type=inp.type==='password'?'text':'password';}

function doUnlock(){
  var flist=Object.values(files1);
  if(!flist.length){alert('Please upload at least one PDF first.');return;}
  sessionId=null; runUnlock(flist,document.getElementById('pw1').value);
}
function doRetry(){runUnlock([],document.getElementById('pw2').value);}
function runUnlock(flist,pw){
  var btn=document.getElementById('unlockBtn');
  btn.disabled=true; btn.innerHTML='<span class="spin"></span> Unlocking...';
  showProg('prog1','pfill1',20);
  var fd=new FormData();
  fd.append('password',pw);
  if(sessionId)fd.append('session_id',sessionId);
  flist.forEach(function(f){fd.append('pdfs',f);});
  fetch('/unlock',{method:'POST',body:fd})
    .then(function(r){return r.json();})
    .then(function(data){
      showProg('prog1','pfill1',100);
      setTimeout(function(){hideProg('prog1','pfill1');},500);
      sessionId=data.session_id;
      renderUnlockResults(data);
    })
    .catch(function(e){alert('Error: '+e.message);})
    .finally(function(){btn.disabled=false;btn.innerHTML='&#128275; Unlock All';});
}
function renderUnlockResults(data){
  document.getElementById('unlockResults').style.display='block';
  document.getElementById('sTotal').textContent=data.total_unlocked+data.locked.length;
  document.getElementById('sOk').textContent=data.total_unlocked;
  document.getElementById('sFail').textContent=data.locked.length;
  if(data.total_unlocked>0){
    document.getElementById('okCard').style.display='block';
    document.getElementById('okCount').textContent=data.total_unlocked+' file'+(data.total_unlocked!==1?'s':'')+' ready';
    var ol=document.getElementById('okList');
    data.unlocked.forEach(function(name){
      var el=document.createElement('div');el.className='fitem';
      el.innerHTML='<span class="dot dg"></span><span class="fname">'+name+'</span><span class="badge bg" style="font-size:.68rem">unlocked</span>';
      ol.appendChild(el);
    });
    document.getElementById('goConsBtn').style.display='inline-flex';
    document.getElementById('tab2').classList.add('done');
  }
  if(data.locked.length>0){
    document.getElementById('retryWrap').style.display='block';
    document.getElementById('retrySubText').textContent=data.locked.length+' file'+(data.locked.length!==1?'s':'')+' could not be unlocked — try a different password';
    document.getElementById('failCount').textContent=data.locked.length+' file'+(data.locked.length!==1?'s':'');
    var fl=document.getElementById('failList');fl.innerHTML='';
    data.locked.forEach(function(name){
      var el=document.createElement('div');el.className='fitem';
      el.innerHTML='<span class="dot dr"></span><span class="fname">'+name+'</span><span class="badge br" style="font-size:.68rem">wrong password</span>';
      fl.appendChild(el);
    });
  } else {document.getElementById('retryWrap').style.display='none';}
  document.getElementById('unlockResults').scrollIntoView({behavior:'smooth',block:'start'});
}
function downloadUnlocked(){if(sessionId)window.location.href='/download/unlocked/'+sessionId;}

function doConsolidate(){
  var btn=document.getElementById('consolidateBtn');
  btn.disabled=true;btn.innerHTML='<span class="spin"></span> Extracting transactions...';
  showProg('prog2','pfill2',15);
  var fd=new FormData();
  var fromSession=document.getElementById('step2FromSession').style.display!=='none';
  if(fromSession&&sessionId){
    fd.append('session_id',sessionId);
  } else {
    var flist=Object.values(freshFiles2);
    if(!flist.length){alert('Please upload PDF files first.');btn.disabled=false;btn.innerHTML='&#128202; Generate Consolidated Excel';hideProg('prog2','pfill2');return;}
    flist.forEach(function(f){fd.append('pdfs',f);});
  }
  showProg('prog2','pfill2',50);
  fetch('/consolidate',{method:'POST',body:fd})
    .then(function(r){return r.json();})
    .then(function(data){
      showProg('prog2','pfill2',100);
      setTimeout(function(){hideProg('prog2','pfill2');},500);
      if(data.error){alert('Error: '+data.error);return;}
      downloadId=data.download_id;
      document.getElementById('consolidateResult').style.display='block';
      document.getElementById('xlsxMeta').textContent=data.total_transactions+' transactions from '+data.files_processed+' file'+(data.files_processed!==1?'s':', across multiple banks');
      if(data.parse_errors&&data.parse_errors.length){
        var eb=document.getElementById('parseErrors');eb.style.display='block';
        eb.textContent='Note: some files had issues — '+data.parse_errors.join(' | ');
      }
      document.getElementById('consolidateResult').scrollIntoView({behavior:'smooth',block:'start'});
    })
    .catch(function(e){alert('Error: '+e.message);})
    .finally(function(){btn.disabled=false;btn.innerHTML='&#128202; Generate Consolidated Excel';});
}
function downloadExcel(){if(downloadId)window.location.href='/download/excel/'+downloadId;}

function showProg(b,f,p){document.getElementById(b).style.display='block';document.getElementById(f).style.width=p+'%';}
function hideProg(b,f){document.getElementById(b).style.display='none';document.getElementById(f).style.width='0%';}

function resetAll(){
  if(sessionId)fetch('/clear/'+sessionId,{method:'POST'});
  sessionId=null;downloadId=null;files1={};freshFiles2={};
  setupDrop('dz1','fi1','pills1',files1);
  setupDrop('dz2','fi2','pills2',freshFiles2);
  ['pills1','pills2','okList','failList'].forEach(function(id){document.getElementById(id).innerHTML='';});
  ['pw1','pw2'].forEach(function(id){document.getElementById(id).value='';});
  ['unlockResults','okCard','retryWrap','consolidateResult','parseErrors','step2FromSession'].forEach(function(id){document.getElementById(id).style.display='none';});
  document.getElementById('step2Fresh').style.display='block';
  document.getElementById('goConsBtn').style.display='none';
  document.getElementById('tab2').classList.remove('done');
  switchTab(1);
  window.scrollTo({top:0,behavior:'smooth'});
}
// Tab clicks handled via onclick on the button elements above
setupDrop('dz1','fi1','pills1',files1);
setupDrop('dz2','fi2','pills2',freshFiles2);
</script>
</body>
</html>"""

# ─── ROUTES ──────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return HTML   # No render_template — works on Render without templates/ folder

@app.route('/unlock', methods=['POST'])
def unlock():
    password = request.form.get('password', '')
    session_id = request.form.get('session_id', str(uuid.uuid4()))
    files = request.files.getlist('pdfs')
    has_session_data = session_id in SESSIONS and SESSIONS[session_id].get('locked_data')
    if not files and not has_session_data:
        return jsonify({'error': 'No files uploaded'}), 400
    pending = {}
    if has_session_data:
        pending = dict(SESSIONS[session_id]['locked_data'])
    for f in files:
        if f.filename: pending[f.filename] = f.read()
    unlocked_data = SESSIONS.get(session_id, {}).get('unlocked_data', {})
    newly_locked, unlocked, locked = {}, [], []
    for filename, file_bytes in pending.items():
        try:
            reader = PdfReader(io.BytesIO(file_bytes))
            if reader.is_encrypted:
                if reader.decrypt(password) == 0:
                    newly_locked[filename] = file_bytes; locked.append(filename); continue
            writer = PdfWriter()
            for page in reader.pages: writer.add_page(page)
            buf = io.BytesIO(); writer.write(buf)
            unlocked_data[filename] = buf.getvalue(); unlocked.append(filename)
        except Exception:
            newly_locked[filename] = file_bytes; locked.append(filename)
    SESSIONS[session_id] = {'unlocked_data': unlocked_data, 'locked_data': newly_locked}
    return jsonify({'session_id': session_id, 'unlocked': unlocked, 'locked': locked, 'total_unlocked': len(unlocked_data)})

@app.route('/consolidate', methods=['POST'])
def consolidate():
    session_id = request.form.get('session_id', '')
    files = request.files.getlist('pdfs')
    pdf_sources = {}
    if session_id and session_id in SESSIONS:
        pdf_sources = dict(SESSIONS[session_id].get('unlocked_data', {}))
    for f in files:
        if f.filename:
            data = f.read()
            try:
                reader = PdfReader(io.BytesIO(data))
                if reader.is_encrypted:
                    return jsonify({'error': f'{f.filename} is still password-protected. Please unlock it in Step 1 first.'}), 400
            except Exception:
                pass
            pdf_sources[f.filename] = data
    if not pdf_sources:
        return jsonify({'error': 'No PDFs found. Please upload files or complete Step 1 first.'}), 400
    all_transactions, account_infos, parse_errors = [], [], []
    for filename, pdf_bytes in pdf_sources.items():
        try:
            info, txns = parse_pdf(pdf_bytes, filename)
            account_infos.append(info); all_transactions.extend(txns)
        except Exception as e:
            parse_errors.append(f'{filename}: {str(e)}')
    if not all_transactions:
        return jsonify({'error': 'No transactions could be extracted. PDFs may be scanned images or in an unsupported format.', 'parse_errors': parse_errors}), 422
    xlsx_bytes = build_excel(all_transactions, account_infos)
    dl_id = str(uuid.uuid4())
    SESSIONS[dl_id] = {'xlsx': xlsx_bytes}
    return jsonify({'download_id': dl_id, 'total_transactions': len(all_transactions), 'files_processed': len(account_infos), 'parse_errors': parse_errors})

@app.route('/download/excel/<dl_id>')
def download_excel(dl_id):
    if dl_id not in SESSIONS or 'xlsx' not in SESSIONS[dl_id]:
        return jsonify({'error': 'File not found — please regenerate.'}), 404
    return send_file(io.BytesIO(SESSIONS[dl_id]['xlsx']),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name='Consolidated_Credit_Card_Statement.xlsx')

@app.route('/download/unlocked/<session_id>')
def download_unlocked(session_id):
    if session_id not in SESSIONS:
        return jsonify({'error': 'Session not found'}), 404
    unlocked_data = SESSIONS[session_id].get('unlocked_data', {})
    if not unlocked_data:
        return jsonify({'error': 'No unlocked files'}), 404
    if len(unlocked_data) == 1:
        fname, data = next(iter(unlocked_data.items()))
        return send_file(io.BytesIO(data), mimetype='application/pdf', as_attachment=True, download_name=f'unlocked_{fname}')
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname, data in unlocked_data.items():
            zf.writestr(f'unlocked_{fname}', data)
    zip_buf.seek(0)
    return send_file(zip_buf, mimetype='application/zip', as_attachment=True, download_name='unlocked_statements.zip')

@app.route('/clear/<session_id>', methods=['POST'])
def clear(session_id):
    SESSIONS.pop(session_id, None)
    return jsonify({'ok': True})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
