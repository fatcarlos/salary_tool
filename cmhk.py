import os
import json
from datetime import date, datetime, timedelta
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, session, send_file
import pandas as pd

app = Flask(__name__)
app.secret_key = os.urandom(24)

app.config['SESSION_COOKIE_NAME'] = 'cmhk_session'
app.config['SESSION_COOKIE_SAMESITE'] = 'None'
app.config['SESSION_COOKIE_SECURE'] = True

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "salary_cmhk.json")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

DEFAULT_HOURS = 8
DEFAULT_RATE = 125

PORTAL_ACCESS_COOKIE = "portal_access"
PORTAL_SECRET = os.environ.get("IFRAME_SECRET_KEY", "")

def _portal_access_valid():
    token = request.cookies.get(PORTAL_ACCESS_COOKIE, "")
    if not token or not PORTAL_SECRET:
        return False
    import hmac, hashlib
    expected = hmac.new(PORTAL_SECRET.encode(), b"salary-portal", hashlib.sha256).hexdigest()
    return hmac.compare_digest(token, expected)

@app.before_request
def require_portal_login():
    if not _portal_access_valid():
        return "需要先由 iframe.py 登入。", 401
WEEKDAY_CN = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"]
WEEKDAY_EXCEL = ["一", "二", "三", "四", "五", "六", "日"]

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def resolve_note_form(form):
                                         
    choice = form.get("note_type")
    if choice is None:
        return str(form.get("note", "")).strip()
    if str(choice).strip().lower() == "other":
        return str(form.get("custom_note", "")).strip()
    allowed = {"null": "null", "paid": "Paid", "off": "Off"}
    return allowed.get(str(choice).strip().lower(), str(choice).strip())

def calc_record(r):
    note = str(r.get("note", "")).strip().lower()
    if note == "off":
        return 0, 0
    try:
        hours = int(float(r.get("hours", DEFAULT_HOURS) or 0))
        rate = float(r.get("rate", DEFAULT_RATE) or 0)
    except (TypeError, ValueError):
        return 0, 0
    pay = int(hours * rate)
    if pay == 0:
        return 0, 0
    return hours, pay

def get_pay_period(data=None, period="current"):
\
\
\
\
\
\
\
       
    today = date.today()

    if period == "previous":
        if today.month == 1:
            year, month = today.year - 1, 12
        else:
            year, month = today.year, today.month - 1
    else:
        year, month = today.year, today.month

    start = date(year, month, 1)
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)
    end = next_month - timedelta(days=1)
    return start, end

def filter_period_records(data, period="current"):
    start, end = get_pay_period(data, period)
    records = {}
    total_h = total_t = work_days = 0
    for r in data:
        try:
            d = datetime.strptime(str(r.get("date", "")), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            continue
        if start <= d <= end:
            h, t = calc_record(r)
            records[d] = (h, t, str(r.get("note", "")))
            total_h += h
            total_t += t
            if h > 0 and t > 0:
                work_days += 1
    return start, end, records, total_h, total_t, work_days

def generate_excel(data, period="current"):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

    start_date, end_date, records, total_h, total_t, work_days = filter_period_records(data, period)
    wb = Workbook()
    ws = wb.active
    ws.title = "工資表"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 30

    center = Alignment(horizontal="center", vertical="center")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    def set_cell(cell, value, bg=None):
        c = ws[cell]
        c.value = value
        c.alignment = center
        c.border = border
        if bg:
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")

    row = 1
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"] = f"每月工資表（{start_date} ～ {end_date}）"
    ws[f"A{row}"].alignment = center
    ws[f"A{row}"].font = Font(size=14, bold=True)

    row += 2
    set_cell(f"A{row}", "工作日數", "E8F5E9")
    set_cell(f"B{row}", work_days)
    set_cell(f"C{row}", "總工資", "E8F5E9")
    set_cell(f"D{row}", f"$ {int(total_t)}", "FFF45C")

    row += 2
    for col, h in zip(["A", "B", "C", "D"], ["日期", "星期", "工時", "每日工資"]):
        set_cell(f"{col}{row}", h, "7ED957")

    row += 1
    for d in sorted(records.keys()):
        h, t, _ = records[d]
        values = [d.strftime("%m/%d"), f"星期{WEEKDAY_EXCEL[d.weekday()]}", int(h), f"$ {int(t)}"]
        for col, v in zip(["A", "B", "C", "D"], values):
            set_cell(f"{col}{row}", v)
        row += 1

    ws.merge_cells(f"A{row}:B{row}")
    set_cell(f"A{row}", f"總工時：{int(total_h)}", "E3F2FD")
    ws.merge_cells(f"C{row}:D{row}")
    set_cell(f"C{row}", f"總工資：$ {int(total_t)}", "FFF9C4")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_pdf(data, period="current"):
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.lib.styles import getSampleStyleSheet

    FONT_NAME = "MSung-Light"
    pdfmetrics.registerFont(UnicodeCIDFont(FONT_NAME))

    start, end, records, _, total_pay, work_days = filter_period_records(data, period)
    rows = [["Date", "Day", "Hours", "Daily Pay"]]
    for d in sorted(records.keys()):
        h, p, _ = records[d]
        if h > 0:
            rows.append([d.strftime("%m/%d"), d.strftime("%a"), str(int(h)), f"$ {int(p)}"])

    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4)
    styles = getSampleStyleSheet()
    styles["Normal"].fontName = FONT_NAME

    elements = [
        Paragraph(f"Salary Report ({start} to {end})", styles["Normal"]),
        Paragraph(f"Working Days: {work_days} | Total Pay: $ {int(total_pay)}", styles["Normal"]),
        Table(rows, colWidths=[100]*4, style=TableStyle([
            ("FONT", (0, 0), (-1, -1), FONT_NAME),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7ED957")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ]))
    ]
    doc.build(elements)
    output.seek(0)
    return output

def generate_jpg(data, period="current"):
    from PIL import Image, ImageDraw, ImageFont

    start_date, end_date, records, _, total_t, work_days = filter_period_records(data, period)
    filtered_days = [d for d in sorted(records.keys()) if records[d][0] > 0 and records[d][1] > 0]

    row_h = 55
    img_h = 240 + row_h * max(len(filtered_days), 1)
    font_paths = [os.path.join(BASE_DIR, "NotoSansCJK-Regular.ttc")]
    font = None
    for path in font_paths:
        if os.path.exists(path):
            try:
                font = ImageFont.truetype(path, 26)
                break
            except Exception:
                pass
    if font is None:
        font = ImageFont.load_default()

    img = Image.new("RGB", (760, img_h), "white")
    draw = ImageDraw.Draw(img)

    def draw_cell(x, y, w, h, text, bg="#FFFFFF", border="#000000"):
        draw.rectangle([x, y, x + w, y + h], fill=bg, outline=border)
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((x + (w - tw) // 2, y + (h - th) // 2), text, fill="black", font=font)

    x0, y = 20, 20
    draw.text((x0, y), f"每月工資表（{start_date} ～ {end_date}）", fill="black", font=font)
    y += 60

    draw_cell(x0, y, 160, row_h, "工作日數", bg="#E8F5E9")
    draw_cell(x0 + 160, y, 120, row_h, str(work_days))
    draw_cell(x0 + 280, y, 160, row_h, "總工資", bg="#E8F5E9")
    draw_cell(x0 + 440, y, 260, row_h, f"$ {int(total_t)}", bg="#FFF45C")
    y += row_h + 20

    col_w = [150, 150, 150, 250]
    x = x0
    for i, h_text in enumerate(["日期", "星期", "工時 (Hrs)", "每日工資"]):
        draw_cell(x, y, col_w[i], row_h, h_text, bg="#7ED957")
        x += col_w[i]
    y += row_h

    for d in filtered_days:
        h, t, _ = records[d]
        values = [d.strftime("%m/%d"), f"星期{WEEKDAY_EXCEL[d.weekday()]}", str(int(h)), f"$ {int(t)}"]
        x = x0
        for i, v in enumerate(values):
            draw_cell(x, y, col_w[i], row_h, v)
            x += col_w[i]
        y += row_h

    output = BytesIO()
    img.save(output, format="JPEG", quality=95)
    output.seek(0)
    return output

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-HK">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📶CMHK📡</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #e3f2fd 0%, #f0f7f4 50%, #e8f5e9 100%);
            background-attachment: fixed;
            color: #2c3e50;
            margin: 0;
            padding-top: 50px;
            transition: opacity 0.6s ease, transform 0.6s ease, background 0.5s ease, color 0.5s ease;
        }

        body.dark-mode {
            background: linear-gradient(135deg, #0d1b1e 0%, #111a14 50%, #1a241b 100%) !important;
            color: #e2e8f0 !important;
        }
        /* Mobile home title: hide the trailing antenna emoji on phones only. */
        @media (max-width: 600px) {
            .cmhk-title-mobile-hide { display: none !important; }
        }
        body.dark-mode .top-console {
            background: rgba(18, 24, 20, 0.85) !important;
            border-bottom: 1px solid rgba(76, 175, 80, 0.3) !important;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3) !important;
        }
        /* Monthly speech button: maintain contrast and visibility in dark mode. */
        body.dark-mode #cmhkMonthlySpeechBtn {
            background: rgba(24, 38, 30, 0.96) !important;
            color: #ffffff !important;
            border: 1px solid rgba(76, 175, 80, 0.85) !important;
            box-shadow: 0 0 10px rgba(76, 175, 80, 0.28) !important;
            opacity: 1 !important;
            filter: none !important;
        }
        body.dark-mode #cmhkMonthlySpeechStopBtn {
            background: rgba(58, 25, 28, 0.96) !important;
            color: #ffffff !important;
            border: 1px solid rgba(248, 113, 113, 0.9) !important;
            box-shadow: 0 0 10px rgba(248, 113, 113, 0.25) !important;
        }
        body.dark-mode #cmhkMonthlySpeechStopBtn:hover,
        body.dark-mode #cmhkMonthlySpeechStopBtn:focus-visible {
            background: rgba(88, 32, 36, 1) !important;
            box-shadow: 0 0 14px rgba(248, 113, 113, 0.45) !important;
            outline: 2px solid rgba(248, 113, 113, 0.55);
            outline-offset: 2px;
        }
        body.dark-mode #cmhkMonthlySpeechBtn:hover,
        body.dark-mode #cmhkMonthlySpeechBtn:focus-visible {
            background: rgba(38, 62, 46, 1) !important;
            border-color: #81c784 !important;
            box-shadow: 0 0 14px rgba(129, 199, 132, 0.5) !important;
            outline: 2px solid rgba(129, 199, 132, 0.65);
            outline-offset: 2px;
        }

        body.dark-mode .top-console input,
        body.dark-mode .top-console select,
        body.dark-mode .top-console button {
            background: rgba(20, 30, 24, 0.6) !important;
            color: #e2e8f0 !important;
            border: 1px solid rgba(76, 175, 80, 0.4) !important;
        }
        body.dark-mode .card {
            background: rgba(22, 32, 26, 0.45) !important;
            border: 1px solid rgba(76, 175, 80, 0.25) !important;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.35) !important;
            color: #e2e8f0 !important;
        }
        body.dark-mode .stat-box {
            background: rgba(30, 50, 38, 0.4) !important;
            border: 1px solid rgba(76, 175, 80, 0.3) !important;
            color: #e2e8f0 !important;
            border-left: 5px solid #4caf50;
        }
        body.dark-mode table {
            background: rgba(20, 30, 24, 0.35) !important;
            color: #e2e8f0 !important;
        }
        body.dark-mode .calendar-card { background: rgba(20,28,24,0.32) !important; border-color: rgba(255,255,255,0.20) !important; }
        body.dark-mode .calendar-card table { background: rgba(20,30,24,0.22) !important; }
        body.dark-mode .calendar-card th { background: rgba(255,255,255,0.08) !important; }
        body.dark-mode th, body.dark-mode td {
            border: 1px solid rgba(76, 175, 80, 0.2) !important;
        }
        body.dark-mode th {
            background-color: rgba(30, 50, 38, 0.6) !important;
            color: #81c784 !important;
        }
        body.dark-mode details summary {
            color: #81c784 !important;
        }
        body.dark-mode details summary:hover {
            color: #a5d6a7 !important;
        }
        body.dark-mode .login-card {
            background: rgba(22, 32, 26, 0.65) !important;
            border: none !important;
            box-shadow: none !important;
        }
        body.dark-mode .login-card h2 {
            color: #a5d6a7 !important;
        }
        body.dark-mode .login-card input,
        body.dark-mode .login-card select {
            background: rgba(20, 30, 24, 0.8) !important;
            border: 2px solid rgba(76, 175, 80, 0.5) !important;
            color: #e2e8f0 !important;
        }
        body.dark-mode .system-switch-link {
            color: #81c784 !important;
            background: rgba(76, 175, 80, 0.15) !important;
            border: 1px solid rgba(76, 175, 80, 0.3) !important;
        }
        body.dark-mode .system-switch-link:hover {
            background: rgba(76, 175, 80, 0.3) !important;
        }

        body.fade-out {
            opacity: 0;
            transform: scale(1.05);
        }

        .top-console {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border-bottom: 1px solid rgba(163, 217, 177, 0.6);
            box-shadow: 0 4px 20px rgba(46, 125, 50, 0.08);
            z-index: 9999;
            padding: 0 15px;
            box-sizing: border-box;
        }
        .top-console summary::-webkit-details-marker {
            display: none;
        }
        /* Match the Cargo control title pill: same glass finish, typography and depth. */
        #consoleDetails > summary > span:first-of-type {
            font-family: inherit !important;
            font-size: 0.85rem !important;
            font-weight: 700 !important;
            letter-spacing: 0 !important;
            line-height: 1.2 !important;
            color: #ffffff !important;
            text-shadow: 0 1px 2px rgba(0,0,0,.22);
        }
        #consoleDetails > summary > span:first-of-type > span {
            margin-left: 0.25rem;
        }
        #consoleDetails > summary > span:first-of-type {
            padding: 5px 15px !important;
            border-radius: 999px !important;
            background: linear-gradient(135deg, rgba(70,85,78,.94), rgba(28,40,33,.88)) !important;
            border: 1px solid rgba(220,255,232,.40) !important;
            box-shadow: 0 4px 14px rgba(0,0,0,.24), inset 0 1px 1px rgba(255,255,255,.48) !important;
            -webkit-backdrop-filter: blur(18px) saturate(180%);
            backdrop-filter: blur(18px) saturate(180%);
        }
        [data-bs-theme="light"] #consoleDetails > summary > span:first-of-type,
        body.light-mode #consoleDetails > summary > span:first-of-type {
            color: #111111 !important;
            background: linear-gradient(135deg, rgba(255,255,255,.96), rgba(230,240,235,.86)) !important;
            border-color: rgba(255,255,255,.95) !important;
            box-shadow: 0 4px 14px rgba(20,40,30,.16), inset 0 1px 1px rgba(255,255,255,.95) !important;
            text-shadow: none !important;
        }
        body.salary-3d-mode #consoleDetails > summary > span:first-of-type {
            transform: translateZ(12px);
            box-shadow: 0 7px 0 rgba(0,0,0,.22), 0 12px 20px rgba(0,0,0,.30), inset 0 2px 0 rgba(255,255,255,.5) !important;
        }
        .top-console input, .top-console select, .top-console button {
            padding: 4px 6px;
            border: 1px solid rgba(200, 230, 201, 0.8);
            border-radius: 6px;
            font-size: 13px;
            background: rgba(252, 253, 252, 0.7);
            backdrop-filter: blur(4px);
        }

        .top-console button, .btn-success { background: linear-gradient(135deg, #4caf50, #388e3c); color: white; border: none; cursor: pointer; box-shadow: 0 2px 6px rgba(76, 175, 80, 0.2); }
        .top-console button:hover, .btn-success:hover { background: linear-gradient(135deg, #388e3c, #2e7d32); }
        .btn-danger { background: linear-gradient(135deg, #e57373, #d32f2f); color: white; border: none; cursor: pointer; box-shadow: 0 2px 6px rgba(229, 115, 115, 0.2); }
        .btn-danger:hover { background: linear-gradient(135deg, #d32f2f, #c62828); }

        /* iPad / desktop: use the full available viewport width for every main frame. */
        .container { width: 100%; max-width: none; margin: 0; padding: 10px 10px 15px 10px; box-sizing: border-box; }
        .container > .grid-2 { width: 100%; grid-template-columns: minmax(0, 1fr); gap: 12px; }
        .container .card { width: 100%; max-width: none; min-width: 0; box-sizing: border-box; }
        .container .grid-2 > div { width: 100%; min-width: 0; }
        .container table { width: 100%; }

        .system-switch-link {
            display: inline-block;
            margin-top: 6px;
            padding: 4px 10px;
            background: rgba(232, 245, 233, 0.7);
            border: 1px solid rgba(163, 217, 177, 0.8);
            border-radius: 6px;
            color: #2e7d32;
            text-decoration: none;
            font-size: 13px;
            font-weight: bold;
            backdrop-filter: blur(4px);
            transition: all 0.2s ease;
        }
        .system-switch-link:hover {
            background: rgba(200, 230, 201, 0.9);
            color: #1b5e20;
        }

        .card {
            background: rgba(255, 255, 255, 0.45);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            padding: 15px;
            border-radius: 12px;
            box-shadow: 0 8px 32px 0 rgba(46, 125, 50, 0.08);
            margin-bottom: 15px;
            border: 1px solid rgba(255, 255, 255, 0.6);
        }

        .grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
        .grid-2 { display: grid; grid-template-columns: 2fr 1fr; gap: 15px; }
        @media(max-width: 768px) {
            .grid-3 { grid-template-columns: 1fr; }
            .grid-2 { grid-template-columns: 1fr; }
        }

        .stat-box {
            background: rgba(232, 245, 233, 0.5);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            padding: 12px;
            border-radius: 8px;
            border-left: 5px solid #4caf50;
            border: 1px solid rgba(200, 230, 201, 0.4);
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            text-align: center;
            font-size: 14px;
            background: rgba(255, 255, 255, 0.3);
            backdrop-filter: blur(6px);
            border-radius: 8px;
            overflow: hidden;
        }
        th, td { padding: 8px; border: 1px solid rgba(200, 230, 201, 0.5); word-break: normal; overflow-wrap: normal; }
        /* Liquid Glass calendar */
        .calendar-card { background: rgba(255,255,255,0.20) !important; backdrop-filter: blur(24px) saturate(180%); -webkit-backdrop-filter: blur(24px) saturate(180%); border: 1px solid rgba(255,255,255,0.48) !important; box-shadow: 0 8px 32px rgba(31,38,135,0.14), inset 0 1px 0 rgba(255,255,255,0.35) !important; }
        .calendar-card table { background: rgba(255,255,255,0.12) !important; backdrop-filter: blur(16px) saturate(160%); -webkit-backdrop-filter: blur(16px) saturate(160%); border: 1px solid rgba(255,255,255,0.35); }
        .calendar-card th { background: rgba(255,255,255,0.26) !important; backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px); }
        .calendar-card td { background-color: rgba(255,255,255,0.10); border-color: rgba(255,255,255,0.30) !important; }
        .calendar-card td[style*="background-color: #FFD54F"] { background-color: #FFD54F !important; }
        .calendar-card td[style*="background-color: #388E3C"] { background-color: #388E3C !important; }
        .calendar-card td[style*="background-color: #4A2E2B"] { background-color: #4A2E2B !important; }
        .calendar-heading { display:flex; align-items:center; justify-content:space-between; gap:8px; flex-wrap:wrap; margin:0 0 10px 0; font-size:16px; font-weight:bold; }
        .calendar-heading .income-total { display:inline-block; font-size:16px; font-weight:bold; text-align:right; white-space:nowrap; padding:4px 8px; line-height:1.35; color:#664d03; background:#fff3cd; border:1px solid #ffecb5; border-radius:6px; box-shadow:0 2px 6px rgba(0,0,0,.12); }
        .calendar-pay { display:inline-block; white-space:nowrap; word-break:keep-all; overflow-wrap:normal; font-size:clamp(9px, 2.7vw, 11px); }
        .calendar-day-cell { cursor:pointer; transition:filter .15s ease; }
        .calendar-day-cell:hover { filter:brightness(1.12); }
        .calendar-day-cell:focus-visible { outline:2px solid #2196f3; outline-offset:-2px; }
        #cmhkDayPopup { position:fixed; z-index:10050; display:none; gap:7px; align-items:center; padding:8px; border-radius:13px;
            background:rgba(255,255,255,.94); border:1px solid rgba(100,160,120,.45); box-shadow:0 8px 28px rgba(0,0,0,.24);
            backdrop-filter:blur(18px); -webkit-backdrop-filter:blur(18px); }
        body.dark-mode #cmhkDayPopup { background:rgba(20,30,24,.96); border-color:rgba(129,199,132,.45); }
        #cmhkDayPopup { flex-direction:row !important; flex-wrap:nowrap !important; justify-content:center !important; }
        #cmhkDayPopup button[data-action] {
            box-sizing:border-box !important;
            display:inline-flex;
            flex:0 0 42px !important;
            width:42px !important;
            min-width:42px !important;
            max-width:42px !important;
            height:42px !important;
            min-height:42px !important;
            max-height:42px !important;
            margin:0 !important;
            padding:0 !important;
            align-items:center !important;
            justify-content:center !important;
            text-align:center !important;
            text-indent:0 !important;
            line-height:1 !important;
            letter-spacing:normal !important;
            font-family:system-ui, -apple-system, "Segoe UI Emoji", "Apple Color Emoji", sans-serif !important;
            font-size:21px !important;
            vertical-align:middle !important;
            border:1px solid rgba(120,150,130,.35) !important;
            border-radius:10px !important;
            background:rgba(120,180,130,.16) !important;
            cursor:pointer !important;
        }
        #cmhkDayPopup button[data-action]:hover { transform:translateY(-2px); background:rgba(76,175,80,.28) !important; }
        th { background-color: rgba(232, 245, 233, 0.7); color: #2e7d32; }
        .password-container input[type="password"] {
            flex: 1;
            margin: 0;
        }

        .alert-error { color: #d32f2f; background: rgba(255, 235, 238, 0.9); padding: 10px 15px; border-radius: 10px; margin-bottom: 10px; font-size: 16px; text-align: center; }
        .alert-success { color: #2e7d32; background: rgba(232, 245, 233, 0.9); padding: 10px 15px; border-radius: 10px; margin-bottom: 10px; font-size: 16px; text-align: center; }

        details summary { cursor: pointer; font-weight: bold; color: #2e7d32; padding: 5px 0; user-select: none; }
        details summary:hover { color: #1b5e20; }
    </style>

<style id="salaryTool3DModeStyles">
/* Unified EXTRA-DEPTH 3D theme shared by Cargo and CMHK; disabled in normal 2D mode. */
body.salary-3d-mode { perspective: 1800px; perspective-origin: 50% 0; }
body.salary-3d-mode .container,
body.salary-3d-mode main,
body.salary-3d-mode .content,
body.salary-3d-mode .main-content { transform-style: preserve-3d; }
body.salary-3d-mode .card,
body.salary-3d-mode .calendar-card,
body.salary-3d-mode .liquid-glass-card,
body.salary-3d-mode .chart-container,
body.salary-3d-mode .chart-card,
body.salary-3d-mode .report-card,
body.salary-3d-mode .table-responsive {
  position: relative;
  transform: translateZ(10px) rotateX(1.8deg);
  transform-style: preserve-3d;
  border: 2px solid rgba(255,255,255,.42) !important;
  border-radius: 16px !important;
  box-shadow: 0 8px 0 rgba(0,0,0,.18), 0 18px 28px rgba(0,0,0,.30), 0 32px 48px rgba(0,0,0,.18), inset 0 2px 0 rgba(255,255,255,.42), inset 0 -3px 0 rgba(0,0,0,.12) !important;
  transition: transform .28s cubic-bezier(.2,.8,.2,1), box-shadow .28s ease, filter .28s ease;
  filter: drop-shadow(0 8px 8px rgba(0,0,0,.13));
}
body.salary-3d-mode .card:hover:not(.sticky-control-panel),
body.salary-3d-mode .calendar-card:hover,
body.salary-3d-mode .chart-container:hover,
body.salary-3d-mode .chart-card:hover,
body.salary-3d-mode .report-card:hover {
  transform: translateY(-9px) translateZ(24px) rotateX(4deg) rotateY(-1deg) scale(1.008);
  box-shadow: 0 12px 0 rgba(0,0,0,.20), 0 26px 38px rgba(0,0,0,.34), 0 44px 58px rgba(0,0,0,.20), inset 0 2px 0 rgba(255,255,255,.5) !important;
  filter: saturate(1.12) brightness(1.035);
}
body.salary-3d-mode .card-header,
body.salary-3d-mode .card-title,
body.salary-3d-mode .section-title { text-shadow: 0 2px 0 rgba(0,0,0,.18), 0 5px 8px rgba(0,0,0,.20); }
body.salary-3d-mode table { transform: translateZ(8px); transform-style: preserve-3d; border-collapse: separate !important; border-spacing: 3px !important; }
body.salary-3d-mode th {
  transform: translateZ(7px);
  border-radius: 9px;
  box-shadow: 0 4px 0 rgba(0,0,0,.20), 0 7px 10px rgba(0,0,0,.16), inset 0 2px 0 rgba(255,255,255,.38) !important;
}
body.salary-3d-mode td {
  transform: translateZ(3px);
  border-radius: 8px;
  box-shadow: 0 4px 0 rgba(0,0,0,.18), 0 7px 10px rgba(0,0,0,.14), inset 0 2px 0 rgba(255,255,255,.24) !important;
  transition: transform .2s ease, box-shadow .2s ease, filter .2s ease;
}
body.salary-3d-mode td:hover {
  transform: translateY(-4px) translateZ(16px) rotateX(3deg) scale(1.025);
  box-shadow: 0 7px 0 rgba(0,0,0,.22), 0 12px 16px rgba(0,0,0,.22), inset 0 2px 0 rgba(255,255,255,.4) !important;
  filter: brightness(1.08);
}
body.salary-3d-mode canvas,
body.salary-3d-mode svg,
body.salary-3d-mode .chart-dark,
body.salary-3d-mode .chart-light,
body.salary-3d-mode [id*="chart" i],
body.salary-3d-mode [class*="chart" i] {
  transform: translateZ(12px);
  transform-style: preserve-3d;
  filter: drop-shadow(0 8px 5px rgba(0,0,0,.26)) drop-shadow(0 16px 12px rgba(0,0,0,.12));
}
body.salary-3d-mode .btn,
body.salary-3d-mode button,
body.salary-3d-mode .form-control,
body.salary-3d-mode .form-select {
  transform: translateZ(5px);
  border: 1px solid rgba(255,255,255,.25);
  box-shadow: 0 5px 0 rgba(0,0,0,.22), 0 9px 14px rgba(0,0,0,.18), inset 0 1px 0 rgba(255,255,255,.30) !important;
  transition: transform .18s ease, box-shadow .18s ease;
}
body.salary-3d-mode .btn:active,
body.salary-3d-mode button:active { transform: translateY(3px) translateZ(2px); box-shadow: 0 2px 0 rgba(0,0,0,.22), 0 4px 7px rgba(0,0,0,.16) !important; }
@media (prefers-reduced-motion: reduce) {
  body.salary-3d-mode *, body.salary-3d-mode *::before, body.salary-3d-mode *::after { transition-duration: .01ms !important; animation-duration: .01ms !important; }
}
/* 3D mode: keep large Cargo/CMHK panels seated in the page instead of
   lifting the whole panel above nearby controls. Retain their glass frame,
   subtle depth and all other 3D effects (buttons, charts and table cells). */
body.salary-3d-mode .card,
body.salary-3d-mode .calendar-card,
body.salary-3d-mode .liquid-glass-card,
body.salary-3d-mode .chart-container,
body.salary-3d-mode .chart-card,
body.salary-3d-mode .report-card,
body.salary-3d-mode .table-responsive {
  transform: none !important;
  transform-style: flat !important;
  box-shadow: 0 2px 8px rgba(0,0,0,.16), 0 5px 14px rgba(0,0,0,.10),
              inset 0 1px 0 rgba(255,255,255,.32) !important;
  filter: none !important;
}
body.salary-3d-mode .card:hover:not(.sticky-control-panel),
body.salary-3d-mode .calendar-card:hover,
body.salary-3d-mode .chart-container:hover,
body.salary-3d-mode .chart-card:hover,
body.salary-3d-mode .report-card:hover {
  transform: none !important;
  box-shadow: 0 3px 10px rgba(0,0,0,.18), 0 6px 16px rgba(0,0,0,.12),
              inset 0 1px 0 rgba(255,255,255,.36) !important;
  filter: brightness(1.02) !important;
}
</style>
<script>
(function(){
  function set3DMode(enabled){ document.body.classList.toggle('salary-3d-mode', !!enabled); }
  window.addEventListener('message', function(event){
    if(event.data && event.data.type === 'salary-tool-3d-mode') set3DMode(event.data.enabled);
  });
  try { set3DMode(localStorage.getItem('salary_tool_3d_mode') === 'on'); } catch(e) {}
  window.parent.postMessage({type:'salary-tool-3d-mode-ready'}, '*');
})();
</script>
</head>
<body class="{{ session.get('theme', 'dark-mode') }}">

    <details class="top-console" id="consoleDetails">
        <summary style="cursor: pointer; position: relative; font-weight: bold; font-size: 14px; display: flex; justify-content: space-between; align-items: center; outline: none; min-height: 26px; line-height: 26px;">
            <div style="display: flex; align-items: center; gap: 4px;" onclick="event.stopPropagation()"></div>
            <span style="position: absolute; left: 50%; transform: translateX(-50%); display: flex; align-items: center; justify-content: center; gap: 6px; white-space: nowrap;">⚙️ CMHK <span>▼</span></span>
        </summary>
        <div style="margin-top: 6px; padding-top: 6px; border-top: 1px dashed rgba(200, 230, 201, 0.6);">
            <div style="margin-bottom: 6px;">
                <a href="{{ url_for('index', mode='add') }}" style="text-decoration: none; font-weight: bold; font-size: 13px; color: {% if mode == 'add' %}#4caf50{% else %}#7f8c8d{% endif %};">➕ 新增</a> |
                <a href="{{ url_for('index', mode='copy') }}" style="margin-left: 4px; text-decoration: none; font-weight: bold; font-size: 13px; color: {% if mode == 'copy' %}#4caf50{% else %}#7f8c8d{% endif %};">📋 複製</a> |
                <a href="{{ url_for('index', mode='edit') }}" style="margin-left: 4px; text-decoration: none; font-weight: bold; font-size: 13px; color: {% if mode == 'edit' %}#4caf50{% else %}#7f8c8d{% endif %};">✏️ 修改/刪除</a> |
            </div>

            {% if flash_error %}
            <div style="background: rgba(244, 67, 54, 0.2); border: 1px solid #f44336; color: #f44336; padding: 4px 8px; border-radius: 6px; font-size: 12px; margin: 4px 0; font-weight: bold;">
                {{ flash_error }}
            </div>
            {% endif %}

            {% if flash_success %}
            <div style="background: rgba(76, 175, 80, 0.2); border: 1px solid #4caf50; color: #4caf50; padding: 4px 8px; border-radius: 6px; font-size: 12px; margin: 4px 0; font-weight: bold;">
                {{ flash_success }}
            </div>
            {% endif %}

            <div>
                {% if mode == 'add' %}
                <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
                <form method="POST" action="{{ url_for('add_record') }}" style="display: flex; gap: 6px; align-items: center; flex-wrap: wrap; margin: 0;">
                    <span style="font-size:12px;">日期:</span><input type="date" name="date" value="{{ selected_date or today_str }}" required>
                    <span style="font-size:12px;">備註:</span><span><select name="note_type" id="add_note_type" onchange="toggleCmhkCustomNote('add')" style="width:85px;"><option value="null" selected>null</option><option value="Paid">Paid</option><option value="Off">Off</option><option value="other">other</option></select><input type="text" name="custom_note" id="add_custom_note" placeholder="自訂備註" style="display:none;width:100px;margin-top:3px;"></span>
                    <span style="font-size:12px;">工時:</span><input type="number" name="hours" value="8" style="width: 50px;" min="0">
                    <span style="font-size:12px;">時薪:</span><input type="number" name="rate" value="125" style="width: 50px;" min="0">
                    <button type="submit" class="btn-success" style="padding:4px 8px;">確認新增</button>
                </form>
                </div>
                {% elif mode == 'copy' %}
                {% if data %}
                <form method="POST" action="{{ url_for('copy_record') }}" style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin:0;">
                    <span style="font-size:12px;">📋 複製原有日期:</span><select name="source_date" required style="max-width:190px;">{% for item in data|sort(attribute='date', reverse=True) %}<option value="{{ item.date }}" {% if selected_date == item.date %}selected{% endif %}>{{ item.date }}｜{{ item.hours }}h｜{{ item.note }}</option>{% endfor %}</select>
                    <span style="font-size:12px;">開始日期:</span><input type="date" name="target_start" value="{{ today_str }}" required>
                    <span style="font-size:12px;">結束日期（單日可留空）:</span><input type="date" name="target_end">
                    <button type="submit" class="btn-success" style="padding:4px 8px;">📋 批量複製</button>
                </form>
                {% else %}
                <span style="font-size:12px;">暫無記錄可供複製。</span>
                {% endif %}
                {% else %}
                <div style="display: flex; gap: 6px; align-items: center; flex-wrap: wrap; margin: 0;">
                    <form method="GET" action="{{ url_for('index') }}" style="margin:0; display:flex; align-items:center; gap:4px;">
                        <input type="hidden" name="mode" value="edit">
                        <span style="font-size:12px;">紀錄:</span>
                        <select name="edit_idx" onchange="this.form.submit()" style="max-width: 200px;">
                            {% for idx, item in sorted_indexed_data %}
                                <option value="{{ idx }}" {% if selected_edit_idx == idx %}selected{% endif %}>
                                    [{{ item.date }}] {{ item.hours }}h ({{ item.note }})
                                </option>
                            {% endfor %}
                        </select>
                    </form>
                    {% if edit_item %}
                    <form method="POST" action="{{ url_for('update_record') }}" style="margin:0; display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                        <input type="hidden" name="idx" value="{{ selected_edit_idx }}">
                        <input type="date" name="date" value="{{ edit_item.date }}" required style="width: 110px;">
                        {% set edit_note_lower = edit_item.note|string|lower %}
                        <span><select name="note_type" id="edit_note_type" onchange="toggleCmhkCustomNote('edit')" style="width:85px;">{% if edit_note_lower == 'null' %}<option value="null" selected>null</option>{% else %}<option value="null">null</option>{% endif %}{% if edit_note_lower == 'paid' %}<option value="Paid" selected>Paid</option>{% else %}<option value="Paid">Paid</option>{% endif %}{% if edit_note_lower == 'off' %}<option value="Off" selected>Off</option>{% else %}<option value="Off">Off</option>{% endif %}<option value="other" {% if edit_note_lower not in ['null','paid','off'] %}selected{% endif %}>other</option></select><input type="text" name="custom_note" id="edit_custom_note" value="{% if edit_note_lower not in ['null','paid','off'] %}{{ edit_item.note }}{% endif %}" placeholder="自訂備註" style="{% if edit_note_lower in ['null','paid','off'] %}display:none;{% endif %}width:100px;margin-top:3px;"></span>
                        <input type="number" name="hours" value="{{ edit_item.hours }}" style="width: 45px;" min="0">
                        <input type="number" name="rate" value="{{ edit_item.rate }}" style="width: 45px;" min="0">
                        <button type="submit" name="action" value="update" class="btn-success" style="padding:3px 6px;">更新</button>
                        <button type="submit" name="action" value="delete" class="btn-danger" style="padding:3px 6px;">刪除</button>
                    </form>
                    {% endif %}
                </div>
                {% endif %}
            </div>
        </div>
    </details>

    <div id="cmhkDayPopup" role="dialog" aria-label="日曆日期操作">
        <button type="button" data-action="add" title="新增" aria-label="新增">➕</button>
        <button type="button" data-action="copy" title="複製" aria-label="複製">📋</button>
        <button type="button" data-action="edit" title="修改" aria-label="修改">✏️</button>
        <button type="button" data-action="delete" title="刪除" aria-label="刪除">🗑️</button>
    </div>

    <script>
        (function handleCalendarDeleteRequest() {
            const params = new URLSearchParams(window.location.search);
            if (params.get('confirm_delete') !== '1') return;
            const form = document.querySelector('form[action*="/update"]');
            if (!form) return;
            const action = document.createElement('input');
            action.type = 'hidden';
            action.name = 'action';
            action.value = 'delete';
            form.appendChild(action);
            form.requestSubmit();
        })();

        function toggleCmhkCustomNote(which) {
            var select = document.getElementById(which === 'add' ? 'add_note_type' : 'edit_note_type');
            var input = document.getElementById(which === 'add' ? 'add_custom_note' : 'edit_custom_note');
            if (!select || !input) return;
            input.style.display = select.value === 'other' ? '' : 'none';
            input.required = select.value === 'other';
        }
        (function initCalendarDayActions() {
            const popup = document.getElementById('cmhkDayPopup');
            if (!popup) return;
            let activeCell = null;
            function hidePopup() { popup.style.display = 'none'; activeCell = null; }
            function showPopup(cell) {
                activeCell = cell;
                const recordIndex = (cell.getAttribute('data-record-index') || '').trim();
                const hasRecord = recordIndex !== '' && recordIndex !== 'None' && recordIndex !== 'null';
                // Force visibility with inline display so existing button CSS
                // cannot override the no-record state.
                popup.querySelectorAll('button[data-action]').forEach(button => {
                    const shouldShow = button.dataset.action === 'add' || hasRecord;
                    button.hidden = !shouldShow;
                    button.style.display = shouldShow ? 'inline-flex' : 'none';
                });
                const rect = cell.getBoundingClientRect();
                popup.style.display = 'flex';
                const width = popup.offsetWidth || 190, height = popup.offsetHeight || 55;
                const left = Math.min(Math.max(8, rect.left + rect.width / 2 - width / 2), window.innerWidth - width - 8);
                let top = rect.bottom + 6;
                if (top + height > window.innerHeight - 8) top = rect.top - height - 6;
                popup.style.left = left + 'px'; popup.style.top = Math.max(8, top) + 'px';
            }
            // Event delegation is required because this script is rendered
            // before the calendar tables later in the HTML document.
            document.addEventListener('click', event => {
                const cell = event.target.closest('.calendar-day-cell');
                if (cell) {
                    event.stopPropagation();
                    showPopup(cell);
                }
            });
            document.addEventListener('keydown', event => {
                const cell = event.target.closest('.calendar-day-cell');
                if (cell && (event.key === 'Enter' || event.key === ' ')) {
                    event.preventDefault();
                    showPopup(cell);
                }
            });
            popup.addEventListener('click', event => {
                const button = event.target.closest('button[data-action]');
                if (!button || !activeCell) return;
                const action = button.dataset.action, date = activeCell.dataset.date;
                const recordIndex = activeCell.dataset.recordIndex;
                const params = new URLSearchParams();
                if (action === 'add') { params.set('mode','add'); params.set('selected_date',date); }
                else if (action === 'copy') { params.set('mode','copy'); params.set('selected_date',date); }
                else {
                    if (!recordIndex) { hidePopup(); return; }
                    params.set('mode','edit'); params.set('edit_idx',recordIndex); params.set('selected_date',date);
                    if (action === 'delete') params.set('confirm_delete','1');
                }
                window.location.href = window.location.pathname + '?' + params.toString();
            });
            document.addEventListener('click', event => {
                if (!popup.contains(event.target) && !event.target.closest('.calendar-day-cell')) hidePopup();
            });
            window.addEventListener('resize',hidePopup); window.addEventListener('scroll',hidePopup,true);
        })();

        const consoleDetails = document.getElementById('consoleDetails');
        if (consoleDetails) {
            const savedState = localStorage.getItem('cmhk_console_open');
            const hasCalendarAction = new URLSearchParams(window.location.search).has('selected_date');
            if (hasCalendarAction || savedState === 'true') {
                consoleDetails.open = true;
            } else {
                consoleDetails.open = false;
            }

            function updatePadding() {
                if (consoleDetails.open) {
                    document.body.style.paddingTop = (consoleDetails.offsetHeight + 10) + 'px';
                } else {
                    document.body.style.paddingTop = '38px';
                }
            }

            updatePadding();
            window.addEventListener('resize', updatePadding);

            consoleDetails.addEventListener('toggle', function() {
                localStorage.setItem('cmhk_console_open', consoleDetails.open);
                updatePadding();
            });
        }
    </script>

    <div class="container">
        <div class="grid-2">
            <div>
                <div class="card calendar-card">
                    <div class="calendar-heading"><span style="display:inline-flex;align-items:center;gap:8px;"><span class="cmhk-home-title">📶CMHK<span class="cmhk-title-mobile-hide">📡</span></span>{% set speech_month = recent_calendar_months.get(current_ym, {}) %}<button type="button" id="cmhkMonthlySpeechBtn" style="width:38px;height:38px;font-size:1.15rem;padding:0;border-radius:50%;cursor:pointer;" title="讀出本月薪資摘要" aria-label="讀出本月薪資摘要" data-hours="{{ speech_month.get('hours', 0) }}" data-month-pay="{{ speech_month.get('total_pay', 0) }}" data-total="{{ total_sum }}">🔊</button><button type="button" id="cmhkMonthlySpeechStopBtn" hidden style="width:38px;height:38px;font-size:1.15rem;padding:0;border-radius:50%;cursor:pointer;background:#fff1f1;color:#b91c1c;border:1px solid #ef4444;" title="停止朗讀" aria-label="停止朗讀">⏹️</button></span><span class="income-total">總金💰 <span style="font-family:monospace;">$ {{ "{:,}".format(total_sum) }} HKD</span></span></div>
                    {% for m, info in recent_calendar_months.items() %}
                        <h4 style="margin: 10px 0 5px 0; font-size: 14px;">📅 【{{ m }}】總工時：<code>{{ info.hours }}h</code> | 月收入：<code>$ {{ info.total_pay }}</code></h4>
                        <div style="overflow-x: auto;">
                            <table>
                                <tr><th>日</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th></tr>
                                {% for week in info.cal %}
                                    <tr>
                                        {% for d in week %}
                                            {% if d == 0 %}
                                                <td style="background:transparent;"></td>
                                            {% else %}
                                                {% set day_info = info.days.get(d) %}
                                                {% set is_today = (m == current_ym and d == current_day) %}
                                                <td class="calendar-day-cell" data-date="{{ '%s-%02d'|format(m, d) }}"
                                                    data-record-index="{{ day_info.index if day_info else '' }}" tabindex="0" role="button"
                                                    aria-label="{{ m }}-{{ '%02d'|format(d) }} 日記錄操作" style="
                                                    {% if day_info and day_info.note == 'paid' %}background-color: #FFD54F; color: #000000;
                                                    {% elif day_info and day_info.pay > 0 %}background-color: #388E3C; color: #ffffff;
                                                    {% elif day_info and day_info.note == 'off' %}background-color: #4A2E2B; color: #E57373;{% endif %}
                                                    {% if is_today %}
                                                        {% if session.get('theme') == 'light-mode' %}
                                                            border: 2px solid #0056b3 !important; background-color: {% if day_info and day_info.note == 'paid' %}#FFD54F{% elif day_info and day_info.pay > 0 %}#2e7d32{% elif day_info and day_info.note == 'off' %}#5c3835{% else %}#cfe2ff{% endif %} !important;
                                                        {% else %}
                                                            border: 2px solid #38ef7d !important; box-shadow: inset 0 0 8px rgba(56, 239, 125, 0.6);
                                                        {% endif %}
                                                    {% endif %}
                                                ">
                                                    <b style="{% if day_info and day_info.note == 'paid' %}color: {% if day_info.hours == 0 %}#ff0000{% else %}#0000ff{% endif %}; font-size: 15px;{% elif is_today %}{% if session.get('theme') == 'light-mode' %}color: #004085; font-size: 15px;{% else %}color: #38ef7d; font-size: 15px; text-shadow: 0 0 5px rgba(56, 239, 125, 0.8);{% endif %}{% endif %}">{{ d }}</b>
                                                    {% if day_info and day_info.note == 'paid' %}
                                                        <br><span style="color:#000000; font-size:11px;">Paid</span>
                                                    {% elif day_info and day_info.pay > 0 %}
                                                        <br><span class="calendar-pay" style="color:#ffffff;">${{ day_info.pay }}</span>
                                                    {% elif day_info and day_info.note == 'off' %}
                                                        <br><span style="color:#E57373; font-size:11px;">off</span>
                                                    {% endif %}
                                                </td>
                                            {% endif %}
                                        {% endfor %}
                                    </tr>
                                {% endfor %}
                            </table>
                        </div>
                    {% endfor %}

                    {% if older_calendar_months %}
                        <details style="margin-top: 15px; border-top: 1px dashed rgba(200, 230, 201, 0.4); padding-top: 10px;">
                            <summary style="font-size: 14px;">📂 展開所有以往月份的日曆 (共 {{ older_calendar_months|length }} 個月份)</summary>
                            <div style="margin-top: 10px;">
                                {% for m, info in older_calendar_months.items() %}
                                    <h4 style="margin: 15px 0 5px 0; font-size: 14px;">📅 【{{ m }}】總工時：<code>{{ info.hours }}h</code> | 月收入：<code>$ {{ info.total_pay }}</code></h4>
                                    <div style="overflow-x: auto;">
                                        <table>
                                            <tr><th>日</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th></tr>
                                            {% for week in info.cal %}
                                                <tr>
                                                    {% for d in week %}
                                                        {% if d == 0 %}
                                                            <td style="background:transparent;"></td>
                                                        {% else %}
                                                            {% set day_info = info.days.get(d) %}
                                                            {% set is_today = (m == current_ym and d == current_day) %}
                                                            <td class="calendar-day-cell" data-date="{{ '%s-%02d'|format(m, d) }}"
                                                                data-record-index="{{ day_info.index if day_info else '' }}" tabindex="0" role="button"
                                                                aria-label="{{ m }}-{{ '%02d'|format(d) }} 日記錄操作" style="
                                                                {% if day_info and day_info.note == 'paid' %}background-color: #FFD54F; color: #000000;
                                                                {% elif day_info and day_info.pay > 0 %}background-color: #388E3C; color: #ffffff;
                                                                {% elif day_info and day_info.note == 'off' %}background-color: #4A2E2B; color: #E57373;{% endif %}
                                                                {% if is_today %}
                                                                    {% if session.get('theme') == 'light-mode' %}
                                                                        border: 2px solid #0056b3 !important; background-color: {% if day_info and day_info.note == 'paid' %}#FFD54F{% elif day_info and day_info.pay > 0 %}#2e7d32{% elif day_info and day_info.note == 'off' %}#5c3835{% else %}#cfe2ff{% endif %} !important;
                                                                    {% else %}
                                                                        border: 2px solid #38ef7d !important; box-shadow: inset 0 0 8px rgba(56, 239, 125, 0.6);
                                                                    {% endif %}
                                                                {% endif %}
                                                            ">
                                                                <b style="{% if day_info and day_info.note == 'paid' %}color: {% if day_info.hours == 0 %}#ff0000{% else %}#0000ff{% endif %}; font-size: 15px;{% elif is_today %}{% if session.get('theme') == 'light-mode' %}color: #004085; font-size: 15px;{% else %}color: #38ef7d; font-size: 15px; text-shadow: 0 0 5px rgba(56, 239, 125, 0.8);{% endif %}{% endif %}">{{ d }}</b>
                                                                {% if day_info and day_info.note == 'paid' %}
                                                                    <br><span style="color:#000000; font-size:11px;">Paid</span>
                                                                {% elif day_info and day_info.pay > 0 %}
                                                                    <br><span class="calendar-pay" style="color:#ffffff;">${{ day_info.pay }}</span>
                                                                {% elif day_info and day_info.note == 'off' %}
                                                                    <br><span style="color:#E57373; font-size:11px;">off</span>
                                                                {% endif %}
                                                            </td>
                                                        {% endif %}
                                                    {% endfor %}
                                                </tr>
                                            {% endfor %}
                                        </table>
                                    </div>
                                {% endfor %}
                            </div>
                        </details>
                    {% endif %}
                </div>

                <div class="card">
                    <h3 style="margin-top:0; font-size: 16px;">📊 本期出糧週期明細</h3>
                    {% if not data %}
                        <p style="font-size: 13px;">💡 提示：請先在頂部控制台新增第一筆資料！</p>
                    {% else %}
                        <div class="grid-3" style="margin-bottom: 10px;">
                            <div class="stat-box">
                                <span style="font-size: 12px; display: block; opacity: 0.8;">📆 今期（本月）</span>
                                <span style="font-size: 14px; font-weight: bold;">{{ start.strftime('%Y/%m/%d') }} ~ {{ end.strftime('%m/%d') }}</span>
                            </div>
                            <div class="stat-box">
                                <span style="font-size: 12px; display: block; opacity: 0.8;">⏱️ 今期總工時</span>
                                <span style="font-size: 14px; font-weight: bold;">{{ total_h|int }} 小時 ({{ work_days }} 天)</span>
                            </div>
                            <div class="stat-box">
                                <span style="font-size: 12px; display: block; opacity: 0.8;">💰 今期薪水</span>
                                <span style="font-size: 14px; font-weight: bold;">$ {{ total_t|int }}</span>
                            </div>
                        </div>
                        <div class="grid-3">
                            <div class="stat-box">
                                <span style="font-size: 12px; display: block; opacity: 0.8;">📆 上一期（上月）</span>
                                <span style="font-size: 14px; font-weight: bold;">{{ previous_start.strftime('%Y/%m/%d') }} ~ {{ previous_end.strftime('%m/%d') }}</span>
                            </div>
                            <div class="stat-box">
                                <span style="font-size: 12px; display: block; opacity: 0.8;">⏱️ 上一期總工時</span>
                                <span style="font-size: 14px; font-weight: bold;">{{ previous_total_h|int }} 小時 ({{ previous_work_days }} 天)</span>
                            </div>
                            <div class="stat-box">
                                <span style="font-size: 12px; display: block; opacity: 0.8;">💰 上一期薪水</span>
                                <span style="font-size: 14px; font-weight: bold;">$ {{ previous_total_t|int }}</span>
                            </div>
                        </div>
                    {% endif %}
                    <div style="margin-top:12px; padding-top:10px; border-top:1px solid rgba(200,230,201,0.45);">
                        <details>
                            <summary>📥 點擊展開：匯出薪資單 (.xlsx / .pdf / .jpg)</summary>
                            <div style="margin-top: 10px;">
                                {% if data %}
                                    <div style="display:flex; flex-direction:column; gap:8px; width:100%;">
                                        <div style="display:flex; gap:8px; flex-wrap:wrap;">
                                            <a href="{{ url_for('download', file_type='excel', period='current') }}" class="btn-success" style="padding:6px 10px; text-decoration:none; border-radius:4px; font-size:13px;">🟢 Excel｜匯出今期</a>
                                            <a href="{{ url_for('download', file_type='excel', period='previous') }}" class="btn-success" style="padding:6px 10px; text-decoration:none; border-radius:4px; font-size:13px; opacity:.8;">🟢 Excel｜匯出上一期</a>
                                        </div>
                                        <div style="display:flex; gap:8px; flex-wrap:wrap;">
                                            <a href="{{ url_for('download', file_type='pdf', period='current') }}" class="btn-danger" style="padding:6px 10px; text-decoration:none; border-radius:4px; font-size:13px;">🔴 PDF｜匯出今期</a>
                                            <a href="{{ url_for('download', file_type='pdf', period='previous') }}" class="btn-danger" style="padding:6px 10px; text-decoration:none; border-radius:4px; font-size:13px; opacity:.8;">🔴 PDF｜匯出上一期</a>
                                        </div>
                                        <div style="display:flex; gap:8px; flex-wrap:wrap;">
                                            <a href="{{ url_for('download', file_type='jpg', period='current') }}" style="background:#81c784; color:white; padding:6px 10px; text-decoration:none; border-radius:4px; font-size:13px;" class="btn-info">🔵 JPG｜匯出今期</a>
                                            <a href="{{ url_for('download', file_type='jpg', period='previous') }}" style="background:#81c784; color:white; padding:6px 10px; text-decoration:none; border-radius:4px; font-size:13px; opacity:.8;" class="btn-info">🔵 JPG｜匯出上一期</a>
                                        </div>
                                    </div>
                                {% else %}
                                    <p style="opacity:0.6; font-size:13px; margin:5px 0 0 0;">目前無資料可供匯出。</p>
                                {% endif %}
                            </div>
                        </details>
                    </div>
                </div>

                <div class="card">
                    <h3 style="margin-top:0; font-size: 16px;">📋 每日薪資明細表</h3>
                    <div style="overflow-x: auto;">
                        <table>
                            <tr><th>日期</th><th>星期</th><th>工時</th><th>當日薪資</th><th>備註</th></tr>
                            {% for row in table_data %}
                                <tr>
                                    <td>{{ row.date }}</td>
                                    <td>{{ row.weekday }}</td>
                                    <td>{{ row.hours }}</td>
                                    <td>{{ row.pay }}</td>
                                    <td>{{ row.note }}</td>
                                </tr>
                            {% else %}
                                <tr><td colspan="5" style="opacity: 0.6;">本期尚無工作薪資記錄。</td></tr>
                            {% endfor %}
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <br><br>
    <p style="opacity: 0.7; margin: 4px 0 0 0; font-size: 12px;">Copyright &copy; 2026 คาร์ลอส. All Rights Reserved. | 系統時間: {{ today_str }}</p>
<script>
(function () {
  const btn = document.getElementById('cmhkMonthlySpeechBtn');
  const stopBtn = document.getElementById('cmhkMonthlySpeechStopBtn');
  if (!btn || !stopBtn) return;
  function setSpeaking(isSpeaking) { stopBtn.hidden = !isSpeaking; }
  stopBtn.addEventListener('click', function () {
    try { window.speechSynthesis.cancel(); } catch (_) {}
    setSpeaking(false);
  });
  if (!('speechSynthesis' in window) || !('SpeechSynthesisUtterance' in window)) {
    btn.addEventListener('click', function () { alert('此瀏覽器不支援人聲朗讀。'); });
    return;
  }
  const synth = window.speechSynthesis;
  let voices = [];
  function loadVoices() { try { voices = synth.getVoices() || []; } catch (_) { voices = []; } }
  loadVoices();
  if (typeof synth.addEventListener === 'function') synth.addEventListener('voiceschanged', loadVoices);
  else if ('onvoiceschanged' in synth) synth.onvoiceschanged = loadVoices;
  btn.addEventListener('click', function () {
    let lang = String(document.documentElement.lang || 'zh-Hant').toLowerCase();
    // Fetch the shared language setting synchronously from the already-applied page language when possible.
    lang = lang.startsWith('th') ? 'th' : (lang === 'en' || lang.startsWith('en-')) ? 'en' : 'zh-Hant';
    const hours = Number(btn.dataset.hours || 0), monthPay = Number(btn.dataset.monthPay || 0), total = Number(btn.dataset.total || 0);
    const locale = lang === 'th' ? 'th-TH' : lang === 'en' ? 'en-HK' : 'zh-HK';
    const fmt = n => new Intl.NumberFormat(locale, {maximumFractionDigits: 2}).format(n);
    let text, voiceLang, voicePattern;
    if (lang === 'en') {
      text = `This month's total working hours are ${fmt(hours)}. Monthly income is ${fmt(monthPay)} Hong Kong dollars. Total earnings are ${fmt(total)} Hong Kong dollars. Keep up the good work!`;
      voiceLang = 'en-HK'; voicePattern = /^en(-|$)/i;
    } else if (lang === 'th') {
      text = `ชั่วโมงทำงานรวมของเดือนนี้คือ ${fmt(hours)} ชั่วโมง รายได้ต่อเดือนคือ ${fmt(monthPay)} ดอลลาร์ฮ่องกง ยอดเงินรวมคือ ${fmt(total)} ดอลลาร์ฮ่องกง สู้ ๆ ตั้งใจทำงานต่อไป!`;
      voiceLang = 'th-TH'; voicePattern = /^th(-|$)/i;
    } else {
      text = `今個月總工時係 ${fmt(hours)} 小時，人工係 ${fmt(monthPay)} 港元，總金額係 ${fmt(total)} 港元，加油努力！禮拜六仲有雞翼`;
      voiceLang = 'zh-HK'; voicePattern = /^(zh-hk|yue|zh-mo)(-|$)/i;
    }
    loadVoices();
    const norm = v => String(v || '').toLowerCase().replace(/_/g, '-');
    let voice = voices.find(v => norm(v.lang) === norm(voiceLang)) || voices.find(v => voicePattern.test(norm(v.lang)));
    if (!voice && lang === 'zh-Hant') voice = voices.find(v => /^zh(-|$)/i.test(norm(v.lang)));
    try {
      synth.cancel();
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.lang = voice ? (voice.lang || voiceLang) : voiceLang;
      if (voice) utterance.voice = voice;
      utterance.rate = 0.9;
      utterance.onstart = function () { setSpeaking(true); };
      utterance.onend = function () { setSpeaking(false); };
      utterance.onerror = function (event) {
        setSpeaking(false);
        if (event && event.error !== 'canceled' && event.error !== 'interrupted') alert('朗讀未能開始（' + (event && event.error ? event.error : 'unknown') + '）。請檢查裝置語音及音量設定。');
      };
      setSpeaking(true);
      synth.speak(utterance);
      if (typeof synth.resume === 'function') synth.resume();
    } catch (_) { alert('無法啟動朗讀，請確認裝置有可用語音。'); }
  });
})();
</script>
<script>
(async function(){
 const maps={
 en:{'新增':'Add','修改':'Edit','刪除':'Delete','複製':'Copy','儲存':'Save','取消':'Cancel','搜尋':'Search','清除':'Clear','完成':'Completed','未完成':'Incomplete','Paid':'Paid','休息':'Off','其他':'Other','工時':'Hours','時薪':'Hourly rate','收入':'Income','總收入':'Total income','本月':'This month','上月':'Last month','今日':'Today','日期':'Date','備註':'Notes','選擇月份':'Select month','匯出':'Export','下載':'Download','服務統計':'Service statistics','圖表':'Charts','夜間模式':'Dark mode','日間模式':'Light mode','確定':'Confirm','返回':'Back','複製到日期':'Copy to date','開始日期':'Start date','結束日期':'End date','月份':'Month','年':'Year','月':'Month','日':'Day','時':'Hours','分鐘':'Minutes','新增記錄':'Add record','修改記錄':'Edit record','刪除記錄':'Delete record','選擇日期':'Select date','開始':'Start','結束':'End','上一月':'Previous month','下一月':'Next month','上一期':'Previous period','今期':'Current period','匯出今期':'Export current period','匯出上一期':'Export previous period','目前無資料可供匯出。':'No data available to export.','歷史以來總收入累計金額：':'All-time cumulative income:','系統時間：':'System time:','圖表載入失敗，請稍後再試':'Chart loading failed. Please try again later.','新增成功':'Added successfully','儲存成功':'Saved successfully','確定刪除？':'Are you sure you want to delete this?','備註（可選）':'Notes (optional)','自訂備註':'Custom note','Paid（已出糧）':'Paid','Off（休息）':'Off','其他（自訂）':'Other (custom)','複製記錄':'Copy records','單日':'Single day','日期範圍':'Date range','套用':'Apply','設定':'Settings','收入統計':'Income statistics','工作日':'Work days','總工時':'Total hours','平均工時':'Average hours','總額':'Total','金額':'Amount','操作':'Actions','狀態':'Status','全部':'All','載入中…':'Loading…','載入中...':'Loading...','週一':'Mon','週二':'Tue','週三':'Wed','週四':'Thu','週五':'Fri','週六':'Sat','週日':'Sun','月度工資日曆':'Monthly Salary Calendar','月收入':'Monthly income','總工時':'Total hours','每日工資':'Daily pay','當日薪資':'Daily pay','本期出糧週期明細':'Current pay-period details','本期出糧週期':'Current pay period','本期':'Current period','上一期':'Previous period','今期總工時':'Current period total hours','上一期總工時':'Previous period total hours','點擊展開':'Click to expand','展開所有以往月份的日曆':'Expand all previous monthly calendars','匯出薪資單':'Export payslip','匯出今期':'Export current period','匯出上一期':'Export previous period','半月工資表':'Semi-monthly salary report','工資管理系統':'Salary Management System','輸入自訂備註':'Enter custom note','自訂備註':'Custom note','確認新增':'Confirm add','批量複製':'Copy in bulk','複製原有日期':'Source date to copy','複製開始日期':'Copy start date','複製結束日期':'Copy end date','單日可留空':'Leave blank for a single day','選擇記錄':'Select record','最新在上':'Newest first','暫無記錄可供複製':'No records available to copy','本期尚無工作薪資記錄':'No salary records for this period','提示':'Tip','請先在頂部控制台新增第一筆資料':'Please add the first record in the top control panel','數據統計與圖表分析':'Statistics and chart analysis','查看圖表':'View charts','收起圖表':'Hide charts','正在載入圖表分析':'Loading chart analysis','展開':'Expand','目前無資料可供匯出':'No data available to export','成功新增記錄':'Record added successfully','找不到要複製的原有日期記錄':'Source record to copy was not found','請選擇有效的複製日期範圍':'Please select a valid copy date range','批量複製完成':'Bulk copy completed','記錄已刪除':'Record deleted','成功新增':'Added successfully','新增失敗':'Failed to add','不能重複新增相同日子的記錄':'A record already exists for this date','系統時間':'System time','歷史以來總收入累計金額':'All-time cumulative income','工資神器':'Salary Manager'},
 th:{'新增':'เพิ่ม','修改':'แก้ไข','刪除':'ลบ','複製':'คัดลอก','儲存':'บันทึก','取消':'ยกเลิก','搜尋':'ค้นหา','清除':'ล้าง','完成':'เสร็จแล้ว','未完成':'ยังไม่เสร็จ','Paid':'จ่ายแล้ว','休息':'หยุด','其他':'อื่นๆ','工時':'ชั่วโมงทำงาน','時薪':'ค่าจ้างต่อชั่วโมง','收入':'รายได้','總收入':'รายได้รวม','本月':'เดือนนี้','上月':'เดือนก่อน','今日':'วันนี้','日期':'วันที่','備註':'หมายเหตุ','選擇月份':'เลือกเดือน','匯出':'ส่งออก','下載':'ดาวน์โหลด','服務統計':'สถิติบริการ','圖表':'กราฟ','夜間模式':'โหมดกลางคืน','日間模式':'โหมดกลางวัน','確定':'ยืนยัน','返回':'กลับ','複製到日期':'คัดลอกไปยังวันที่','開始日期':'วันที่เริ่มต้น','結束日期':'วันที่สิ้นสุด','月份':'เดือน','年':'ปี','月':'เดือน','日':'วัน','新增記錄':'เพิ่มรายการ','修改記錄':'แก้ไขรายการ','刪除記錄':'ลบรายการ','選擇日期':'เลือกวันที่','開始':'เริ่มต้น','結束':'สิ้นสุด','上一月':'เดือนก่อน','下一月':'เดือนถัดไป','上一期':'งวดก่อน','今期':'งวดปัจจุบัน','匯出今期':'ส่งออกงวดปัจจุบัน','匯出上一期':'ส่งออกงวดก่อน','目前無資料可供匯出。':'ไม่มีข้อมูลสำหรับส่งออก','歷史以來總收入累計金額：':'รายได้สะสมทั้งหมด:','系統時間：':'เวลาระบบ:','圖表載入失敗，請稍後再試':'โหลดกราฟไม่สำเร็จ กรุณาลองอีกครั้ง','新增成功':'เพิ่มสำเร็จ','儲存成功':'บันทึกสำเร็จ','確定刪除？':'ยืนยันการลบหรือไม่?','備註（可選）':'หมายเหตุ (ไม่บังคับ)','自訂備註':'หมายเหตุที่กำหนดเอง','Paid（已出糧）':'จ่ายแล้ว','Off（休息）':'หยุด','其他（自訂）':'อื่นๆ (กำหนดเอง)','複製記錄':'คัดลอกรายการ','單日':'วันเดียว','日期範圍':'ช่วงวันที่','套用':'ใช้','設定':'ตั้งค่า','收入統計':'สถิติรายได้','工作日':'วันทำงาน','總工時':'ชั่วโมงรวม','平均工時':'ชั่วโมงเฉลี่ย','總額':'ยอดรวม','金額':'จำนวนเงิน','操作':'การดำเนินการ','狀態':'สถานะ','全部':'ทั้งหมด','載入中…':'กำลังโหลด…','載入中...':'กำลังโหลด...','週一':'จ.','週二':'อ.','週三':'พ.','週四':'พฤ.','週五':'ศ.','週六':'ส.','週日':'อา.','月度工資日曆':'ปฏิทินเงินเดือนรายเดือน','月收入':'รายได้ต่อเดือน','總工時':'ชั่วโมงรวม','每日工資':'ค่าจ้างรายวัน','當日薪資':'ค่าจ้างรายวัน','本期出糧週期明細':'รายละเอียดงวดจ่ายเงินปัจจุบัน','本期出糧週期':'งวดจ่ายเงินปัจจุบัน','本期':'งวดปัจจุบัน','上一期':'งวดก่อนหน้า','今期總工時':'ชั่วโมงรวมงวดปัจจุบัน','上一期總工時':'ชั่วโมงรวมงวดก่อนหน้า','點擊展開':'คลิกเพื่อขยาย','展開所有以往月份的日曆':'ขยายปฏิทินเดือนก่อนหน้าทั้งหมด','匯出薪資單':'ส่งออกสลิปเงินเดือน','匯出今期':'ส่งออกงวดปัจจุบัน','匯出上一期':'ส่งออกงวดก่อนหน้า','半月工資表':'รายงานเงินเดือนรายครึ่งเดือน','工資管理系統':'ระบบจัดการเงินเดือน','輸入自訂備註':'ป้อนหมายเหตุที่กำหนดเอง','自訂備註':'หมายเหตุที่กำหนดเอง','確認新增':'ยืนยันการเพิ่ม','批量複製':'คัดลอกหลายรายการ','複製原有日期':'วันที่ต้นทาง','複製開始日期':'วันที่เริ่มคัดลอก','複製結束日期':'วันที่สิ้นสุดการคัดลอก','單日可留空':'เว้นว่างได้หากเป็นวันเดียว','選擇記錄':'เลือกรายการ','最新在上':'รายการล่าสุดอยู่ด้านบน','暫無記錄可供複製':'ไม่มีรายการให้คัดลอก','本期尚無工作薪資記錄':'ไม่มีบันทึกค่าจ้างในงวดนี้','提示':'คำแนะนำ','請先在頂部控制台新增第一筆資料':'กรุณาเพิ่มรายการแรกที่แผงควบคุมด้านบน','數據統計與圖表分析':'สถิติและการวิเคราะห์กราฟ','查看圖表':'ดูกราฟ','收起圖表':'ซ่อนกราฟ','正在載入圖表分析':'กำลังโหลดการวิเคราะห์กราฟ','展開':'ขยาย','目前無資料可供匯出':'ไม่มีข้อมูลสำหรับส่งออก','成功新增記錄':'เพิ่มรายการสำเร็จ','找不到要複製的原有日期記錄':'ไม่พบรายการต้นทางที่จะคัดลอก','請選擇有效的複製日期範圍':'กรุณาเลือกช่วงวันที่คัดลอกที่ถูกต้อง','批量複製完成':'คัดลอกหลายรายการเสร็จแล้ว','記錄已刪除':'ลบรายการแล้ว','成功新增':'เพิ่มสำเร็จ','新增失敗':'เพิ่มไม่สำเร็จ','不能重複新增相同日子的記錄':'มีรายการของวันที่นี้อยู่แล้ว ไม่สามารถเพิ่มซ้ำได้','系統時間':'เวลาระบบ','歷史以來總收入累計金額':'รายได้สะสมทั้งหมด','工資神器':'ตัวจัดการเงินเดือน'}
 };

 Object.assign(maps.en, {'總金💰':'Total 💰'});
 Object.assign(maps.th, {'總金💰':'ยอดรวม 💰'});
 // Expanded common UI vocabulary so labels, mixed text, and dynamically rendered controls are localized.
 Object.assign(maps.en, {'新增資料': 'Add data', '新增工資': 'Add salary', '工作記錄': 'Work records', '工資記錄': 'Salary records', '工資日曆': 'Salary calendar', '出糧日期': 'Payday', '出糧日': 'Payday', '月薪': 'Monthly salary', '日薪': 'Daily pay', '時薪': 'Hourly pay', '工資': 'Salary', '薪金': 'Wages', '金額': 'Amount', '工作時數': 'Working hours', '總時數': 'Total hours', '收入合計': 'Total income', '本月總收入': "This month's total income", '上月總收入': "Last month's total income", '本期收入': 'Current period income', '上期收入': 'Previous period income', '備註': 'Notes', '狀態': 'Status', '已付款': 'Paid', '未付款': 'Unpaid', '休假': 'Day off', '自訂': 'Custom', '其他': 'Other', '選擇': 'Select', '請選擇': 'Please select', '請輸入': 'Please enter', '新增日期': 'Add date', '修改日期': 'Edit date', '刪除記錄': 'Delete record', '確認刪除': 'Confirm deletion', '確定要刪除嗎？': 'Are you sure you want to delete this?', '複製日期': 'Copy date', '複製記錄至其他日期': 'Copy records to another date', '開始日期': 'Start date', '結束日期': 'End date', '單日複製': 'Copy one day', '批量複製': 'Bulk copy', '複製成功': 'Copied successfully', '複製失敗': 'Copy failed', '儲存成功': 'Saved successfully', '儲存失敗': 'Save failed', '更新成功': 'Updated successfully', '刪除成功': 'Deleted successfully', '載入失敗': 'Failed to load', '資料載入中': 'Loading data', '沒有資料': 'No data', '暫無資料': 'No data available', '服務統計': 'Service statistics', '收入統計': 'Income statistics', '工時統計': 'Hours statistics', '每月統計': 'Monthly statistics', '圖表分析': 'Chart analysis', '顯示圖表': 'Show charts', '隱藏圖表': 'Hide charts', '展開全部': 'Expand all', '收起全部': 'Collapse all', '匯出 Excel': 'Export Excel', '匯出 CSV': 'Export CSV', '匯出 PDF': 'Export PDF', '下載圖片': 'Download image', '日間': 'Light', '夜間': 'Dark', '主題': 'Theme', '設定': 'Settings', '返回主頁': 'Back to home', '上一頁': 'Previous', '下一頁': 'Next', '上個月': 'Previous month', '下個月': 'Next month', '星期日': 'Sunday', '星期一': 'Monday', '星期二': 'Tuesday', '星期三': 'Wednesday', '星期四': 'Thursday', '星期五': 'Friday', '星期六': 'Saturday', '總計': 'Total', '平均': 'Average', '累計': 'Cumulative', '歷史累計收入': 'Historical cumulative income', '系統時間': 'System time', '載入圖表失敗，請稍後再試': 'Chart loading failed. Please try again later.'});
 Object.assign(maps.th, {'新增資料': 'เพิ่มข้อมูล', '新增工資': 'เพิ่มค่าจ้าง', '工作記錄': 'บันทึกการทำงาน', '工資記錄': 'บันทึกเงินเดือน', '工資日曆': 'ปฏิทินเงินเดือน', '出糧日期': 'วันจ่ายเงิน', '出糧日': 'วันจ่ายเงิน', '月薪': 'เงินเดือนรายเดือน', '日薪': 'ค่าจ้างรายวัน', '時薪': 'ค่าจ้างต่อชั่วโมง', '工資': 'เงินเดือน', '薪金': 'ค่าจ้าง', '金額': 'จำนวนเงิน', '工作時數': 'ชั่วโมงทำงาน', '總時數': 'ชั่วโมงรวม', '收入合計': 'รายได้รวม', '本月總收入': 'รายได้รวมเดือนนี้', '上月總收入': 'รายได้รวมเดือนก่อน', '本期收入': 'รายได้งวดปัจจุบัน', '上期收入': 'รายได้งวดก่อน', '備註': 'หมายเหตุ', '狀態': 'สถานะ', '已付款': 'จ่ายแล้ว', '未付款': 'ยังไม่จ่าย', '休假': 'วันหยุด', '自訂': 'กำหนดเอง', '其他': 'อื่นๆ', '選擇': 'เลือก', '請選擇': 'กรุณาเลือก', '請輸入': 'กรุณากรอก', '新增日期': 'เพิ่มวันที่', '修改日期': 'แก้ไขวันที่', '刪除記錄': 'ลบบันทึก', '確認刪除': 'ยืนยันการลบ', '確定要刪除嗎？': 'คุณแน่ใจหรือไม่ว่าต้องการลบ?', '複製日期': 'คัดลอกวันที่', '複製記錄至其他日期': 'คัดลอกบันทึกไปยังวันอื่น', '開始日期': 'วันที่เริ่มต้น', '結束日期': 'วันที่สิ้นสุด', '單日複製': 'คัดลอกวันเดียว', '批量複製': 'คัดลอกหลายรายการ', '複製成功': 'คัดลอกสำเร็จ', '複製失敗': 'คัดลอกไม่สำเร็จ', '儲存成功': 'บันทึกสำเร็จ', '儲存失敗': 'บันทึกไม่สำเร็จ', '更新成功': 'อัปเดตสำเร็จ', '刪除成功': 'ลบสำเร็จ', '載入失敗': 'โหลดไม่สำเร็จ', '資料載入中': 'กำลังโหลดข้อมูล', '沒有資料': 'ไม่มีข้อมูล', '暫無資料': 'ไม่มีข้อมูล', '服務統計': 'สถิติบริการ', '收入統計': 'สถิติรายได้', '工時統計': 'สถิติชั่วโมงทำงาน', '每月統計': 'สถิติรายเดือน', '圖表分析': 'วิเคราะห์กราฟ', '顯示圖表': 'แสดงกราฟ', '隱藏圖表': 'ซ่อนกราฟ', '展開全部': 'ขยายทั้งหมด', '收起全部': 'ย่อทั้งหมด', '匯出 Excel': 'ส่งออก Excel', '匯出 CSV': 'ส่งออก CSV', '匯出 PDF': 'ส่งออก PDF', '下載圖片': 'ดาวน์โหลดรูปภาพ', '日間': 'กลางวัน', '夜間': 'กลางคืน', '主題': 'ธีม', '設定': 'ตั้งค่า', '返回主頁': 'กลับหน้าหลัก', '上一頁': 'ก่อนหน้า', '下一頁': 'ถัดไป', '上個月': 'เดือนก่อน', '下個月': 'เดือนถัดไป', '星期日': 'วันอาทิตย์', '星期一': 'วันจันทร์', '星期二': 'วันอังคาร', '星期三': 'วันพุธ', '星期四': 'วันพฤหัสบดี', '星期五': 'วันศุกร์', '星期六': 'วันเสาร์', '總計': 'รวม', '平均': 'เฉลี่ย', '累計': 'สะสม', '歷史累計收入': 'รายได้สะสมย้อนหลัง', '系統時間': 'เวลาระบบ', '載入圖表失敗，請稍後再試': 'โหลดกราฟไม่สำเร็จ กรุณาลองอีกครั้ง'});
 // Targeted fixes for labels previously remaining in Chinese.
 Object.assign(maps.en, {'總金💰': 'Total💰', '總金': 'Total', '更新': 'Update', '更新記錄': 'Update record', '更新成功': 'Updated successfully', '更新失敗': 'Update failed', '修改/刪除': 'Edit / Delete', '月度工資日曆': 'Monthly Salary Calendar', '月度工資日曆 (最近)': 'Monthly Salary Calendar (Recent)', '(最近)': '(Recent)', '最近': 'Recent', '星期日': 'Sunday', '星期天': 'Sunday', '星期一': 'Monday', '星期二': 'Tuesday', '星期三': 'Wednesday', '星期四': 'Thursday', '星期五': 'Friday', '星期六': 'Saturday', '星期': 'Day', '每日薪資明細表': 'Daily Salary Details', '日期': 'Date', '工時': 'Hours', '當日薪資': 'Daily Pay', '備註': 'Notes', '本期出糧週期明細': 'Current Pay-period Details', '出糧週期': 'Pay Period', '週期開始': 'Period Start', '週期結束': 'Period End', '總工時': 'Total Hours', '總薪資': 'Total Pay', '本期': 'Current Period', '上期': 'Previous Period', '新增': 'Add', '儲存': 'Save', '取消': 'Cancel', '確定': 'Confirm', '選擇日期': 'Select Date', '日期範圍': 'Date Range', '匯出': 'Export', '下載': 'Download', '操作': 'Actions', '金額': 'Amount', '總計': 'Total', '工資': 'Salary', '薪資': 'Pay', '目前無資料': 'No data available', '沒有資料': 'No data'});
 Object.assign(maps.th, {'總金💰': 'ยอดรวม💰', '總金': 'ยอดรวม', '更新': 'อัปเดต', '更新記錄': 'อัปเดตรายการ', '更新成功': 'อัปเดตสำเร็จ', '更新失敗': 'อัปเดตไม่สำเร็จ', '修改/刪除': 'แก้ไข / ลบ', '月度工資日曆': 'ปฏิทินเงินเดือนรายเดือน', '月度工資日曆 (最近)': 'ปฏิทินเงินเดือนรายเดือน (ล่าสุด)', '(最近)': '(ล่าสุด)', '最近': 'ล่าสุด', '星期日': 'วันอาทิตย์', '星期天': 'วันอาทิตย์', '星期一': 'วันจันทร์', '星期二': 'วันอังคาร', '星期三': 'วันพุธ', '星期四': 'วันพฤหัสบดี', '星期五': 'วันศุกร์', '星期六': 'วันเสาร์', '星期': 'วัน', '每日薪資明細表': 'รายละเอียดค่าจ้างรายวัน', '日期': 'วันที่', '工時': 'ชั่วโมงทำงาน', '當日薪資': 'ค่าจ้างรายวัน', '備註': 'หมายเหตุ', '本期出糧週期明細': 'รายละเอียดงวดจ่ายเงินปัจจุบัน', '出糧週期': 'งวดจ่ายเงิน', '週期開始': 'เริ่มงวด', '週期結束': 'สิ้นสุดงวด', '總工時': 'ชั่วโมงรวม', '總薪資': 'ค่าจ้างรวม', '本期': 'งวดปัจจุบัน', '上期': 'งวดก่อนหน้า', '新增': 'เพิ่ม', '儲存': 'บันทึก', '取消': 'ยกเลิก', '確定': 'ยืนยัน', '選擇日期': 'เลือกวันที่', '日期範圍': 'ช่วงวันที่', '匯出': 'ส่งออก', '下載': 'ดาวน์โหลด', '操作': 'การดำเนินการ', '金額': 'จำนวนเงิน', '總計': 'รวม', '工資': 'ค่าจ้าง', '薪資': 'ค่าจ้าง', '目前無資料': 'ไม่มีข้อมูล', '沒有資料': 'ไม่มีข้อมูล'});
 let lang='zh-Hant';
 try { const r=await fetch('/get_language',{credentials:'same-origin',cache:'no-store'}); if(r.ok){const d=await r.json();lang=d.language||lang;} } catch(e) { console.warn('Cannot read shared language setting',e); }
 document.documentElement.lang=lang==='en'?'en':lang==='th'?'th':'zh-Hant';
 if(lang==='zh-Hant') return;
 const map=maps[lang]||{};
 function translate(root){
  const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT); const nodes=[]; let n;
  while((n=walker.nextNode())) nodes.push(n);
  const keys=Object.keys(map).sort((a,b)=>b.length-a.length);
  for(const node of nodes){
   let raw=node.nodeValue;
   if(!raw || !/[\u3400-\u9fff]/.test(raw)) continue;
   // Translate the complete dynamic label, including the rendered month count.
   const olderMonths = raw.match(/📂\\s*展開所有以往月份的日曆\\s*\\(共\\s*(\\d+)\\s*個月份\\)/);
   if (olderMonths) {
    const count = olderMonths[1];
    const localized = lang === 'en'
      ? `📂 Expand all previous monthly calendars (${count} months)`
      : `📂 ขยายปฏิทินเดือนก่อนหน้าทั้งหมด (${count} เดือน)`;
    raw = raw.replace(olderMonths[0], localized);
   }
   // CMHK edit/delete record selector label.
   if (lang === 'en') raw = raw.replace(/紀錄:/g, 'Record:');
   else if (lang === 'th') raw = raw.replace(/紀錄:/g, 'รายการ:');
   const trimmed=raw.trim();
   const weekdayShort = lang==='en' ? {'一':'Mon','二':'Tue','三':'Wed','四':'Thu','五':'Fri','六':'Sat','日':'Sun'} : {'一':'จ.','二':'อ.','三':'พ.','四':'พฤ.','五':'ศ.','六':'ส.','日':'อา.'};
   if (weekdayShort[trimmed]) { node.nodeValue=raw.replace(trimmed, weekdayShort[trimmed]); continue; }
   for(const key of keys){ if(key && raw.includes(key)) raw=raw.split(key).join(map[key]); }
   if(raw!==node.nodeValue) node.nodeValue=raw;
  }
  root.querySelectorAll('input,textarea,button,option,label,th,a,summary,[title],[aria-label]').forEach(el=>{
   for(const attr of ['placeholder','title','aria-label','value']){let v=el.getAttribute(attr);if(v){for(const key of keys){if(key&&v.includes(key))v=v.split(key).join(map[key]);}el.setAttribute(attr,v);}}
   if(el.children.length===0){let t=el.textContent;for(const key of keys){if(key&&t.includes(key))t=t.split(key).join(map[key]);}if(t!==el.textContent)el.textContent=t;}
  });
 }

 // Targeted fixes for remaining hard-coded Chinese labels and units.
 const extraEn = {
  '✈️ Cathay Cargo Pay管理系統':'✈️ Cathay Cargo Pay Management System',
  'Cathay Cargo Pay管理系統':'Cathay Cargo Pay Management System',
  '⚙️ Cargo':'⚙️ Cargo',
  '⚙️ CMHK':'⚙️ CMHK','⚙️ CMHK':'⚙️ CMHK','⚙️ CMHK':'⚙️ CMHK',
  '總工時':'Total Hours','小時':'hours','工作日數':'Work Days','天':'days',
  '週一':'Monday','週二':'Tuesday','週三':'Wednesday','週四':'Thursday','週五':'Friday','週六':'Saturday','週日':'Sunday',
  '星期一':'Monday','星期二':'Tuesday','星期三':'Wednesday','星期四':'Thursday','星期五':'Friday','星期六':'Saturday','星期日':'Sunday',
  '更新':'Update','修改/刪除':'Edit / Delete','最近':'Recent','本期出糧週期明細':'Current Pay-period Details',
  '📊 本期出糧週期明細':'📊 Current Pay-period Details','📅 工作日數':'📅 Work Days','⏱️ 總工時':'⏱️ Total Hours'
 };
 const extraTh = {
  '✈️ Cathay Cargo Pay管理系統':'✈️ ระบบจัดการเงินเดือน Cathay Cargo Pay',
  'Cathay Cargo Pay管理系統':'ระบบจัดการเงินเดือน Cathay Cargo Pay',
  '⚙️ Cargo':'⚙️ Cargo','⚙️ CMHK':'⚙️ CMHK','⚙️ CMHK':'⚙️ CMHK','⚙️ CMHK':'⚙️ CMHK',
  '總工時':'ชั่วโมงรวม','小時':'ชั่วโมง','工作日數':'วันทำงาน','天':'วัน',
  '週一':'จันทร์','週二':'อังคาร','週三':'พุธ','週四':'พฤหัสบดี','週五':'ศุกร์','週六':'เสาร์','週日':'อาทิตย์',
  '星期一':'วันจันทร์','星期二':'วันอังคาร','星期三':'วันพุธ','星期四':'วันพฤหัสบดี','星期五':'วันศุกร์','星期六':'วันเสาร์','星期日':'วันอาทิตย์',
  '更新':'อัปเดต','修改/刪除':'แก้ไข / ลบ','最近':'ล่าสุด','本期出糧週期明細':'รายละเอียดงวดจ่ายเงินปัจจุบัน',
  '📊 本期出糧週期明細':'📊 รายละเอียดงวดจ่ายเงินปัจจุบัน','📅 工作日數':'📅 วันทำงาน','⏱️ 總工時':'⏱️ ชั่วโมงรวม'
 };
 Object.assign(maps.en, extraEn);
 Object.assign(maps.th, extraTh);

 translate(document.body);
 // Also translate UI created later by calendar, modal and chart interactions.
 let queued=false; new MutationObserver(()=>{if(queued)return;queued=true;requestAnimationFrame(()=>{queued=false;translate(document.body);});}).observe(document.body,{childList:true,subtree:true,characterData:true});
})();
</script>
<script>
(function(){
  function applyPortalTheme(theme){
    const dark = theme !== 'light';
    document.body.classList.toggle('dark-mode', dark);
    document.body.classList.toggle('light-mode', !dark);
    // 同步保存到 CMHK session，讓重新整理後仍跟隨主頁主題。
    try {
      const body = new URLSearchParams(); body.set('theme', dark ? 'dark-mode' : 'light-mode');
      fetch('set_theme', {method:'POST', credentials:'same-origin', headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'}, body:body.toString(), redirect:'manual'}).catch(()=>{});
    } catch(e) {}
  }
  window.addEventListener('message', function(event){
    if(event.data && event.data.type === 'salary-tool-theme') applyPortalTheme(event.data.theme);
  });
  window.parent.postMessage({type:'salary-tool-theme-ready'}, '*');
})();
</script>
<script>
(function(){
  // Built-in synthesized click sound: no external MP3/WAV file required.
  let audioContext=null;
  let modeEnabled=null;
  function enabled(){
    if(modeEnabled!==null) return modeEnabled;
    try { return localStorage.getItem('salary_tool_3d_mode')==='on'; } catch(e) { return false; }
  }
  function play(){
    // Sound volume works independently of the 3D visual mode.
    try {
      const AudioCtx=window.AudioContext||window.webkitAudioContext;
      if(!AudioCtx) return;
      if(!audioContext) audioContext=new AudioCtx();
      if(audioContext.state==='suspended') audioContext.resume();
      const now=audioContext.currentTime;
      let volume=70;
      try { volume=Math.max(0,Math.min(100,Number(localStorage.getItem('salary_tool_sound_volume')||70))); } catch(e) {}
      if(volume<=0) return;
      const volumeScale=volume/100;
      const osc=audioContext.createOscillator();
      const gain=audioContext.createGain();
      osc.type='sine';
      osc.frequency.setValueAtTime(1050,now);
      osc.frequency.exponentialRampToValueAtTime(720,now+0.045);
      gain.gain.setValueAtTime(0.0001,now);
      gain.gain.exponentialRampToValueAtTime(0.16*volumeScale,now+0.006);
      gain.gain.exponentialRampToValueAtTime(0.0001,now+0.075);
      osc.connect(gain); gain.connect(audioContext.destination);
      osc.start(now); osc.stop(now+0.08);
    } catch(_) {}
  }
  // Delegated capture listener covers regular and dynamically-created controls.
  document.addEventListener('click',function(e){
    const target=e.target;
    if(!target || !target.closest) return;
    const el=target.closest('button,input,select,textarea,a,summary,label,[role],[onclick],[tabindex],.btn,[data-bs-toggle],td,th,.calendar-card,.calendar-day,.day-cell,.date-cell,.clickable');
    if(el) play();
  },true);
  window.addEventListener('message',function(e){
    if(e.data&&e.data.type==='salary-tool-3d-mode'){
      modeEnabled=!!e.data.enabled;
      try { localStorage.setItem('salary_tool_3d_mode',modeEnabled?'on':'off'); } catch(_) {}
    }
    if(e.data&&e.data.type==='salary-tool-sound-volume'){
      const volume=Math.max(0,Math.min(100,Number(e.data.volume)||0));
      try { localStorage.setItem('salary_tool_sound_volume',String(volume)); } catch(_) {}
    }
  });
  try { modeEnabled=localStorage.getItem('salary_tool_3d_mode')==='on'; } catch(_) {}
})();

</script>
<style id="cmhkCustomDatePickerStyles">
#cmhkDatePickerPopup{position:fixed;z-index:20000;width:270px;padding:12px;border-radius:14px;
 background:var(--card-bg,rgba(255,255,255,.98));color:var(--text-color,#222);
 border:1px solid rgba(120,160,130,.45);box-shadow:0 12px 32px rgba(0,0,0,.28);
 font:14px system-ui,-apple-system,"Segoe UI",sans-serif;box-sizing:border-box}
#cmhkDatePickerPopup *{box-sizing:border-box}
#cmhkDatePickerPopup .dp-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:10px}
#cmhkDatePickerPopup .dp-title{font-weight:700;flex:1;text-align:center}
#cmhkDatePickerPopup button{border:0;border-radius:8px;cursor:pointer;color:inherit;background:rgba(120,180,130,.18)}
#cmhkDatePickerPopup .dp-nav{width:34px;height:32px;font-size:19px}
#cmhkDatePickerPopup .dp-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;text-align:center}
#cmhkDatePickerPopup .dp-week{font-size:12px;font-weight:700;opacity:.7;padding:5px 0}
#cmhkDatePickerPopup .dp-day{height:31px;font-size:13px;background:transparent}
#cmhkDatePickerPopup .dp-day:hover{background:rgba(76,175,80,.2)}
#cmhkDatePickerPopup .dp-day.dp-selected{background:#1976d2;color:#fff;font-weight:700}
#cmhkDatePickerPopup .dp-day.dp-today{outline:1px solid #4caf50}
#cmhkDatePickerPopup .dp-foot{display:flex;justify-content:space-between;margin-top:9px}
#cmhkDatePickerPopup .dp-foot button{padding:6px 10px;font-size:12px}
#cmhkDatePickerPopup.cmhk-datepicker-dark{background:#17231f!important;color:#e8f1eb!important;border:1px solid rgba(129,199,132,.6)!important;box-shadow:0 12px 34px rgba(0,0,0,.65)!important}
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-title,
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-week,
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-day{color:#e8f1eb!important}
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-nav,
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-foot button{color:#a5e6ad!important;background:rgba(76,175,80,.18)!important}
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-day:hover{background:rgba(129,199,132,.25)!important}
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-day.dp-selected{background:#2e7d32!important;color:#fff!important}
#cmhkDatePickerPopup.cmhk-datepicker-dark .dp-day.dp-today{outline-color:#81c784!important}
.cmhk-date-trigger{display:inline-flex!important;align-items:center!important;justify-content:space-between!important;gap:8px;
 min-height:30px;padding:5px 9px!important;border:1px solid #b8c9bd!important;border-radius:6px!important;
 background:#ffffff!important;color:#26352b!important;font:13px system-ui,-apple-system,"Segoe UI",sans-serif!important;
 cursor:pointer;white-space:nowrap}
.cmhk-date-trigger .dp-calendar-icon{font-size:15px;line-height:1}
.cmhk-date-trigger.cmhk-datepicker-dark{background:#17231f!important;color:#e8f1eb!important;border-color:#81c784!important}

/* CMHK date picker: explicit dark-theme palette */
body.dark-mode #cmhkDatePickerPopup,
body.night-mode #cmhkDatePickerPopup,
.dark-mode #cmhkDatePickerPopup {
    background: #17231f !important;
    color: #e8f1eb !important;
    border: 1px solid rgba(129, 199, 132, .55) !important;
    box-shadow: 0 12px 34px rgba(0, 0, 0, .62), 0 0 12px rgba(76, 175, 80, .12) !important;
}
body.dark-mode #cmhkDatePickerPopup .dp-title,
body.dark-mode #cmhkDatePickerPopup .dp-week,
body.dark-mode #cmhkDatePickerPopup .dp-day,
body.night-mode #cmhkDatePickerPopup .dp-title,
body.night-mode #cmhkDatePickerPopup .dp-week,
body.night-mode #cmhkDatePickerPopup .dp-day,
.dark-mode #cmhkDatePickerPopup .dp-title,
.dark-mode #cmhkDatePickerPopup .dp-week,
.dark-mode #cmhkDatePickerPopup .dp-day {
    color: #e8f1eb !important;
}
body.dark-mode #cmhkDatePickerPopup .dp-nav,
body.dark-mode #cmhkDatePickerPopup .dp-foot button,
body.night-mode #cmhkDatePickerPopup .dp-nav,
body.night-mode #cmhkDatePickerPopup .dp-foot button,
.dark-mode #cmhkDatePickerPopup .dp-nav,
.dark-mode #cmhkDatePickerPopup .dp-foot button {
    color: #a5e6ad !important;
    background: rgba(76, 175, 80, .16) !important;
}
body.dark-mode #cmhkDatePickerPopup .dp-day:hover,
body.dark-mode #cmhkDatePickerPopup .dp-nav:hover,
body.dark-mode #cmhkDatePickerPopup .dp-foot button:hover,
body.night-mode #cmhkDatePickerPopup .dp-day:hover,
body.night-mode #cmhkDatePickerPopup .dp-nav:hover,
body.night-mode #cmhkDatePickerPopup .dp-foot button:hover,
.dark-mode #cmhkDatePickerPopup .dp-day:hover,
.dark-mode #cmhkDatePickerPopup .dp-nav:hover,
.dark-mode #cmhkDatePickerPopup .dp-foot button:hover {
    background: rgba(129, 199, 132, .24) !important;
}
body.dark-mode #cmhkDatePickerPopup .dp-day.dp-selected,
body.night-mode #cmhkDatePickerPopup .dp-day.dp-selected,
.dark-mode #cmhkDatePickerPopup .dp-day.dp-selected {
    background: #2e7d32 !important;
    color: #ffffff !important;
}
body.dark-mode #cmhkDatePickerPopup .dp-day.dp-today,
body.night-mode #cmhkDatePickerPopup .dp-day.dp-today,
.dark-mode #cmhkDatePickerPopup .dp-day.dp-today {
    outline-color: #81c784 !important;
}

</style>
<script id="cmhkCustomDatePickerScript">
(function(){
 function initCmhkDatePicker(){
  if(window.__cmhkCustomDatePickerReady)return;
  window.__cmhkCustomDatePickerReady=true;
  const popup=document.createElement('div');popup.id='cmhkDatePickerPopup';popup.hidden=true;
  popup.setAttribute('role','dialog');popup.setAttribute('aria-label','選擇日期');
  document.body.appendChild(popup);
  // Prevent navigation clicks from reaching the outside-click closer.
  // render() replaces popup.innerHTML and detaches the clicked arrow.
  popup.addEventListener('click',e=>e.stopPropagation());
  function syncPickerTheme(){
   const cls=(document.body.className||'').toLowerCase();
   const isDark=/(dark|night|black)/.test(cls);
   popup.classList.toggle('cmhk-datepicker-dark',isDark);
   document.querySelectorAll('.cmhk-date-trigger').forEach(el=>el.classList.toggle('cmhk-datepicker-dark',isDark));
  }
  syncPickerTheme();
  new MutationObserver(syncPickerTheme).observe(document.body,{attributes:true,attributeFilter:['class','style']});
  let activeInput=null,viewYear=0,viewMonth=0;
  const weekdays=['日','一','二','三','四','五','六'];
  const pad=n=>String(n).padStart(2,'0');
  const iso=(y,m,d)=>`${y}-${pad(m+1)}-${pad(d)}`;
  function parseISO(v){const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(v||'');return m?{y:+m[1],m:+m[2]-1,d:+m[3]}:null;}
  function today(){const d=new Date();return {y:d.getFullYear(),m:d.getMonth(),d:d.getDate()};}
  function format(v){const d=parseISO(v);return d?`${pad(d.d)}/${pad(d.m+1)}/${d.y}`:'選擇日期';}
  function close(){popup.hidden=true;activeInput=null;}
  function render(){
   const selected=parseISO(activeInput.value),first=new Date(viewYear,viewMonth,1).getDay(),days=new Date(viewYear,viewMonth+1,0).getDate();
   popup.innerHTML='';
   const head=document.createElement('div');head.className='dp-head';
   const prev=document.createElement('button');prev.type='button';prev.className='dp-nav';prev.textContent='‹';prev.setAttribute('aria-label','上一個月');
   prev.addEventListener('click',()=>{viewMonth--;if(viewMonth<0){viewMonth=11;viewYear--;}render();});
   const title=document.createElement('div');title.className='dp-title';title.textContent=new Date(viewYear,viewMonth,1).toLocaleDateString('en',{month:'long',year:'numeric'});
   const next=document.createElement('button');next.type='button';next.className='dp-nav';next.textContent='›';next.setAttribute('aria-label','下一個月');
   next.addEventListener('click',()=>{viewMonth++;if(viewMonth>11){viewMonth=0;viewYear++;}render();});
   head.append(prev,title,next);popup.append(head);
   const grid=document.createElement('div');grid.className='dp-grid';
   weekdays.forEach(w=>{const el=document.createElement('div');el.className='dp-week';el.textContent=w;grid.append(el);});
   for(let i=0;i<first;i++){const blank=document.createElement('span');grid.append(blank);}
   for(let d=1;d<=days;d++){
    const b=document.createElement('button');b.type='button';b.className='dp-day';b.textContent=d;
    if(selected&&selected.y===viewYear&&selected.m===viewMonth&&selected.d===d)b.classList.add('dp-selected');
    const t=today();if(t.y===viewYear&&t.m===viewMonth&&t.d===d)b.classList.add('dp-today');
    b.addEventListener('click',()=>{activeInput.value=iso(viewYear,viewMonth,d);activeInput.dispatchEvent(new Event('input',{bubbles:true}));activeInput.dispatchEvent(new Event('change',{bubbles:true}));const tr=activeInput._cmhkDateTrigger;if(tr)tr.querySelector('.dp-date-label').textContent=format(activeInput.value);close();});
    grid.append(b);
   }
   popup.append(grid);
   const foot=document.createElement('div');foot.className='dp-foot';
   const clear=document.createElement('button');clear.type='button';clear.textContent='清除';clear.addEventListener('click',()=>{activeInput.value='';activeInput.dispatchEvent(new Event('change',{bubbles:true}));const tr=activeInput._cmhkDateTrigger;if(tr)tr.querySelector('.dp-date-label').textContent='選擇日期';close();});
   const now=document.createElement('button');now.type='button';now.textContent='今天';now.addEventListener('click',()=>{const t=today();activeInput.value=iso(t.y,t.m,t.d);activeInput.dispatchEvent(new Event('input',{bubbles:true}));activeInput.dispatchEvent(new Event('change',{bubbles:true}));const tr=activeInput._cmhkDateTrigger;if(tr)tr.querySelector('.dp-date-label').textContent=format(activeInput.value);close();});
   foot.append(clear,now);popup.append(foot);
  }
  function open(input,trigger){
   activeInput=input;const d=parseISO(input.value)||today();viewYear=d.y;viewMonth=d.m;render();popup.hidden=false;
   const r=trigger.getBoundingClientRect(),w=270,h=popup.offsetHeight||300;
   popup.style.left=Math.max(8,Math.min(r.left,window.innerWidth-w-8))+'px';
   popup.style.top=(r.bottom+h+8>window.innerHeight?Math.max(8,r.top-h-8):r.bottom+6)+'px';
  }
  document.querySelectorAll('input[type="date"]').forEach(input=>{
   if(input.dataset.cmhkPickerBound==='1')return;
   input.dataset.cmhkPickerBound='1';
   const trigger=document.createElement('button');trigger.type='button';trigger.className='cmhk-date-trigger';
   const label=document.createElement('span');label.className='dp-date-label';label.textContent=format(input.value);
   const icon=document.createElement('span');icon.className='dp-calendar-icon';icon.textContent='📅';
   trigger.append(label,icon);input._cmhkDateTrigger=trigger;input.dataset.requiredDate=input.required?'1':'0';
   input.type='hidden';input.required=false;input.insertAdjacentElement('afterend',trigger);
   trigger.addEventListener('click',()=>open(input,trigger));
  });
  document.addEventListener('click',e=>{if(!popup.hidden&&!popup.contains(e.target)&&!e.target.closest('.cmhk-date-trigger'))close();});
  document.addEventListener('keydown',e=>{if(e.key==='Escape')close();});
  document.querySelectorAll('form').forEach(form=>form.addEventListener('submit',e=>{
   const empty=[...form.querySelectorAll('input[type="hidden"]')].filter(i=>i._cmhkDateTrigger&&i.dataset.requiredDate==='1'&&!i.value);
   if(empty.length){e.preventDefault();empty[0]._cmhkDateTrigger.click();}
  }));
  window.addEventListener('resize',close);window.addEventListener('scroll',close,true);
 }
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initCmhkDatePicker,{once:true});
 else initCmhkDatePicker();
})();
</script>

</body>
</html>
"""

@app.errorhandler(Exception)
def handle_unexpected_error(exc):
                                                                                    
    import traceback
    traceback.print_exc()
    return (
        "<!doctype html><html lang=zh-HK><head><meta charset=utf-8>"
        "<title>CMHK Internal Server Error</title></head>"
        '<body style="font-family:monospace;background:#111;color:#eee;padding:20px">' 
        "<h2>CMHK 發生程式錯誤</h2>"
        f'<pre style="white-space:pre-wrap;color:#ffb4b4">{type(exc).__name__}: {exc}</pre>' 
        "</body></html>", 500
    )

@app.route("/get_language", methods=["GET"])
def get_language():
                                                                                   
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        lang = cfg.get("language", "zh-Hant")
    except Exception:
        lang = "zh-Hant"
    if lang not in ("zh-Hant", "en", "th"):
        lang = "zh-Hant"
    from flask import jsonify
    return jsonify({"language": lang})

@app.route("/", methods=["GET"])
def index():
    data = load_data()
    mode = request.args.get("mode", "add")
    flash_error = session.pop("flash_error", None)
    flash_success = session.pop("flash_success", None)

    start, end, records_dict, total_h, total_t, work_days = filter_period_records(data)
    previous_start, previous_end, previous_records_dict, previous_total_h, previous_total_t, previous_work_days = filter_period_records(data, "previous")

    indexed_data = [{"original_idx": i, "data": r} for i, r in enumerate(data)]
    sorted_indexed_data = sorted(indexed_data, key=lambda x: str(x["data"].get("date", "")), reverse=True)

    selected_date = request.args.get("selected_date", "").strip()
    selected_edit_idx = request.args.get("edit_idx", type=int)
    if selected_edit_idx is None and selected_date:
        selected_edit_idx = next((i for i, item in enumerate(data) if item.get("date") == selected_date), None)
    if selected_edit_idx is None and sorted_indexed_data:
        selected_edit_idx = sorted_indexed_data[0]["original_idx"]

    edit_item = data[selected_edit_idx] if (selected_edit_idx is not None and 0 <= selected_edit_idx < len(data)) else None

    months = {}
    for r in data:
        try:
            d = datetime.strptime(str(r.get("date", "")), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            continue
        h, t = calc_record(r)
        key = d.strftime("%Y-%m")
        months.setdefault(key, {"days": {}, "hours": 0, "total_pay": 0})
        months[key]["days"][d.day] = {"pay": int(t), "note": str(r.get("note", "")).lower().strip(), "hours": h, "index": data.index(r)}
        months[key]["hours"] += h

    all_calendar_months = {}
    sorted_months_keys = sorted(months.keys(), reverse=True)
    for m in sorted_months_keys:
        y, mo = map(int, m.split("-"))
        total_m = sum(info["pay"] for info in months[m]["days"].values())

        first_day = date(y, mo, 1)
        next_month = date(y + 1, 1, 1) if mo == 12 else date(y, mo + 1, 1)
        last_day = next_month - timedelta(days=1)
        start_wday = (first_day.weekday() + 1) % 7

        cal = []
        current_week = [0] * 7
        slot = start_wday
        for d_num in range(1, last_day.day + 1):
            current_week[slot] = d_num
            slot += 1
            if slot == 7:
                cal.append(current_week)
                current_week = [0] * 7
                slot = 0
        if slot > 0:
            cal.append(current_week)

        all_calendar_months[m] = {
            "hours": months[m]["hours"],
            "total_pay": total_m,
            "days": months[m]["days"],
            "cal": cal
        }

    recent_calendar_months = {}
    older_calendar_months = {}
    for i, m in enumerate(sorted_months_keys):
        if i < 2:
            recent_calendar_months[m] = all_calendar_months[m]
        else:
            older_calendar_months[m] = all_calendar_months[m]

    table_data = []
    curr = start
    while curr <= end:
        h, t, note = records_dict.get(curr, (0, 0, ""))
        if h > 0 and t > 0:
            table_data.append({
                "date": curr.strftime("%m/%d"),
                "weekday": WEEKDAY_CN[curr.weekday()],
                "hours": int(h),
                "pay": f"$ {int(t)}",
                "note": "🟢 Paid" if note.lower().strip() == "paid" else ("" if note.lower().strip() in ("null", "none") else note)
            })
        curr += timedelta(days=1)

    total_sum = 0
    if data:
        df_all = pd.DataFrame(data)
        for col, default in [("hours", 0), ("rate", 125), ("note", "")]:
            if col not in df_all.columns:
                df_all[col] = default
        df_all["pay"] = df_all.apply(
            lambda row: 0 if (str(row["note"]).lower().strip() == "off" or float(row["hours"]) * float(row["rate"]) == 0) else int(float(row["hours"]) * float(row["rate"])),
            axis=1
        )
        total_sum = int(df_all["pay"].sum())

    today_date = date.today()

    return render_template_string(
        HTML_TEMPLATE,
        data=data,
        mode=mode,
        today_str=str(today_date),
        selected_date=selected_date,
        start=start,
        end=end,
        total_h=total_h,
        total_t=total_t,
        work_days=work_days,
        previous_start=previous_start,
        previous_end=previous_end,
        previous_total_h=previous_total_h,
        previous_total_t=previous_total_t,
        previous_work_days=previous_work_days,
        sorted_indexed_data=[(item["original_idx"], item["data"]) for item in sorted_indexed_data],
        selected_edit_idx=selected_edit_idx,
        edit_item=edit_item,
        recent_calendar_months=recent_calendar_months,
        older_calendar_months=older_calendar_months,
        current_ym=today_date.strftime("%Y-%m"),
        current_day=today_date.day,
        table_data=table_data,
        total_sum=total_sum,
        flash_error=flash_error,
        flash_success=flash_success
    )

@app.route("/set_theme", methods=["POST"])
def set_theme():
    theme = request.form.get("theme", "dark-mode")
    session["theme"] = theme
    return redirect(request.referrer or url_for("index"))

@app.route("/add", methods=["POST"])
def add_record():
    data = load_data()
    dt = request.form.get("date")
    note = resolve_note_form(request.form)

    if any(item.get("date") == dt for item in data):
        session["flash_error"] = f"⚠️ 新增失敗：日期 {dt} 已經存在，不能重複新增相同日子的記錄！"
        return redirect(url_for("index", mode="add"))

    is_off = note.lower() == "off"
    hours = 0 if is_off else int(request.form.get("hours", DEFAULT_HOURS))
    rate = 0 if is_off else int(request.form.get("rate", DEFAULT_RATE))

    data.append({"date": dt, "hours": hours, "rate": rate, "note": note})
    data.sort(key=lambda x: x["date"])
    save_data(data)
    return redirect(url_for("index", mode="add"))

@app.route("/copy", methods=["POST"])
def copy_record():
    data = load_data()
    source_date = request.form.get("source_date", "").strip()
    target_start = request.form.get("target_start", "").strip()
    target_end = request.form.get("target_end", "").strip() or target_start
    source = next((item for item in data if item.get("date") == source_date), None)
    try:
        start_d = datetime.strptime(target_start, "%Y-%m-%d").date()
        end_d = datetime.strptime(target_end, "%Y-%m-%d").date()
        if end_d < start_d:
            raise ValueError
    except (ValueError, TypeError):
        start_d = end_d = None

    if not source:
        session["flash_error"] = "⚠️ 找不到要複製的原有日期記錄！"
    elif start_d is None:
        session["flash_error"] = "⚠️ 請選擇有效的複製日期範圍（結束日期不可早於開始日期）。"
    else:
        existing = {item.get("date") for item in data}
        added = skipped = 0
        day = start_d
        while day <= end_d:
            target_date = day.isoformat()
            if target_date in existing:
                skipped += 1
            else:
                data.append({"date": target_date, "hours": source.get("hours", DEFAULT_HOURS), "rate": source.get("rate", DEFAULT_RATE), "note": source.get("note", "")})
                existing.add(target_date)
                added += 1
            day += timedelta(days=1)
        if added:
            data.sort(key=lambda item: item["date"])
            save_data(data)
        session["flash_success"] = f"✅ 批量複製完成：新增 {added} 日，略過已有記錄 {skipped} 日。"
    return redirect(url_for("index", mode="add"))

@app.route("/update", methods=["POST"])
def update_record():
    data = load_data()
    idx = request.form.get("idx", type=int)
    action = request.form.get("action")

    if idx is not None and 0 <= idx < len(data):
        if action == "delete":
            data.pop(idx)
        elif action == "update":
            dt = request.form.get("date")
            note = resolve_note_form(request.form)
            is_off = note.lower() == "off"
            hours = 0 if is_off else int(request.form.get("hours", 0))
            rate = 0 if is_off else int(request.form.get("rate", 0))

            data[idx] = {"date": dt, "hours": hours, "rate": rate, "note": note}
            data.sort(key=lambda x: x["date"])
        save_data(data)

    return redirect(url_for("index", mode="edit", edit_idx=idx if action != "delete" else None))

@app.route("/download/<file_type>")
def download(file_type):
    data = load_data()
    period = request.args.get("period", "current")
    if period not in ("current", "previous"):
        period = "current"
    start_dl, end_dl, _, _, _, _ = filter_period_records(data, period)

    if file_type == "excel":
        file_obj = generate_excel(data, period)
        return send_file(file_obj, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name=f"pay_{start_dl}_{end_dl}.xlsx")
    elif file_type == "pdf":
        file_obj = generate_pdf(data, period)
        return send_file(file_obj, mimetype="application/pdf",
                         as_attachment=True, download_name=f"salary_{start_dl}_{end_dl}.pdf")
    elif file_type == "jpg":
        file_obj = generate_jpg(data, period)
        return send_file(file_obj, mimetype="image/jpeg",
                         as_attachment=True, download_name=f"pay_{start_dl}_{end_dl}.jpg")

    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8504, debug=True)
