from flask import (
    Flask,
    render_template_string,
    request,
    redirect,
    url_for,
    send_file,
    jsonify,
)
import json
import os
from datetime import date, datetime, timedelta
from io import BytesIO

app = Flask(__name__)
app.secret_key = os.urandom(24)

app.config["SESSION_COOKIE_NAME"] = "cargo_session"
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = True                                      

app.jinja_env.globals.update(enumerate=enumerate)
app.jinja_env.filters["money"] = lambda v: f"{int(v or 0):,}"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "salary_cargo.json")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
PAID_REPORT_DIR = os.path.join(BASE_DIR, "export")
DEFAULT_HOURS = 10
DEFAULT_RATE = 75
WEEKDAY_CN = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"]
WEEKDAY_EXCEL = ["一", "二", "三", "四", "五", "六", "日"]

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

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def resolve_note_form(form, allowed):
                                        
    choice = form.get("note_type")
    if choice is None:
        return str(form.get("note", "")).strip()
    choice = str(choice).strip()
    if choice.lower() == "other":
        return str(form.get("custom_note", "")).strip()
    allowed_map = {str(x).lower(): str(x) for x in allowed}
    return allowed_map.get(choice.lower(), choice)

def calc_record(r):
    if r["note"].lower() in ("off", "cut"):
        return 0, 0
    hours = int(r["hours"])
    pay = int(hours * r["rate"])
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
       
    today = date.today()
    data = data or []
    dated_records = []
    paid_dates = []

    for record in data:
        try:
            record_date = datetime.strptime(record.get("date", ""), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            continue
        dated_records.append(record_date)
        if str(record.get("note", "")).strip().lower() == "paid":
            paid_dates.append(record_date)

    paid_dates = sorted(set(paid_dates))
    earliest_date = min(dated_records) if dated_records else today

    if period == "previous":
        if not paid_dates:
                                          
            return today, today - timedelta(days=1)
        latest_paid = paid_dates[-1]
        previous_start = paid_dates[-2] + timedelta(days=1) if len(paid_dates) >= 2 else earliest_date
        return previous_start, latest_paid

    current_start = paid_dates[-1] + timedelta(days=1) if paid_dates else earliest_date
    return current_start, today

def get_display_period_start(start, records):
                                        
    positive_pay_dates = [d for d, values in records.items() if len(values) > 1 and values[1] > 0]
    return min(positive_pay_dates) if positive_pay_dates else start

def filter_period_records(data, period="current"):
    start, end = get_pay_period(data, period)
    records = {}
    total_h = total_t = work_days = 0
    if start > end:
        return start, end, records, total_h, total_t, work_days

    for r in data:
        try:
            d = datetime.strptime(r["date"], "%Y-%m-%d").date()
        except (KeyError, TypeError, ValueError):
            continue
        if start <= d <= end:
            h, t = calc_record(r)
                                  
            if t <= 0:
                continue
            records[d] = (h, t, r.get("note", ""))
            total_h += h
            total_t += t
            if h > 0 and t > 0:
                work_days += 1
    return start, end, records, total_h, total_t, work_days

def get_calendar_matrix(y, mo):
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
    return cal

def generate_excel(data, period="current"):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

    start_date, end_date, records, total_h, total_t, work_days = (
        filter_period_records(data, period)
    )
    start_date = get_display_period_start(start_date, records)
    wb = Workbook()
    ws = wb.active
    ws.title = "工資表"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 30
    center = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    def set_cell(cell, value, bg=None):
        c = ws[cell]
        c.value = value
        c.alignment = center
        c.border = border
        if bg:
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")

    row = 1
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"] = f"出糧週期工資表（{start_date} ～ {end_date}）"
    ws[f"A{row}"].alignment = center
    ws[f"A{row}"].font = Font(size=14, bold=True)
    row += 2
    set_cell(f"A{row}", "工作日數", "E8F5E9")
    set_cell(f"B{row}", work_days)
    set_cell(f"C{row}", "總工資", "E8F5E9")
    set_cell(f"D{row}", f"$ {int(total_t):,}", "FFF45C")
    row += 2
    for col, h in zip(["A", "B", "C", "D"], ["日期", "星期", "工時", "每日工資"]):
        set_cell(f"{col}{row}", h, "7ED957")
    row += 1
    for d in sorted(records.keys()):
        h, t, _ = records[d]
        values = [
            d.strftime("%m/%d"),
            f"星期{WEEKDAY_EXCEL[d.weekday()]}",
            int(h),
            f"$ {int(t):,}",
        ]
        for col, v in zip(["A", "B", "C", "D"], values):
            set_cell(f"{col}{row}", v)
        row += 1
    ws.merge_cells(f"A{row}:B{row}")
    set_cell(f"A{row}", f"總工時：{int(total_h)}", "E3F2FD")
    ws.merge_cells(f"C{row}:D{row}")
    set_cell(f"C{row}", f"總工資：$ {int(total_t):,}", "FFF9C4")
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

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
    start = get_display_period_start(start, records)
    rows = [["Date", "Day", "Hours", "Daily Pay"]]
    for d in sorted(records.keys()):
        h, p, _ = records[d]
        if h > 0:
            rows.append(
                [d.strftime("%m/%d"), d.strftime("%a"), str(int(h)), f"$ {int(p):,}"]
            )
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4)
    styles = getSampleStyleSheet()
    styles["Normal"].fontName = FONT_NAME
    elements = [
        Paragraph(f"Salary Report ({start} to {end})", styles["Normal"]),
        Paragraph(
            f"Working Days: {work_days} | Total Pay: $ {int(total_pay):,}",
            styles["Normal"],
        ),
        Table(
            rows,
            colWidths=[100] * 4,
            style=TableStyle([
                ("FONT", (0, 0), (-1, -1), FONT_NAME),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7ED957")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]),
        ),
    ]
    doc.build(elements)
    return output.getvalue()

def generate_jpg(data, period="current"):
    from PIL import Image, ImageDraw, ImageFont

    start_date, end_date, records, _, total_t, work_days = filter_period_records(
        data, period
    )
    start_date = get_display_period_start(start_date, records)
    filtered_days = [
        d for d in sorted(records.keys()) if records[d][0] > 0 and records[d][1] > 0
    ]
    row_h = 55
    img_h = 240 + row_h * max(len(filtered_days), 1)
    font_path = os.path.join(BASE_DIR, "NotoSansCJK-Regular.ttc")
    try:
        font = ImageFont.truetype(font_path, 26)
    except Exception:
        font = ImageFont.load_default()
    img = Image.new("RGB", (760, img_h), "white")
    draw = ImageDraw.Draw(img)

    def draw_cell(x, y, w, h, text, bg="#FFFFFF", border="#000000"):
        draw.rectangle([x, y, x + w, y + h], fill=bg, outline=border)
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        except Exception:
            tw, th = 50, 20
        draw.text((x + (w - tw) // 2, y + (h - th) // 2), text, fill="black", font=font)

    x0, y = 20, 20
    draw.text(
        (x0, y),
        f"出糧週期工資表（{start_date} ～ {end_date}）",
        fill="black",
        font=font,
    )
    y += 60
    draw_cell(x0, y, 160, row_h, "工作日數", bg="#E8F5E9")
    draw_cell(x0 + 160, y, 120, row_h, str(work_days))
    draw_cell(x0 + 280, y, 160, row_h, "總工資", bg="#E8F5E9")
    draw_cell(x0 + 440, y, 260, row_h, f"$ {int(total_t):,}", bg="#FFF45C")
    y += row_h + 20
    col_w = [150, 150, 150, 250]
    x = x0
    for i, h_text in enumerate(["日期", "星期", "工時 (Hrs)", "每日工資"]):
        draw_cell(x, y, col_w[i], row_h, h_text, bg="#7ED957")
        x += col_w[i]
    y += row_h
    for d in filtered_days:
        h, t, _ = records[d]
        values = [
            d.strftime("%m/%d"),
            f"星期{WEEKDAY_EXCEL[d.weekday()]}",
            str(int(h)),
            f"$ {int(t):,}",
        ]
        x = x0
        for i, v in enumerate(values):
            draw_cell(x, y, col_w[i], row_h, v)
            x += col_w[i]
        y += row_h
    output = BytesIO()
    img.save(output, format="JPEG", quality=95)
    return output.getvalue()

@app.route("/get_language", methods=["GET"])
def get_language():
                                                                                 
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        lang = cfg.get("language", "zh-Hant")
    except (OSError, ValueError, TypeError):
        lang = "zh-Hant"
    if lang not in ("zh-Hant", "en", "th"):
        lang = "zh-Hant"
    return jsonify({"language": lang})

@app.route("/", methods=["GET", "POST"])
def index():
    data = load_data()
    msg = None
    msg_type = "success"

    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            input_date = request.form.get("date")
            note = resolve_note_form(request.form, ("null", "Paid", "Off", "Cut"))
            try:
                hours = int(request.form.get("hours", 0))
                rate = int(request.form.get("rate", 0))
            except ValueError:
                hours, rate = 0, 0

            if any(r["date"] == input_date for r in data):
                msg = f"該日期 ({input_date}) 的記錄已存在！"
                msg_type = "warning"
            else:
                data.append(
                    {"date": input_date, "hours": hours, "rate": rate, "note": note}
                )
                data.sort(key=lambda x: x["date"])
                save_data(data)
                msg = "成功新增記錄！"

        elif action == "copy":
            source_date = request.form.get("source_date", "").strip()
            target_start = request.form.get("target_start", "").strip()
            target_end = request.form.get("target_end", "").strip() or target_start
            source = next((r for r in data if r.get("date") == source_date), None)
            try:
                start_d = datetime.strptime(target_start, "%Y-%m-%d").date()
                end_d = datetime.strptime(target_end, "%Y-%m-%d").date()
                if end_d < start_d:
                    raise ValueError
            except (ValueError, TypeError):
                start_d = end_d = None
            if not source:
                msg, msg_type = "找不到要複製的原有日期記錄！", "warning"
            elif start_d is None:
                msg, msg_type = "請選擇有效的複製日期範圍（結束日期不可早於開始日期）。", "warning"
            else:
                existing = {r.get("date") for r in data}
                added = skipped = 0
                day = start_d
                while day <= end_d:
                    target_date = day.isoformat()
                    if target_date in existing:
                        skipped += 1
                    else:
                        data.append({"date": target_date, "hours": source.get("hours", 10), "rate": source.get("rate", 75), "note": source.get("note", "")})
                        existing.add(target_date)
                        added += 1
                    day += timedelta(days=1)
                if added:
                    data.sort(key=lambda x: x["date"])
                    save_data(data)
                msg = f"批量複製完成：新增 {added} 日，略過已有記錄 {skipped} 日。"
                if skipped:
                    msg_type = "warning"

        elif action == "update":
            try:
                idx = int(request.form.get("index"))
            except (TypeError, ValueError):
                idx = -1
            edit_date = request.form.get("date")
            edit_note = resolve_note_form(request.form, ("null", "Paid", "Off", "Cut"))
            edit_is_off = edit_note.lower() in ("off", "cut")
            try:
                edit_hours = (
                    0 if edit_is_off else int(request.form.get("hours", 0))
                )
                edit_rate = 0 if edit_is_off else int(request.form.get("rate", 0))
            except ValueError:
                edit_hours, edit_rate = 0, 0

            if 0 <= idx < len(data):
                data[idx] = {
                    "date": edit_date,
                    "hours": edit_hours,
                    "rate": edit_rate,
                    "note": edit_note,
                }
                data.sort(key=lambda x: x["date"])
                save_data(data)
                msg = "記錄更新成功！"

        elif action == "delete":
            try:
                idx = int(request.form.get("index"))
            except (TypeError, ValueError):
                idx = -1
            if 0 <= idx < len(data):
                del data[idx]
                save_data(data)
                msg = "記錄已刪除。"
                msg_type = "warning"

    selected_month = request.args.get("month", None)
    months = {}
    for r in data:
        d = datetime.strptime(r["date"], "%Y-%m-%d").date()
        h, t = calc_record(r)
        key = d.strftime("%Y-%m")
        months.setdefault(key, {"days": {}, "hours": 0})
        note_val = str(r.get("note", "")).lower().strip()
        months[key]["days"][d.day] = {
            "pay": int(t),
            "note": note_val,
            "hours": h,
        }
        months[key]["hours"] += h

    sorted_months = sorted(months.keys(), reverse=True)
    if not selected_month or selected_month not in months:
        selected_month = (
            sorted_months[0] if sorted_months else datetime.now().strftime("%Y-%m")
        )

    all_calendar_months = {}
    for m in sorted_months:
        y_m, mo_m = map(int, m.split("-"))
        cal_weeks_m = get_calendar_matrix(y_m, mo_m)
        total_pay_m = sum(int(day_info.get("pay", 0)) for day_info in months[m]["days"].values())
        all_calendar_months[m] = {
            "hours": months[m]["hours"],
            "total_pay": total_pay_m,
            "days": months[m]["days"],
            "cal": cal_weeks_m,
        }

    recent_calendar_months = {}
    older_calendar_months = {}
    for i, m in enumerate(sorted_months):
        if i < 2:
            recent_calendar_months[m] = all_calendar_months[m]
        else:
            older_calendar_months[m] = all_calendar_months[m]

    start, end, records_dict, total_h, total_t, work_days = filter_period_records(
        data, "current"
    )
    previous_start, previous_end, previous_records_dict, previous_total_h, previous_total_t, previous_work_days = filter_period_records(
        data, "previous"
    )
    start = get_display_period_start(start, records_dict)
    previous_start = get_display_period_start(previous_start, previous_records_dict)

    reversed_indexed_data = list(reversed(list(enumerate(data))))

    total_income = 0
    for r in data:
        _, pay = calc_record(r)
        total_income += int(pay)

    charts_html = None

    return render_template_string(
        HTML_TEMPLATE,
        data=data,
        reversed_indexed_data=reversed_indexed_data,
        sorted_months=sorted_months,
        selected_month=selected_month,
        months=months,
        recent_calendar_months=recent_calendar_months,
        older_calendar_months=older_calendar_months,
        start=start,
        end=end,
        records_dict=records_dict,
        total_h=total_h,
        total_t=total_t,
        work_days=work_days,
        previous_start=previous_start,
        previous_end=previous_end,
        previous_records_dict=previous_records_dict,
        previous_total_h=previous_total_h,
        previous_total_t=previous_total_t,
        previous_work_days=previous_work_days,
        msg=msg,
        msg_type=msg_type,
        charts=charts_html,
        total_income=total_income,
        WEEKDAY_CN=WEEKDAY_CN,
        now_date=date.today().strftime("%Y-%m-%d"),
        current_year=date.today().year,
        current_month=date.today().month,
        current_day=date.today().day,
        today_str=datetime.now().strftime("%Y-%m-%d"),
    )

@app.route("/charts")
def charts():
                               
    data = load_data()
    selected_month = request.args.get("month", "")
    theme = request.args.get("theme", "dark")
    lang = request.args.get("lang", "zh-Hant")
    if lang not in ("zh-Hant", "en", "th"):
        lang = "zh-Hant"
    theme_name = "plotly_white" if theme == "light" else "plotly_dark"

    chart_labels = {
        "zh-Hant": {"正常返工": "正常返工", "放假 (Off)": "放假 (Off)", "被削班 (Cut)": "被削班 (Cut)", "狀態": "狀態", "天數": "天數"},
        "en": {"正常返工": "Working", "放假 (Off)": "Day Off", "被削班 (Cut)": "Cut", "狀態": "Status", "天數": "Days"},
        "th": {"正常返工": "ทำงานปกติ", "放假 (Off)": "วันหยุด", "被削班 (Cut)": "ลดกะ (Cut)", "狀態": "สถานะ", "天數": "จำนวนวัน"},
    }[lang]

    if not data:
        return jsonify({"charts": [], "total_income": 0, "prev_month_str": ""})

    import plotly.express as px

    status_all = {}
    month_status = {}
    total_income = 0

    for r in data:
        try:
            d = datetime.strptime(r["date"], "%Y-%m-%d").date()
        except (KeyError, TypeError, ValueError):
            continue
        note = str(r.get("note", "")).lower().strip()
        try:
            hours = float(r.get("hours", 0) or 0)
        except (TypeError, ValueError):
            hours = 0.0
        try:
            rate = float(r.get("rate", DEFAULT_RATE) or DEFAULT_RATE)
        except (TypeError, ValueError):
            rate = float(DEFAULT_RATE)
        pay = hours * rate

        if note == "cut":
            status = "被削班 (Cut)"
        elif note == "off" or pay == 0 or hours == 0:
            status = "放假 (Off)"
        else:
            status = "正常返工"
            total_income += int(pay)

        status_all[status] = status_all.get(status, 0) + 1
        month = d.strftime("%Y-%m")
        month_status.setdefault(month, {})
        month_status[month][status] = month_status[month].get(status, 0) + 1

    if not selected_month or selected_month not in month_status:
        selected_month = max(month_status) if month_status else date.today().strftime("%Y-%m")

    curr_date = datetime.strptime(selected_month + "-01", "%Y-%m-%d")
    prev_date = curr_date.replace(day=1) - timedelta(days=1)
    prev_month_str = prev_date.strftime("%Y-%m")

    color_map = {
        "正常返工": "#00F5D4",
        "放假 (Off)": "#00A6FB",
        "被削班 (Cut)": "#FF2E88",
    }

    def make_pie(status_counts):
                                                                            
        rows = [{chart_labels["狀態"]: chart_labels.get(k, k), chart_labels["天數"]: v} for k, v in status_counts.items()]
        translated_colors = {chart_labels.get(k, k): v for k, v in color_map.items()}
        fig = px.pie(
            rows,
            names=chart_labels["狀態"],
            values=chart_labels["天數"],
            color=chart_labels["狀態"],
            color_discrete_map=translated_colors,
            hole=0.3,
        )
        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(3, 8, 24, 0.18)",
            plot_bgcolor="rgba(3, 8, 24, 0.08)",
            font=dict(color="#B8F7FF", family="Orbitron, Rajdhani, Arial, sans-serif"),
            margin=dict(t=26, b=38, l=24, r=24),
            height=280,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.35,
                xanchor="center",
                x=0.5,
                font=dict(color="#9CFBFF", size=11),
                bgcolor="rgba(0, 12, 30, 0.48)",
                bordercolor="rgba(0, 245, 212, 0.38)",
                borderwidth=1,
            ),
        )
        fig.update_traces(
            textinfo="label+percent+value",
            textfont=dict(size=12, color="#FFFFFF", family="Orbitron, Rajdhani, Arial, sans-serif"),
            pull=[0.055] * len(rows),
            rotation=18,
            sort=False,
            marker=dict(line=dict(color="#07152F", width=3)),
            hovertemplate="<b>%{label}</b><br>數量：%{value}<br>比例：%{percent}<extra></extra>",
        )
        return fig.to_json()

    chart_json = [make_pie(status_all)]
    chart_json.append(make_pie(month_status[selected_month]) if selected_month in month_status else None)
    chart_json.append(make_pie(month_status[prev_month_str]) if prev_month_str in month_status else None)

    return jsonify({
        "charts": chart_json,
        "total_income": total_income,
        "prev_month_str": prev_month_str,
        "selected_month": selected_month,
    })

@app.route("/export/<file_type>")
def export_file(file_type):
    data = load_data()
    period = request.args.get("period", "current")
    if period not in ("current", "previous"):
        period = "current"
    start_dl, end_dl, export_records, _, _, _ = filter_period_records(data, period)
    start_dl = get_display_period_start(start_dl, export_records)
    if file_type == "excel":
        content = generate_excel(data, period)
        return send_file(
            BytesIO(content),
            mimetype=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            as_attachment=True,
            download_name=f"pay_{start_dl}_{end_dl}.xlsx",
        )
    elif file_type == "pdf":
        content = generate_pdf(data, period)
        return send_file(
            BytesIO(content),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"salary_{start_dl}_{end_dl}.pdf",
        )
    elif file_type == "jpg":
        content = generate_jpg(data, period)
                                                                                 
        os.makedirs(PAID_REPORT_DIR, exist_ok=True)
        jpg_name = f"pay_{start_dl}_{end_dl}.jpg"
        archive_name = jpg_name
        archive_path = os.path.join(PAID_REPORT_DIR, archive_name)
                                                                                   
        if os.path.exists(archive_path):
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            archive_name = f"pay_{start_dl}_{end_dl}_{stamp}.jpg"
            archive_path = os.path.join(PAID_REPORT_DIR, archive_name)
        with open(archive_path, "wb") as jpg_file:
            jpg_file.write(content)
                                           
        return redirect(url_for("index") + "#paidReportsCard")
    return redirect(url_for("index"))

@app.route("/paid-reports/list")
def list_paid_reports():
                                                                              
    os.makedirs(PAID_REPORT_DIR, exist_ok=True)
    files = []
    for name in os.listdir(PAID_REPORT_DIR):
        full_path = os.path.join(PAID_REPORT_DIR, name)
        if not os.path.isfile(full_path) or name.startswith("."):
            continue
        try:
            stat = os.stat(full_path)
        except OSError:
            continue
        ext = os.path.splitext(name)[1].lower()
        files.append({
            "name": name,
            "modified": stat.st_mtime,
            "modified_text": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "size": stat.st_size,
            "size_text": (f"{stat.st_size / 1024:.1f} KB" if stat.st_size < 1024 * 1024 else f"{stat.st_size / (1024 * 1024):.2f} MB"),
            "is_image": ext in (".jpg", ".jpeg", ".png", ".webp", ".gif"),
            "preview_url": url_for("preview_paid_report", filename=name),
        })
    files.sort(key=lambda item: item["modified"], reverse=True)
    return jsonify({"files": files, "folder": "export"})

@app.route("/paid-reports/upload", methods=["POST"])
def upload_paid_report():
                                                                          
    uploaded = request.files.get("file")
    if uploaded is None or not uploaded.filename:
        return jsonify({"ok": False, "message": "請先選擇要上載的檔案。"}), 400

    original_name = uploaded.filename
                                                                                       
    safe_name = original_name.replace("\\", "/").rsplit("/", 1)[-1]
    if not safe_name or safe_name in (".", "..") or safe_name.startswith("."):
        return jsonify({"ok": False, "message": "檔案名稱無效。"}), 400

    os.makedirs(PAID_REPORT_DIR, exist_ok=True)
    destination = os.path.join(PAID_REPORT_DIR, safe_name)
    try:
        uploaded.save(destination)
    except OSError:
        app.logger.exception("Failed to save uploaded paid report")
        return jsonify({"ok": False, "message": "檔案上載失敗，請檢查資料夾權限。"}), 500

    return jsonify({"ok": True, "message": "上載成功。", "name": safe_name})

@app.route("/paid-reports/rename", methods=["POST"])
def rename_paid_report():
                                                        
    payload = request.get_json(silent=True) or {}
    old_name = str(payload.get("name", "")).strip()
    new_name = str(payload.get("new_name", "")).strip()
    safe_old = old_name.replace("\\", "/").rsplit("/", 1)[-1]
    safe_new = new_name.replace("\\", "/").rsplit("/", 1)[-1]
    if (not safe_old or safe_old != old_name or safe_old.startswith(".") or safe_old in (".", "..")
            or not safe_new or safe_new != new_name or safe_new.startswith(".") or safe_new in (".", "..")):
        return jsonify({"ok": False, "message": "檔案名稱無效。"}), 400
    os.makedirs(PAID_REPORT_DIR, exist_ok=True)
    old_path = os.path.join(PAID_REPORT_DIR, safe_old)
    new_path = os.path.join(PAID_REPORT_DIR, safe_new)
    if not os.path.isfile(old_path):
        return jsonify({"ok": False, "message": "找不到要改名的檔案。"}), 404
    if safe_old == safe_new:
        return jsonify({"ok": True, "name": safe_new, "message": "檔案名稱沒有變更。"})
    if os.path.exists(new_path):
        return jsonify({"ok": False, "message": "已有相同名稱的檔案，請使用其他名稱。"}), 409
    try:
        os.rename(old_path, new_path)
    except OSError:
        app.logger.exception("Failed to rename paid report")
        return jsonify({"ok": False, "message": "檔案改名失敗。"}), 500
    return jsonify({"ok": True, "name": safe_new, "message": "檔案改名成功。"})

@app.route("/paid-reports/delete", methods=["POST"])
def delete_paid_report():
                                                        
    payload = request.get_json(silent=True) or {}
    name = str(payload.get("name", "")).strip()
    safe_name = name.replace("\\", "/").rsplit("/", 1)[-1]
    if not safe_name or safe_name != name or safe_name.startswith(".") or safe_name in (".", ".."):
        return jsonify({"ok": False, "message": "檔案名稱無效。"}), 400
    full_path = os.path.join(PAID_REPORT_DIR, safe_name)
    if not os.path.isfile(full_path):
        return jsonify({"ok": False, "message": "找不到要刪除的檔案。"}), 404
    try:
        os.remove(full_path)
    except OSError:
        app.logger.exception("Failed to delete paid report")
        return jsonify({"ok": False, "message": "檔案刪除失敗。"}), 500
    return jsonify({"ok": True, "name": safe_name, "message": "檔案已刪除。"})

@app.route("/paid-reports/download/<path:filename>")
def download_paid_report(filename):
                                                        
    safe_name = os.path.basename(filename)
    if safe_name != filename or safe_name.startswith("."):
        return "找不到檔案。", 404
    full_path = os.path.join(PAID_REPORT_DIR, safe_name)
    if not os.path.isfile(full_path):
        return "找不到檔案。", 404
    return send_file(full_path, as_attachment=True, download_name=safe_name, conditional=True)

@app.route("/paid-reports/preview/<path:filename>")
def preview_paid_report(filename):
                                                                               
    safe_name = os.path.basename(filename)
    if safe_name != filename or safe_name.startswith("."):
        return "找不到檔案。", 404
    full_path = os.path.join(PAID_REPORT_DIR, safe_name)
    if not os.path.isfile(full_path):
        return "找不到檔案。", 404
    ext = os.path.splitext(safe_name)[1].lower()
    mimetypes = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp", ".gif": "image/gif"}
    if ext not in mimetypes:
        return "此檔案格式不支援圖片預覽。", 415
    return send_file(full_path, mimetype=mimetypes[ext], as_attachment=False, download_name=safe_name, conditional=True)

HTML_TEMPLATE = """
<html lang="zh-Hant" data-bs-theme="dark" id="htmlRoot">
<head>
    <meta charset="UTF-8">
    <title>✈️Cathay cargo💰 薪資管理系統</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <style>
        body {
            background-color: #0b140e;
            color: #e0e0e0;
            padding-top: 30px !important; /* 避開 iframe 頂部工具列，避免遮住 ⚙️ cargo 標題 */
            transition: background-color 0.3s ease, color 0.3s ease;
        }
        [data-bs-theme="light"] body {
            background-color: #f4f7f6;
            color: #212529;
        }
        .liquid-glass-card {
            background: rgba(20, 32, 23, 0.65);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.08);
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
            border-radius: 12px;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        [data-bs-theme="light"] .liquid-glass-card {
            background: rgba(255, 255, 255, 0.75);
            border: 1px solid rgba(0, 0, 0, 0.06);
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.06);
        }
        .liquid-glass-card:hover {
            box-shadow: 0 15px 35px rgba(40, 167, 69, 0.15);
        }
        [data-bs-theme="dark"] .night-calendar {
            background: rgba(14, 22, 16, 0.85);
            border: 1px solid rgba(126, 217, 87, 0.2);
            box-shadow: inset 0 0 15px rgba(0, 0, 0, 0.6);
        }
        /* 3D calendar treatment: raised day tiles, perspective and soft depth shadows. */
        .night-calendar {
            border-collapse: separate !important;
            border-spacing: 3px !important;
            perspective: 900px;
            border: 0 !important;
            background: transparent !important;
            box-shadow: none !important;
        }
        .night-calendar thead th {
            border: 0 !important;
            border-radius: 7px;
            box-shadow: 0 3px 0 rgba(0, 0, 0, 0.22), 0 5px 10px rgba(0, 0, 0, 0.12);
            transform: translateZ(4px);
        }
        .night-calendar tbody td {
            border: 1px solid rgba(255, 255, 255, 0.12) !important;
            border-radius: 7px;
            overflow: hidden;
            background-clip: padding-box;
            box-shadow: 0 4px 0 rgba(0, 0, 0, 0.24), 0 7px 12px rgba(0, 0, 0, 0.16), inset 0 1px 0 rgba(255, 255, 255, 0.10);
            transition: transform 180ms ease, box-shadow 180ms ease, filter 180ms ease;
            transform: translateZ(0);
        }
        .night-calendar tbody td:not([style*="background:transparent"]):hover {
            position: relative;
            z-index: 3;
            transform: perspective(600px) translateY(-3px) rotateX(4deg) scale(1.025);
            box-shadow: 0 7px 0 rgba(0, 0, 0, 0.24), 0 12px 18px rgba(0, 0, 0, 0.24), inset 0 1px 0 rgba(255, 255, 255, 0.18);
            filter: brightness(1.08);
        }
        [data-bs-theme="light"] .night-calendar tbody td {
            border-color: rgba(0, 0, 0, 0.12) !important;
            box-shadow: 0 4px 0 rgba(0, 0, 0, 0.10), 0 7px 12px rgba(0, 0, 0, 0.08), inset 0 1px 0 rgba(255, 255, 255, 0.75);
        }
        [data-bs-theme="light"] .night-calendar tbody td:hover {
            box-shadow: 0 7px 0 rgba(0, 0, 0, 0.10), 0 12px 18px rgba(0, 0, 0, 0.13), inset 0 1px 0 rgba(255, 255, 255, 0.8);
        }
        .night-calendar .today-glow-cell,
        .night-calendar td[style*="border: 2px solid #38ef7d"] {
            transform: translateZ(8px);
            box-shadow: 0 0 0 1px rgba(56, 239, 125, 0.4), 0 5px 0 rgba(0, 0, 0, 0.22), 0 0 16px rgba(56, 239, 125, 0.32), inset 0 0 10px rgba(56, 239, 125, 0.22) !important;
        }
        @media (prefers-reduced-motion: reduce) {
            .night-calendar tbody td { transition: none !important; }
            .night-calendar tbody td:hover { transform: none !important; }
        }
        [data-bs-theme="dark"] .night-calendar th {
            background-color: rgba(26, 43, 31, 0.9) !important;
            color: #7ED957 !important;
            border-color: rgba(255, 255, 255, 0.08) !important;
            padding: 4px !important;
            font-size: 0.85rem;
        }
        [data-bs-theme="light"] .night-calendar th {
            padding: 4px !important;
            font-size: 0.85rem;
        }
        [data-bs-theme="dark"] .night-calendar td {
            border-color: rgba(255, 255, 255, 0.06) !important;
            transition: all 0.2s ease;
        }
        /* 極度明顯且強烈的 Today 發光特效 */
        .today-glow-cell {
            border: 3px solid #7ED957 !important;
            box-shadow: 0 0 20px #7ED957, inset 0 0 12px rgba(126, 217, 87, 0.9) !important;
            position: relative;
            z-index: 5;
        }
        /* 記錄控制台永遠至頂 (Sticky Top) 且緊貼無留白與超精簡字體 */
        .sticky-control-panel,
        .sticky-control-panel.liquid-glass-card,
        .sticky-control-panel > .card-header {
            border: 0 !important;
            border-top: 0 !important;
            border-right: 0 !important;
            border-bottom: 0 !important;
            border-left: 0 !important;
            box-shadow: none !important;
            outline: 0 !important;
            background-image: none !important;
        }
        .sticky-control-panel {
            position: sticky;
            top: 30px; /* iframe 頂部工具列約 30px 高，固定控制台時亦保留空間 */
            z-index: 1030;
            font-size: 0.8rem;
            background: transparent !important;
        }
        .sticky-control-panel::before,
        .sticky-control-panel::after {
            border: 0 !important;
            box-shadow: none !important;
        }
        .sticky-control-panel > .card-header {
            border: 0 !important;
            box-shadow: none !important;
        }
        .sticky-control-panel .form-label {
            font-size: 0.75rem;
            margin-bottom: 0.1rem;
        }
        .sticky-control-panel .form-control, .sticky-control-panel .form-select {
            font-size: 0.8rem;
            padding: 0.2rem 0.4rem;
            height: auto;
        }
        .sticky-control-panel .btn {
            font-size: 0.8rem;
            padding: 0.2rem 0.5rem;
        }
        .sticky-control-panel .nav-link {
            font-size: 0.8rem;
            padding: 0.25rem 0.5rem;
        }
        /* ⚙️ cargo Liquid Glass 標題膠囊：必須有可見玻璃底色，而非全透明 */
        #cargoControlHeaderTitle { color: #ffffff !important; }
        [data-bs-theme="light"] #cargoControlHeaderTitle { color: #111111 !important; }
        /* 不使用偽元素遮線，避免在頁面頂部形成黑條／黑邊。 */
        #cargoControlHeaderTitle::before,
        #cargoControlHeaderTitle::after {
            content: none !important;
            display: none !important;
        }
        #cargoControlHeaderTitle .cargo-glass-pill {
            display: inline-flex !important;
            align-items: center;
            justify-content: center;
            gap: 3px;
            padding: 5px 15px;
            border-radius: 999px;
            background: linear-gradient(135deg, rgba(255,255,255,.30), rgba(255,255,255,.10)) !important;
            border: 1px solid rgba(255,255,255,.42);
            box-shadow: 0 4px 14px rgba(0,0,0,.24), inset 0 1px 1px rgba(255,255,255,.48);
            -webkit-backdrop-filter: blur(18px) saturate(180%);
            backdrop-filter: blur(18px) saturate(180%);
            text-shadow: 0 1px 2px rgba(0,0,0,.22);
            white-space: nowrap;
        }
        [data-bs-theme="dark"] #cargoControlHeaderTitle .cargo-glass-pill {
            background: linear-gradient(135deg, rgba(70,85,78,.94), rgba(28,40,33,.88)) !important;
            border-color: rgba(220,255,232,.40);
        }
        [data-bs-theme="light"] #cargoControlHeaderTitle .cargo-glass-pill {
            background: linear-gradient(135deg, rgba(255,255,255,.96), rgba(230,240,235,.86)) !important;
            border-color: rgba(255,255,255,.95);
            box-shadow: 0 4px 14px rgba(20,40,30,.16), inset 0 1px 1px rgba(255,255,255,.95);
            text-shadow: none;
        }

        /* 圖表主題動態切換 */
        .chart-dark, .chart-light { display: none; }
        [data-bs-theme="dark"] .chart-dark { display: block; }
        [data-bs-theme="light"] .chart-light { display: block; }
    
.cargo-calendar-day { cursor: pointer; user-select: none; }
.cargo-calendar-day:focus-visible { outline: 2px solid #38ef7d; outline-offset: -2px; }
.cargo-calendar-popup {
    position: fixed; z-index: 3000; display: flex; gap: 5px; padding: 7px;
    border: 1px solid rgba(255,255,255,.25); border-radius: 12px;
    background: rgba(25,30,38,.96); box-shadow: 0 8px 28px rgba(0,0,0,.35);
    backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
}
.cargo-calendar-popup[hidden] { display: none !important; }
.cargo-calendar-popup button {
    width: 36px; height: 36px; padding: 0; border: 0; border-radius: 9px;
    background: rgba(255,255,255,.12); color: inherit; font-size: 19px;
    display: inline-flex; align-items: center; justify-content: center; cursor: pointer;
}
.cargo-calendar-popup button:hover { background: rgba(56,239,125,.28); }

</style>
<style id="salaryTool3DModeStyles">
/* Unified EXTRA-DEPTH 3D theme shared by cargo and CMHK; disabled in normal 2D mode. */
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
/* In 3D mode the ⚙️ cargo control panel scrolls normally; keep it sticky only in 2D mode. */
body.salary-3d-mode .sticky-control-panel {
  position: static !important;
  top: auto !important;
  z-index: auto !important;
  transform: none !important;
  overflow: visible !important;
  perspective: none !important;
}
/* Keep the ⚙️ cargo title fully visible in 3D mode. The original 1px header
   and absolutely-positioned title are optimized for the 2D sticky layout. */
body.salary-3d-mode .sticky-control-panel > .card-header {
  height: auto !important;
  min-height: 38px !important;
  overflow: visible !important;
  padding: 4px 8px !important;
  display: flex !important;
  justify-content: center !important;
  align-items: center !important;
}
body.salary-3d-mode #cargoControlHeaderTitle {
  position: relative !important;
  top: auto !important;
  left: auto !important;
  right: auto !important;
  width: auto !important;
  max-width: 100% !important;
  margin: 0 !important;
  transform: none !important;
  line-height: 1.35 !important;
  z-index: 2 !important;
  pointer-events: auto !important;
}
body.salary-3d-mode #cargoControlHeaderTitle .cargo-glass-pill {
  pointer-events: auto !important;
}
@media (prefers-reduced-motion: reduce) {
  body.salary-3d-mode *, body.salary-3d-mode *::before, body.salary-3d-mode *::after { transition-duration: .01ms !important; animation-duration: .01ms !important; }
}
/* Phone-only compact heading; iPad and desktop keep the full Cathay cargo title. */
.cargo-title-desktop { display: inline; }
.cargo-title-mobile { display: none; }
@media (max-width: 600px) {
  .cargo-title-desktop { display: none; }
  .cargo-title-mobile { display: inline; }
}
/* cargo add-record form: keep all fields on one line on phones without horizontal scrolling. */
.add-record-wrap { overflow-x: visible; min-width: 0; }
.add-record-form > div { min-width: 0; }
.add-record-date { flex: 1 1 150px; min-width: 120px; }
.add-record-note { flex: 1 1 120px; min-width: 100px; }
.add-record-hours, .add-record-rate { flex: 0 0 90px; }
.add-record-submit { flex: 0 0 110px; }
@media (max-width: 600px) {
  #add-tab { min-width: 0; }
  .add-record-wrap { width: 100%; max-width: 100%; overflow: visible; }
  .add-record-form { width: 100%; min-width: 0 !important; gap: 4px !important; align-items: center !important; }
  .add-record-form > div { min-width: 0 !important; }
  .add-record-form .form-label { display: none !important; }
  .add-record-date { flex: 1.45 1 0 !important; }
  .add-record-note { flex: 1 1 0 !important; }
  .add-record-hours, .add-record-rate { flex: .62 1 0 !important; }
  .add-record-submit { flex: .8 1 0 !important; }
  .add-record-form .form-control,
  .add-record-form .form-select { width: 100% !important; min-width: 0 !important; padding: .25rem .2rem !important; font-size: 11px !important; }
  .add-record-form input[type="date"] { min-width: 0 !important; font-size: 10px !important; }
  .add-record-form input[type="number"] { text-align: center; }
  .add-record-form .btn { padding: .35rem .15rem !important; font-size: 10px !important; overflow: hidden; }
  .add-record-submit-text { display: none !important; }
  #add_custom_note { font-size: 10px !important; }
}
/* 3D mode: keep large cargo/CMHK panels seated in the page instead of
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
<body class="pt-0">
    <div class="card liquid-glass-card mb-4 rounded-0 border-0 sticky-control-panel m-0">
        <div class="card-header bg-transparent d-flex justify-content-between align-items-center rounded-0 px-3 py-0 position-relative" style="height: 1px; min-height: 1px; overflow: visible; border: 0;">
            <div class="d-flex align-items-center gap-2"></div>
            <h6 id="cargoControlHeaderTitle" class="mb-0 text-nowrap position-absolute start-50 translate-middle-x w-100 text-center" style="cursor: pointer; font-size: 0.85rem; z-index: 1040; pointer-events: none; top: -19px; line-height: 1.2;" data-bs-toggle="collapse" data-bs-target="#controlPanelCollapse">
                <span class="cargo-glass-pill" style="pointer-events: auto;">⚙️ cargo <span class="ms-1">▼</span></span>
            </h6>
        </div>
        <div class="collapse" id="controlPanelCollapse">
            <div class="card-body px-3 py-2">
                <ul class="nav nav-tabs mb-2" id="controlTab" role="tablist">
                    <li class="nav-item"><button class="nav-link active py-1 px-2" data-bs-toggle="tab" data-bs-target="#add-tab" type="button">新增記錄</button></li>
                    <li class="nav-item"><button class="nav-link py-1 px-2" data-bs-toggle="tab" data-bs-target="#copy-tab" type="button">複製記錄</button></li>
                    <li class="nav-item"><button class="nav-link py-1 px-2" data-bs-toggle="tab" data-bs-target="#edit-tab" type="button">修改 / 刪除記錄</button></li>
                </ul>
                <div class="tab-content">
                    <div class="tab-pane fade show active" id="add-tab">
                        <!-- 手機版新增記錄固定單行顯示，不產生左右捲動 -->
                        <div class="w-100 add-record-wrap">
                        <form method="POST" class="d-flex align-items-end gap-2 flex-nowrap add-record-form" style="width:100%;">
                            <input type="hidden" name="action" value="add">
                            <div class="add-record-date"><label class="form-label">工作日期</label><input type="date" name="date" class="form-control form-control-sm" aria-label="工作日期" title="工作日期" value="{{ now_date }}" required></div>
                            <div class="add-record-note"><label class="form-label">備註</label><select name="note_type" class="form-select form-select-sm" aria-label="備註" title="備註" id="add_note_type" onchange="toggleCustomNote('add')"><option value="null" selected>null</option><option value="Paid">Paid</option><option value="Off">Off</option><option value="Cut">Cut</option><option value="other">other</option></select><input type="text" name="custom_note" id="add_custom_note" class="form-control form-control-sm mt-1" placeholder="輸入自訂備註" style="display:none;"></div>
                            <div class="add-record-hours"><label class="form-label">工時 (小時)</label><input type="number" name="hours" aria-label="工時 (小時)" title="工時 (小時)" class="form-control form-control-sm" value="10"></div>
                            <div class="add-record-rate"><label class="form-label">時薪 (HKD)</label><input type="number" name="rate" aria-label="時薪 (HKD)" title="時薪 (HKD)" class="form-control form-control-sm" value="75"></div>
                            <div class="add-record-submit"><button type="submit" class="btn btn-success btn-sm w-100 py-1 text-nowrap" aria-label="確認新增" title="確認新增">➕<span class="add-record-submit-text"> 確認新增</span></button></div>
                        </form>
                        </div>
                    </div>
                    <div class="tab-pane fade" id="copy-tab">
                        {% if data %}
                        <form method="POST" class="d-flex align-items-end gap-2 flex-wrap">
                            <input type="hidden" name="action" value="copy">
                            <div style="flex:1 1 180px; min-width:160px;"><label class="form-label">📋 複製原有日期</label><select name="source_date" class="form-select form-select-sm" required>{% for r in data|sort(attribute='date', reverse=True) %}<option value="{{ r.date }}">{{ r.date }}｜{{ r.hours }} 小時｜{{ r.note }}</option>{% endfor %}</select></div>
                            <div style="flex:1 1 150px; min-width:140px;"><label class="form-label">複製開始日期</label><input type="date" name="target_start" class="form-control form-control-sm" value="{{ now_date }}" required></div>
                            <div style="flex:1 1 150px; min-width:140px;"><label class="form-label">複製結束日期（單日可留空）</label><input type="date" name="target_end" class="form-control form-control-sm"></div>
                            <div style="flex:0 0 130px;"><button type="submit" class="btn btn-outline-primary btn-sm w-100 py-1 text-nowrap">📋 批量複製</button></div>
                        </form>
                        {% endif %}
                    </div>
                    <div class="tab-pane fade" id="edit-tab">
                        {% if data %}
                        <!-- 修改 / 刪除記錄改為由左向右橫向緊湊排列 -->
                        <div class="d-flex align-items-end gap-2 flex-wrap">
                            <div style="flex: 1 1 180px; min-width: 160px;">
                                <label class="form-label">選擇記錄 (最新在上)</label>
                                <select class="form-select form-select-sm" id="editSelect" onchange="loadEditData(this)">
                                    {% for idx, r in reversed_indexed_data %}
                                    <option value="{{ idx }}" data-date="{{ r.date }}" data-note="{{ r.note }}" data-hours="{{ r.hours }}" data-rate="{{ r.rate }}">{{ r.date }} | 時薪:{{ r.rate }} | 備註:{{ r.note }}</option>
                                    {% endfor %}
                                </select>
                            </div>
                            <form method="POST" id="updateForm" class="d-flex align-items-end gap-2 flex-wrap m-0 p-0 flex-fill">
                                <input type="hidden" name="action" value="update">
                                <input type="hidden" name="index" id="update_index" value="">
                                <div style="flex: 1 1 130px; min-width: 120px;"><label class="form-label">日期</label><input type="date" name="date" id="edit_date" class="form-control form-control-sm" required></div>
                                <div style="flex: 1 1 130px; min-width: 120px;"><label class="form-label">備註</label><select name="note_type" id="edit_note_type" class="form-select form-select-sm" onchange="toggleCustomNote('edit')"><option value="null" selected>null</option><option value="Paid">Paid</option><option value="Off">Off</option><option value="Cut">Cut</option><option value="other">other</option></select><input type="text" name="custom_note" id="edit_custom_note" class="form-control form-control-sm mt-1" placeholder="輸入自訂備註" style="display:none;"></div>
                                <div style="flex: 0 0 80px;"><label class="form-label">工時</label><input type="number" name="hours" id="edit_hours" class="form-control form-control-sm"></div>
                                <div style="flex: 0 0 80px;"><label class="form-label">時薪</label><input type="number" name="rate" id="edit_rate" class="form-control form-control-sm"></div>
                                <div class="d-flex gap-1" style="flex: 0 0 130px;">
                                    <button type="submit" class="btn btn-primary btn-sm flex-fill py-1 text-nowrap">更新</button>
                                    <button type="button" class="btn btn-danger btn-sm flex-fill py-1 text-nowrap" onclick="submitDelete()">刪除</button>
                                </div>
                            </form>
                            <!-- 獨立隱藏的刪除表單，透過 JavaScript 觸發送出 -->
                            <form method="POST" id="deleteForm" style="display: none;">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="index" id="delete_index" value="">
                            </form>
                        </div>
                        <script>
                            function loadEditData(sel) {
                                var opt = sel.options[sel.selectedIndex];
                                var val = sel.value;
                                document.getElementById('update_index').value = val;
                                document.getElementById('delete_index').value = val;
                                document.getElementById('edit_date').value = opt.getAttribute('data-date');
                                var note = opt.getAttribute('data-note') || '';
                                var known = ['null', 'paid', 'off', 'cut'];
                                var noteType = known.includes(note.toLowerCase()) ? (note.toLowerCase() === 'null' ? 'null' : note.charAt(0).toUpperCase() + note.slice(1).toLowerCase()) : 'other';
                                document.getElementById('edit_note_type').value = noteType;
                                document.getElementById('edit_custom_note').value = noteType === 'other' ? note : '';
                                toggleCustomNote('edit');
                                document.getElementById('edit_hours').value = opt.getAttribute('data-hours');
                                document.getElementById('edit_rate').value = opt.getAttribute('data-rate');
                            }
                            function toggleCustomNote(which) {
                                var select = document.getElementById(which === 'add' ? 'add_note_type' : 'edit_note_type');
                                var input = document.getElementById(which === 'add' ? 'add_custom_note' : 'edit_custom_note');
                                if (!select || !input) return;
                                input.style.display = select.value === 'other' ? '' : 'none';
                                input.required = select.value === 'other';
                            }
                            function submitDelete() {
                                var form = document.getElementById('deleteForm');
                                if (form) form.requestSubmit ? form.requestSubmit() : form.submit();
                            }
                            window.addEventListener('DOMContentLoaded', function() { var sel = document.getElementById('editSelect'); if(sel) loadEditData(sel); });
                        </script>
                        {% else %}
                        <p class="text-muted m-0">尚無記錄</p>
                        {% endif %}
                    </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="container my-4">

        <!-- 月度月曆：以 CMHK 日曆為藍本，最近兩個月直接顯示，其餘月份收起 -->
        <div class="card liquid-glass-card p-3 mb-4">
            <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-2">
                <div class="d-flex align-items-center gap-2">
                    <h5 class="mb-0" style="font-size: 1.05rem;"><span>✈️</span><span class="cargo-title-desktop"> Cathay cargo</span><span class="cargo-title-mobile">cargo</span></h5>
                    {% set speech_month = recent_calendar_months.get(now_date[:7], {}) %}
                    <span class="cargo-speech-controls" style="display:inline-flex;align-items:center;justify-content:center;gap:6px;vertical-align:middle;flex:0 0 auto;white-space:nowrap;">
                        <button type="button" id="cargoMonthlySpeechBtn" class="btn btn-outline-info btn-sm rounded-circle" style="display:inline-flex;align-items:center;justify-content:center;flex:0 0 38px;width:38px;height:38px;line-height:1;padding:0;margin:0;box-sizing:border-box;" title="讀出本月薪資摘要" aria-label="讀出本月薪資摘要" data-hours="{{ speech_month.get('hours', 0) }}" data-month-pay="{{ speech_month.get('total_pay', 0) }}" data-total="{{ total_income }}">🔊</button>
                        <button type="button" id="cargoMonthlySpeechStopBtn" class="btn btn-outline-danger btn-sm rounded-circle" style="display:none;align-items:center;justify-content:center;flex:0 0 38px;width:38px;height:38px;line-height:1;padding:0;margin:0;box-sizing:border-box;" title="停止朗讀" aria-label="停止朗讀">⏹️</button>
                    </span>
                </div>
                <div class="alert alert-warning mb-0 py-1 px-2" style="font-size:1.05rem; line-height:1.35; white-space:nowrap;">
                    總金💰 <b>$ <span id="totalIncome">{{ total_income | money }}</span> HKD</b>
                </div>
            </div>

            {% for m, info in recent_calendar_months.items() %}
                <h4 style="margin: 10px 0 5px 0; font-size: 14px;">📅 【{{ m }}】總工時：<code>{{ info.hours }}h</code> | 月收入：<code>$ {{ info.total_pay | money }}</code></h4>
                <div class="table-responsive cargo-calendar-responsive" style="margin-bottom: 12px;">
                    <table class="table table-bordered text-center align-middle night-calendar m-0" style="table-layout: fixed;">
                        <thead class="table-light">
                            <tr><th>日</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th></tr>
                        </thead>
                        <tbody>
                        {% for week in info.cal %}
                            <tr>
                            {% for d in week %}
                                {% if d == 0 %}
                                    <td style="background:transparent;"></td>
                                {% else %}
                                    {% set day_info = info.days.get(d) %}
                                    {% set is_today = (m == current_year|string ~ '-' ~ '%02d' % current_month and d == current_day) %}
                                    <td class="cargo-calendar-day" data-date="{{ m }}-{{ '%02d' % d }}" data-has-record="{{ '1' if day_info else '0' }}" style="height: 44px; padding: 2px; vertical-align: middle;
                                        {% if day_info and day_info.note == 'paid' %}
                                            background-color: #FFD54F; color: #000000; font-weight: bold;
                                        {% elif day_info and day_info.pay > 0 %}
                                            background-color: rgba(126, 217, 87, 0.25); color: #7ED957; font-weight: bold;
                                        {% elif day_info and day_info.note == 'off' %}
                                            background-color: rgba(110, 168, 254, 0.25); color: #6ea8fe; font-weight: bold;
                                        {% elif day_info and day_info.note == 'cut' %}
                                            background-color: rgba(234, 134, 143, 0.25); color: #ea868f; font-weight: bold;
                                        {% endif %}
                                        {% if is_today %}
                                            border: 2px solid #38ef7d !important; box-shadow: inset 0 0 8px rgba(56, 239, 125, 0.6);
                                        {% endif %}
                                    ">
                                        <b style="{% if day_info and day_info.note == 'paid' %}color: {% if day_info.hours == 0 %}#ff0000{% else %}#0000ff{% endif %};{% elif is_today %}color: #38ef7d; font-size: 15px; text-shadow: 0 0 5px rgba(56, 239, 125, 0.8);{% endif %}">{{ d }}</b>
                                        {% if day_info and day_info.note == 'paid' %}
                                            <br><span style="color:#000000; font-size:11px;">Paid</span>
                                        {% elif day_info and day_info.pay > 0 %}
                                            <br><span style="color:#ffffff; font-size:11px;">${{ day_info.pay | money }}</span>
                                        {% elif day_info and day_info.note == 'off' %}
                                            <br><span style="color:#E57373; font-size:11px;">off</span>
                                        {% elif day_info and day_info.note == 'cut' %}
                                            <br><span style="color:#ea868f; font-size:11px;">cut</span>
                                        {% endif %}
                                    </td>
                                {% endif %}
                            {% endfor %}
                            </tr>
                        {% endfor %}
                        </tbody>
                    </table>
                </div>
            {% endfor %}

            {% if older_calendar_months %}
                <details style="margin-top: 15px; border-top: 1px dashed rgba(200, 230, 201, 0.4); padding-top: 10px;">
                    <summary style="font-size: 14px; cursor: pointer;">📂 展開所有以往月份的日曆 (共 {{ older_calendar_months|length }} 個月份)</summary>
                    <div style="margin-top: 10px;">
                        {% for m, info in older_calendar_months.items() %}
                            <h4 style="margin: 15px 0 5px 0; font-size: 14px;">📅 【{{ m }}】總工時：<code>{{ info.hours }}h</code> | 月收入：<code>$ {{ info.total_pay | money }}</code></h4>
                            <div class="table-responsive cargo-calendar-responsive" style="margin-bottom: 12px;">
                                <table class="table table-bordered text-center align-middle night-calendar m-0" style="table-layout: fixed;">
                                    <thead class="table-light">
                                        <tr><th>日</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th></tr>
                                    </thead>
                                    <tbody>
                                    {% for week in info.cal %}
                                        <tr>
                                        {% for d in week %}
                                            {% if d == 0 %}
                                                <td style="background:transparent;"></td>
                                            {% else %}
                                                {% set day_info = info.days.get(d) %}
                                                {% set is_today = (m == current_year|string ~ '-' ~ '%02d' % current_month and d == current_day) %}
                                                <td class="cargo-calendar-day" data-date="{{ m }}-{{ '%02d' % d }}" data-has-record="{{ '1' if day_info else '0' }}" style="height: 44px; padding: 2px; vertical-align: middle;
                                                    {% if day_info and day_info.note == 'paid' %}
                                                        background-color: #FFD54F; color: #000000; font-weight: bold;
                                                    {% elif day_info and day_info.pay > 0 %}
                                                        background-color: rgba(126, 217, 87, 0.25); color: #7ED957; font-weight: bold;
                                                    {% elif day_info and day_info.note == 'off' %}
                                                        background-color: rgba(110, 168, 254, 0.25); color: #6ea8fe; font-weight: bold;
                                                    {% elif day_info and day_info.note == 'cut' %}
                                                        background-color: rgba(234, 134, 143, 0.25); color: #ea868f; font-weight: bold;
                                                    {% endif %}
                                                ">
                                                    <b style="{% if day_info and day_info.note == 'paid' %}color:{% if day_info.hours == 0 %}#ff0000{% else %}#0000ff{% endif %};{% endif %}">{{ d }}</b>
                                                    {% if day_info and day_info.note == 'paid' %}
                                                        <br><span style="color:#000000; font-size:11px;">Paid</span>
                                                    {% elif day_info and day_info.pay > 0 %}
                                                        <br><span style="color:#ffffff; font-size:11px;">${{ day_info.pay | money }}</span>
                                                    {% elif day_info and day_info.note == 'off' %}
                                                        <br><span style="color:#E57373; font-size:11px;">off</span>
                                                    {% elif day_info and day_info.note == 'cut' %}
                                                        <br><span style="color:#ea868f; font-size:11px;">cut</span>
                                                    {% endif %}
                                                </td>
                                            {% endif %}
                                        {% endfor %}
                                        </tr>
                                    {% endfor %}
                                    </tbody>
                                </table>
                            </div>
                        {% endfor %}
                    </div>
                </details>
            {% endif %}
        </div>

        <!-- 本期及上一期出糧週期明細 -->
        <div class="card liquid-glass-card p-4 mb-4">
            <h4>📊 本期出糧週期明細</h4>

            <div class="row text-center mt-3 g-3">
                <!-- 本期 -->
                <div class="col-md-6">
                    <div class="p-3 rounded border-start border-success border-4 h-100">
                        <h5 class="mb-3">🟢 本期（{{ start.strftime('%m/%d') }} ~ {{ end.strftime('%m/%d') }}）</h5>
                        <div class="row g-2">
                            <div class="col-6"><div class="p-2 bg-transparent rounded border"><span>⏱️ 總工時</span><br><b>{{ total_h }} 小時</b></div></div>
                            <div class="col-6"><div class="p-2 bg-transparent rounded border"><span>📅 工作日數</span><br><b>{{ work_days }} 天</b></div></div>
                            <div class="col-12"><div class="p-2 bg-transparent rounded border"><span>💰 累計薪水</span><br><b>$ {{ total_t | money }} HKD</b></div></div>
                        </div>
                    </div>
                </div>

                <!-- 上一期 -->
                <div class="col-md-6">
                    <div class="p-3 rounded border-start border-secondary border-4 h-100">
                        <h5 class="mb-3">🔵 上一期（{{ previous_start.strftime('%m/%d') }} ~ {{ previous_end.strftime('%m/%d') }}）</h5>
                        <div class="row g-2">
                            <div class="col-6"><div class="p-2 bg-transparent rounded border"><span>⏱️ 總工時</span><br><b>{{ previous_total_h }} 小時</b></div></div>
                            <div class="col-6"><div class="p-2 bg-transparent rounded border"><span>📅 工作日數</span><br><b>{{ previous_work_days }} 天</b></div></div>
                            <div class="col-12"><div class="p-2 bg-transparent rounded border"><span>💰 累計薪水</span><br><b>$ {{ previous_total_t | money }} HKD</b></div></div>
                        </div>
                    </div>
                </div>

                <!-- 匯出 -->
                <div class="col-12 mt-3">
                    <div class="d-flex flex-column gap-2 w-100">
                        <div class="d-flex gap-2">
                            <a href="export/excel?period=current" class="btn btn-success flex-fill py-2">📥 Excel｜匯出今期</a>
                            <a href="export/excel?period=previous" class="btn btn-outline-success flex-fill py-2">📥 Excel｜匯出上一期</a>
                        </div>
                        <div class="d-flex gap-2">
                            <a href="export/pdf?period=current" class="btn btn-primary flex-fill py-2">📥 PDF｜匯出今期</a>
                            <a href="export/pdf?period=previous" class="btn btn-outline-primary flex-fill py-2">📥 PDF｜匯出上一期</a>
                        </div>
                        <div class="d-flex gap-2">
                            <a href="export/jpg?period=current" class="btn btn-warning flex-fill py-2 text-dark">📥 JPG｜匯出今期</a>
                            <a href="export/jpg?period=previous" class="btn btn-outline-warning flex-fill py-2">📥 JPG｜匯出上一期</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div id="cargoCalendarActionPopup" class="cargo-calendar-popup" role="dialog" aria-label="日曆日期操作" hidden>
            <button type="button" data-cargo-action="add" title="新增" aria-label="新增">➕</button>
            <button type="button" data-cargo-action="copy" data-requires-record="1" title="複製" aria-label="複製">📋</button>
            <button type="button" data-cargo-action="edit" data-requires-record="1" title="修改" aria-label="修改">✏️</button>
            <button type="button" data-cargo-action="delete" data-requires-record="1" title="刪除" aria-label="刪除">🗑️</button>
        </div>
        <!-- 每日薪資明細表 -->
        <div class="card liquid-glass-card p-4 mb-4">
            <h4>📋 每日薪資明細表</h4>
            <div class="table-responsive mt-3">
                <table class="table table-striped table-hover align-middle">
                    <thead class="table-success"><tr><th>日期</th><th>星期</th><th>工時</th><th>當日薪資</th><th>備註</th></tr></thead>
                    <tbody>
                        {% for d, val in records_dict.items() %}
                        {% if val[0] > 0 and val[1] > 0 %}
                        <tr>
                            <td>{{ d.strftime('%m/%d') }}</td>
                            <td>{{ WEEKDAY_CN[d.weekday()] }}</td>
                            <td>{{ val[0] }} hrs</td>
                            <td>$ {{ val[1] | money }}</td>
                            <td>{% if val[2] and val[2]|string|lower|trim != 'null' %}{{ val[2] }}{% endif %}</td>
                        </tr>
                        {% endif %}
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>

        <!-- 檔案：瀏覽 export 內的歷史匯出檔案 -->
        <div class="card liquid-glass-card p-4 mb-4" id="paidReportsCard">
            <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-3">
                <div class="d-flex align-items-center gap-2">
                    <h4 class="mb-0">🗂️ 檔案</h4>
                    <input type="file" id="paidReportUploadInput" class="d-none" aria-label="選擇要上載的檔案">
                    <button type="button" class="btn btn-success btn-sm" id="paidReportUploadBtn" title="上載檔案" aria-label="上載檔案">⬆️</button>
                    <button type="button" class="btn btn-outline-primary btn-sm" id="paidReportDownloadBtn" title="下載選定檔案" aria-label="下載選定檔案">⬇️</button>
                    <button type="button" class="btn btn-outline-warning btn-sm" id="paidReportRenameBtn" title="重新命名選定檔案" aria-label="重新命名選定檔案">✏️</button>
                    <button type="button" class="btn btn-outline-danger btn-sm" id="paidReportDeleteBtn" title="刪除選定檔案" aria-label="刪除選定檔案">🗑️</button>
                </div>
                <div class="d-flex align-items-center gap-2">
                    <select id="paidReportSort" class="form-select form-select-sm" aria-label="按修改日期排序" title="按修改日期排序" style="width:auto; min-width:145px;">
                        <option value="newest">最新修改在上</option>
                        <option value="oldest">最舊修改在上</option>
                    </select>
                    <button type="button" class="btn btn-success btn-sm paid-whatsapp-btn" id="paidReportWhatsAppBtn" title="以 WhatsApp 發送選定 JPG" aria-label="以 WhatsApp 發送選定 JPG">
                        <svg viewBox="0 0 32 32" width="20" height="20" aria-hidden="true" focusable="false">
                            <circle cx="16" cy="15" r="12.5" fill="currentColor"/>
                            <path fill="#198754" d="M10.2 8.7c-.4.5-.9 1.3-.9 2.4 0 1.1.8 2.2 1 2.5.1.2 1.7 2.8 4.3 4.1 2.1 1.1 2.6.9 3.1.8.5-.1 1.6-.7 1.8-1.3.2-.6.2-1.1.1-1.2-.1-.1-.4-.2-.8-.4l-1.8-.9c-.2-.1-.4-.2-.6.2l-.8 1c-.2.2-.3.2-.6.1-.3-.1-1.1-.4-2.1-1.3-.8-.7-1.3-1.5-1.5-1.8-.2-.3 0-.4.1-.6l.5-.6c.2-.2.2-.3.3-.5.1-.2 0-.4 0-.5l-.8-1.9c-.2-.5-.4-.4-.6-.4h-.5c-.2 0-.5.1-.7.4Z"/>
                            <path fill="currentColor" d="m5.2 23.4-1.7 6 6.2-1.6z"/>
                        </svg>
                    </button>
                    <button type="button" class="btn btn-outline-success btn-sm" id="refreshPaidReports" title="重新整理" aria-label="重新整理">↻</button>
                </div>
            </div>
            <div class="paid-explorer">
                <div class="paid-file-pane">
                    <div class="paid-pane-heading">📁 檔案清單 <span id="paidReportCount" class="badge text-bg-secondary">0</span></div>
                    <div id="paidReportFileList" class="paid-file-list" aria-label="檔案清單">
                        <div class="text-secondary small p-3">正在載入檔案…</div>
                    </div>
                </div>
            </div>
        </div>

        <div id="paidImageFullscreen" role="dialog" aria-modal="true" aria-label="全屏圖片預覽">
            <button type="button" id="paidImageFullscreenClose" aria-label="退出全屏預覽" title="退出全屏預覽">×</button>
            <img id="paidImageFullscreenImg" alt="檔案全屏預覽">
        </div>

        <!-- 數據統計與圖表分析：按需要才載入 -->
        <div class="card liquid-glass-card p-4 mb-4" style="overflow:hidden;">
            <div class="d-flex justify-content-between align-items-center flex-wrap gap-2">
                <h4 class="mb-0">📊 數據統計與圖表分析</h4>
                <button type="button" id="loadChartsBtn" class="btn btn-success btn-sm" onclick="toggleCharts()">📊 查看圖表</button>
            </div>
            <div id="chartsLoading" class="text-center text-muted py-4" style="display:none;">⏳ 正在載入圖表分析…</div>
            <div id="chartsContent" style="display:none; width:100%; max-width:100%; overflow:hidden;">
                <div class="row mt-4 g-3" style="margin-left:0; margin-right:0;">
                    <div class="col-12 col-lg-4 text-center" style="min-width:0;">
                        <h6 class="text-success mb-2">📋 歷史所有資料出勤比例</h6>
                        <div class="chart-box" id="chartAll"></div>
                    </div>
                    <div class="col-12 col-lg-4 text-center" style="min-width:0;">
                        <h6 class="text-success mb-2">📅 本月出勤狀態比例 ({{ selected_month }})</h6>
                        <div class="chart-box" id="chartCurr"><p class="text-muted mt-5">本月尚無資料</p></div>
                    </div>
                    <div class="col-12 col-lg-4 text-center" style="min-width:0;">
                        <h6 class="text-success mb-2">⏮️ 上一個月出勤狀態比例</h6>
                        <div class="chart-box" id="chartPrev"><p class="text-muted mt-5">上一個月尚無資料</p></div>
                    </div>
                </div>
            </div>
        </div>
        <br><br>
        <p style="opacity: 0.7; margin: 4px 0 0 0; font-size: 12px;">Copyright &copy; 2026 คาร์ลอส. All Rights Reserved. | 系統時間: {{ today_str }}</p>
    </div>

    <style>
        .paid-explorer { display:grid; grid-template-columns:minmax(0, 1fr); min-height:260px; height:min(65vh, 620px); }
        .paid-file-pane { min-width:0; min-height:0; border:1px solid var(--bs-border-color); border-radius:10px; overflow:hidden; background:rgba(127,127,127,.045); display:flex; flex-direction:column; }
        .paid-pane-heading { display:flex; justify-content:space-between; align-items:center; gap:8px; padding:10px 12px; font-size:13px; font-weight:700; border-bottom:1px solid var(--bs-border-color); flex:0 0 auto; }
        .paid-file-list { flex:1; min-height:0; overflow:auto; }
        .paid-file-row { width:100%; display:flex; align-items:center; gap:8px; padding:8px 10px; border:0; border-bottom:1px solid var(--bs-border-color); background:transparent; color:var(--bs-body-color); text-align:left; cursor:pointer; font-size:12px; line-height:1.35; }
        .paid-file-row:hover { background:rgba(25,135,84,.10); }
        .paid-file-row.active { background:rgba(25,135,84,.20); box-shadow:inset 3px 0 #198754; }
        .paid-file-icon { flex:0 0 20px; font-size:16px; }
        .paid-file-info { min-width:0; flex:1; }
        .paid-file-name { display:block; overflow-wrap:anywhere; font-weight:650; }
        .paid-file-meta { display:block; margin-top:3px; opacity:.7; font-size:10px; }
        #paidImageFullscreen { display:none; position:fixed; inset:0; z-index:2147483647; background:rgba(0,0,0,.96); align-items:center; justify-content:center; padding:24px; isolation:isolate; transform:none !important; transform-style:flat !important; perspective:none !important; }
        #paidImageFullscreen.open { display:flex; }
        #paidImageFullscreen img { display:block; width:auto; height:auto; max-width:100%; max-height:100%; object-fit:contain; transform:none !important; filter:none !important; }
        #paidImageFullscreenClose { position:absolute; top:12px; left:12px; z-index:2147483647; width:42px; height:42px; border:1px solid #ffffff88; border-radius:50%; background:#151515e8; color:#fff; font-size:27px; line-height:1; cursor:pointer; transform:none !important; }
        /* Keep the full-screen preview above cargo's 3D card/transform stacking layers. */
        body.salary-3d-mode #paidImageFullscreen { position:fixed !important; inset:0 !important; z-index:2147483647 !important; transform:none !important; transform-style:flat !important; perspective:none !important; }
        body.paid-preview-fullscreen-open { perspective:none !important; }
        @media (max-width: 767.98px) { .paid-explorer { grid-template-columns:minmax(0, 1fr); height:min(65vh, 620px); min-height:220px; } .paid-file-pane { min-height:0; } .paid-file-list { max-height:none; min-height:0; } }
        /* 圖表固定在「數據統計與圖表分析」卡片內，不容許內容向外撐開。 */
        #chartsContent, #chartsContent .row, #chartsContent [class*="col-"] {
            max-width: 100%;
            min-width: 0;
        }
        #chartsContent .chart-box {
            width: 100%;
            max-width: 100%;
            min-width: 0;
            height: 320px;
            overflow: hidden;
            position: relative;
            border: 1px solid rgba(0, 245, 212, .24);
            border-radius: 14px;
            background: radial-gradient(ellipse at 50% 42%, rgba(0, 166, 251, .12), rgba(3, 8, 24, .08) 68%, transparent 100%);
            box-shadow: inset 0 0 22px rgba(0, 245, 212, .07), 0 0 14px rgba(0, 166, 251, .08);
        }
        /* Faux 3D extrusion, neon rim and hover lift for the pie slices. */
        #chartsContent .chart-box .surface,
        #chartsContent .chart-box .slices path,
        #chartsContent .chart-box path.slice {
            filter: drop-shadow(0 5px 0 rgba(0, 0, 0, .48)) drop-shadow(0 8px 5px rgba(0, 245, 212, .18));
            transition: filter .22s ease, opacity .22s ease;
        }
        #chartsContent .chart-box .slices path:hover,
        #chartsContent .chart-box path.slice:hover {
            filter: drop-shadow(0 7px 0 rgba(0, 0, 0, .55)) drop-shadow(0 0 10px rgba(0, 245, 212, .58));
            opacity: .96;
        }
        #chartsContent .chart-box .main-svg { overflow: visible; }
        #chartsContent .chart-box .plotly,
        #chartsContent .chart-box .js-plotly-plot,
        #chartsContent .chart-box .plot-container,
        #chartsContent .chart-box .svg-container {
            width: 100% !important;
            max-width: 100% !important;
        }
        @media (max-width: 991.98px) {
            #chartsContent .chart-box { height: 300px; }
        }
    </style>

    <script>
        (function initPaidReportExplorer() {
            const listEl = document.getElementById('paidReportFileList');
            const sortEl = document.getElementById('paidReportSort');
            const countEl = document.getElementById('paidReportCount');
            const refreshBtn = document.getElementById('refreshPaidReports');
            const uploadBtn = document.getElementById('paidReportUploadBtn');
            const downloadBtn = document.getElementById('paidReportDownloadBtn');
            const renameBtn = document.getElementById('paidReportRenameBtn');
            const deleteBtn = document.getElementById('paidReportDeleteBtn');
            const whatsappBtn = document.getElementById('paidReportWhatsAppBtn');
            const uploadInput = document.getElementById('paidReportUploadInput');
            const fullscreen = document.getElementById('paidImageFullscreen');
            // Keep the overlay outside transformed cards/containers so 3D mode cannot cover it.
            if (fullscreen && fullscreen.parentElement !== document.body) document.body.appendChild(fullscreen);
            const fullscreenImg = document.getElementById('paidImageFullscreenImg');
            const closeBtn = document.getElementById('paidImageFullscreenClose');
            let reportFiles = [];
            let selectedName = '';

            function escapeHtml(value) {
                return String(value).replace(/[&<>"']/g, function(ch) {
                    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch];
                });
            }
            function renderFileList() {
                const newestFirst = sortEl.value !== 'oldest';
                const sorted = reportFiles.slice().sort((a, b) => newestFirst ? b.modified - a.modified : a.modified - b.modified);
                countEl.textContent = String(sorted.length);
                if (!sorted.length) {
                    listEl.innerHTML = '<div class="text-secondary small p-3">export 內暫時沒有檔案。匯出 JPG 後會自動儲存到這裡。</div>';
                    return;
                }
                listEl.innerHTML = sorted.map(file => `
                    <button type="button" class="paid-file-row ${file.name === selectedName ? 'active' : ''}" data-name="${escapeHtml(file.name)}">
                        <span class="paid-file-icon">${file.is_image ? '🖼️' : '📄'}</span>
                        <span class="paid-file-info"><span class="paid-file-name">${escapeHtml(file.name)}</span><span class="paid-file-meta">${escapeHtml(file.modified_text)} · ${escapeHtml(file.size_text)}</span></span>
                    </button>`).join('');
                listEl.querySelectorAll('.paid-file-row').forEach(btn => {
                    btn.addEventListener('click', () => selectFile(btn.dataset.name));
                    btn.addEventListener('dblclick', () => {
                        const file = reportFiles.find(item => item.name === btn.dataset.name);
                        if (!file) return;
                        selectFile(file.name);
                        if (file.is_image) {
                            const url = file.preview_url + '?v=' + encodeURIComponent(file.modified);
                            openFullscreen(url, file.name);
                        }
                    });
                });
            }
            function selectFile(name) {
                const file = reportFiles.find(item => item.name === name);
                if (!file) return;
                selectedName = name;
                listEl.querySelectorAll('.paid-file-row').forEach(btn => {
                    btn.classList.toggle('active', btn.dataset.name === selectedName);
                });
            }
            function openFullscreen(url, name) {
                fullscreenImg.src = url;
                fullscreenImg.alt = name || '檔案全屏預覽';
                fullscreen.classList.add('open');
                document.body.classList.add('paid-preview-fullscreen-open');
                document.body.style.overflow = 'hidden';
                closeBtn.focus();
            }
            function closeFullscreen() {
                fullscreen.classList.remove('open');
                document.body.classList.remove('paid-preview-fullscreen-open');
                fullscreenImg.removeAttribute('src');
                document.body.style.overflow = '';
            }
            async function loadFiles() {
                listEl.innerHTML = '<div class="text-secondary small p-3">正在載入檔案…</div>';
                try {
                    const response = await fetch({{ url_for('list_paid_reports') | tojson }}, {cache:'no-store'});
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const result = await response.json();
                    reportFiles = Array.isArray(result.files) ? result.files : [];
                    renderFileList();
                    if (selectedName && reportFiles.some(file => file.name === selectedName)) renderFileList();
                    else if (!reportFiles.length) selectedName = '';
                } catch (error) {
                    listEl.innerHTML = '<div class="text-danger small p-3">無法載入 export 檔案清單。</div>';
                }
            }
            function selectedFileExists() {
                return !!selectedName && reportFiles.some(file => file.name === selectedName);
            }
            sortEl.addEventListener('change', renderFileList);
            refreshBtn.addEventListener('click', loadFiles);
            uploadBtn.addEventListener('click', () => uploadInput.click());
            downloadBtn.addEventListener('click', () => {
                if (!selectedFileExists()) {
                    window.alert('請先在左邊檔案清單選定要下載的檔案。');
                    return;
                }
                window.location.href = {{ url_for('download_paid_report', filename='__FILENAME__') | tojson }}.replace('__FILENAME__', encodeURIComponent(selectedName));
            });
            whatsappBtn.addEventListener('click', async () => {
                if (!selectedFileExists()) {
                    window.alert('請先在左邊檔案清單選定要發送的 JPG 檔案。');
                    return;
                }
                const selectedFile = reportFiles.find(item => item.name === selectedName);
                if (!selectedFile || !/\\.jpe?g$/i.test(selectedFile.name)) {
                    window.alert('WhatsApp 分享功能只支援 JPG / JPEG 檔案。');
                    return;
                }
                if (!navigator.share || !navigator.canShare) {
                    window.alert('此瀏覽器不支援直接分享檔案，請先下載 JPG，再於 WhatsApp 附加原檔發送。');
                    return;
                }
                whatsappBtn.disabled = true;
                try {
                    const response = await fetch(selectedFile.preview_url, {cache:'no-store'});
                    if (!response.ok) throw new Error('無法讀取選定的 JPG 檔案。');
                    const blob = await response.blob();
                    const originalFile = new File([blob], selectedFile.name, {type:'image/jpeg', lastModified:Date.now()});
                    if (!navigator.canShare({files:[originalFile]})) {
                        throw new Error('此裝置或瀏覽器不支援分享 JPG 原檔，請先下載檔案再於 WhatsApp 附加。');
                    }
                    await navigator.share({files:[originalFile], title:selectedFile.name});
                } catch (error) {
                    if (error && error.name === 'AbortError') return;
                    window.alert(error.message || 'WhatsApp 分享失敗，請稍後再試。');
                } finally {
                    whatsappBtn.disabled = false;
                }
            });
            renameBtn.addEventListener('click', async () => {
                if (!selectedFileExists()) {
                    window.alert('請先在左邊檔案清單選定要改名的檔案。');
                    return;
                }
                const current = selectedName;
                const proposed = window.prompt('請輸入新的檔案名稱：', current);
                if (proposed === null) return;
                const newName = proposed.trim();
                if (!newName) {
                    window.alert('檔案名稱不可留空。');
                    return;
                }
                renameBtn.disabled = true;
                try {
                    const response = await fetch({{ url_for('rename_paid_report') | tojson }}, {
                        method: 'POST', headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({name: current, new_name: newName}), cache: 'no-store'
                    });
                    const result = await response.json();
                    if (!response.ok || !result.ok) throw new Error(result.message || '檔案改名失敗。');
                    selectedName = result.name;
                    await loadFiles();
                    if (reportFiles.some(item => item.name === selectedName)) selectFile(selectedName);
                    window.alert(result.message || '檔案改名成功。');
                } catch (error) {
                    window.alert(error.message || '檔案改名失敗。');
                } finally {
                    renameBtn.disabled = false;
                }
            });
            deleteBtn.addEventListener('click', async () => {
                if (!selectedFileExists()) {
                    window.alert('請先在左邊檔案清單選定要刪除的檔案。');
                    return;
                }
                const current = selectedName;
                if (!window.confirm('確定要刪除檔案「' + current + '」？此操作無法復原。')) return;
                deleteBtn.disabled = true;
                try {
                    const response = await fetch({{ url_for('delete_paid_report') | tojson }}, {
                        method: 'POST', headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({name: current}), cache: 'no-store'
                    });
                    const result = await response.json();
                    if (!response.ok || !result.ok) throw new Error(result.message || '檔案刪除失敗。');
                    selectedName = '';
                    await loadFiles();
                    window.alert(result.message || '檔案已刪除。');
                } catch (error) {
                    window.alert(error.message || '檔案刪除失敗。');
                } finally {
                    deleteBtn.disabled = false;
                }
            });
            uploadInput.addEventListener('change', async () => {
                const file = uploadInput.files && uploadInput.files[0];
                if (!file) return;
                const formData = new FormData();
                formData.append('file', file, file.name);
                uploadBtn.disabled = true;
                const originalLabel = uploadBtn.textContent;
                uploadBtn.textContent = '上載中…';
                try {
                    const response = await fetch({{ url_for('upload_paid_report') | tojson }}, {
                        method: 'POST', body: formData, cache: 'no-store'
                    });
                    const result = await response.json();
                    if (!response.ok || !result.ok) throw new Error(result.message || '上載失敗。');
                    selectedName = result.name;
                    await loadFiles();
                    if (reportFiles.some(item => item.name === selectedName)) selectFile(selectedName);
                    window.alert('上載成功：' + result.name);
                } catch (error) {
                    window.alert(error.message || '檔案上載失敗，請稍後再試。');
                } finally {
                    uploadBtn.disabled = false;
                    uploadBtn.textContent = originalLabel;
                    uploadInput.value = '';
                }
            });
            closeBtn.addEventListener('click', closeFullscreen);
            document.addEventListener('keydown', event => {
                if (event.key === 'Escape' && fullscreen.classList.contains('open')) closeFullscreen();
            });
            fullscreen.addEventListener('click', event => { if (event.target === fullscreen) closeFullscreen(); });
            loadFiles();
        })();

        let chartsLoaded = false;
        let chartsVisible = false;
        let plotlyLoading = null;

        // 限制 Plotly 只使用各自欄位內的寬度，避免 SVG / canvas 撐出外框。
        function plotlyConfig() {
            return {responsive:true, displayModeBar:false, scrollZoom:false};
        }

        function fitCharts() {
            if (!window.Plotly) return;
            ['chartAll', 'chartCurr', 'chartPrev'].forEach(function(id) {
                const el = document.getElementById(id);
                if (el && el.data) {
                    try { Plotly.Plots.resize(el); } catch (e) {}
                }
            });
        }

        function loadPlotly() {
            if (window.Plotly) return Promise.resolve();
            if (plotlyLoading) return plotlyLoading;
            plotlyLoading = new Promise((resolve, reject) => {
                const script = document.createElement('script');
                script.src = 'https://cdn.plot.ly/plotly-2.35.2.min.js';
                script.onload = resolve;
                script.onerror = () => reject(new Error('Plotly 載入失敗'));
                document.head.appendChild(script);
            });
            return plotlyLoading;
        }

        function currentChartTheme() {
            return document.getElementById('htmlRoot').getAttribute('data-bs-theme') === 'light' ? 'light' : 'dark';
        }

        async function loadCharts() {
            const loading = document.getElementById('chartsLoading');
            const content = document.getElementById('chartsContent');
            const btn = document.getElementById('loadChartsBtn');
            loading.style.display = 'block';
            content.style.display = 'none';
            btn.disabled = true;

            try {
                await loadPlotly();
                // cargo 可能掛載在 /cargo 下，不能寫死成網站根目錄 /charts。
                const chartsUrl = {{ url_for('charts') | tojson }};
                const params = new URLSearchParams({
                    month: {{ selected_month | tojson }},
                    theme: currentChartTheme(),
                    lang: (document.documentElement.lang || 'zh-Hant')
                });
                const response = await fetch(chartsUrl + '?' + params.toString(), { cache: 'no-store' });
                if (!response.ok) throw new Error('HTTP ' + response.status);
                const result = await response.json();

                const chartOptions = plotlyConfig();
                Plotly.newPlot('chartAll', JSON.parse(result.charts[0]).data, JSON.parse(result.charts[0]).layout, chartOptions);
                if (result.charts[1]) {
                    const curr = JSON.parse(result.charts[1]);
                    Plotly.newPlot('chartCurr', curr.data, curr.layout, chartOptions);
                }
                if (result.charts[2]) {
                    const prev = JSON.parse(result.charts[2]);
                    Plotly.newPlot('chartPrev', prev.data, prev.layout, chartOptions);
                }
                document.getElementById('totalIncome').textContent = Number(result.total_income || 0).toLocaleString();
                chartsLoaded = true;
                chartsVisible = true;
                loading.style.display = 'none';
                content.style.display = 'block';
                requestAnimationFrame(function() {
                    fitCharts();
                    setTimeout(fitCharts, 80);
                });
                btn.disabled = false;
                btn.textContent = '📉 收起圖表';
            } catch (error) {
                loading.textContent = '⚠️ 圖表載入失敗，請稍後再試。';
                btn.disabled = false;
            }
        }

        function toggleCharts() {
            const content = document.getElementById('chartsContent');
            const btn = document.getElementById('loadChartsBtn');
            if (chartsVisible) {
                content.style.display = 'none';
                chartsVisible = false;
                btn.textContent = '📊 查看圖表';
                return;
            }
            if (chartsLoaded) {
                content.style.display = 'block';
                chartsVisible = true;
                btn.textContent = '📉 收起圖表';
                return;
            }
            loadCharts();
        }

        function applyPortalTheme(theme) {
            const root = document.getElementById('htmlRoot');
            if (!root) return;
            root.setAttribute('data-bs-theme', theme === 'light' ? 'light' : 'dark');
            // 讓 Plotly 圖表在切換主題後自動重新計算尺寸以防跑版
            window.dispatchEvent(new Event('resize'));
        }
        window.addEventListener('message', function(event) {
            if (event.data && event.data.type === 'salary-tool-theme') {
                applyPortalTheme(event.data.theme);
            }
        });
        window.parent.postMessage({type: 'salary-tool-theme-ready'}, '*');
    </script>
<script>
(async function(){
 const maps={
 en:{'新增':'Add','修改':'Edit','刪除':'Delete','複製':'Copy','儲存':'Save','取消':'Cancel','搜尋':'Search','清除':'Clear','完成':'Completed','未完成':'Incomplete','Paid':'Paid','休息':'Off','其他':'Other','工時':'Hours','時薪':'Hourly rate','收入':'Income','總收入':'Total income','本月':'This month','上月':'Last month','今日':'Today','日期':'Date','備註':'Notes','選擇月份':'Select month','匯出':'Export','下載':'Download','服務統計':'Service statistics','圖表':'Charts','夜間模式':'Dark mode','日間模式':'Light mode','確定':'Confirm','返回':'Back','複製到日期':'Copy to date','開始日期':'Start date','結束日期':'End date','月份':'Month','年':'Year','月':'Month','日':'Day','時':'Hours','分鐘':'Minutes','新增記錄':'Add record','修改記錄':'Edit record','刪除記錄':'Delete record','選擇日期':'Select date','開始':'Start','結束':'End','上一月':'Previous month','下一月':'Next month','上一期':'Previous period','今期':'Current period','匯出今期':'Export current period','匯出上一期':'Export previous period','目前無資料可供匯出。':'No data available to export.','歷史以來總收入累計金額：':'All-time cumulative income:','系統時間：':'System time:','圖表載入失敗，請稍後再試':'Chart loading failed. Please try again later.','新增成功':'Added successfully','儲存成功':'Saved successfully','確定刪除？':'Are you sure you want to delete this?','備註（可選）':'Notes (optional)','自訂備註':'Custom note','Paid（已出糧）':'Paid','Off（休息）':'Off','其他（自訂）':'Other (custom)','複製記錄':'Copy records','單日':'Single day','日期範圍':'Date range','套用':'Apply','設定':'Settings','收入統計':'Income statistics','工作日':'Work days','總工時':'Total hours','平均工時':'Average hours','總額':'Total','金額':'Amount','操作':'Actions','狀態':'Status','全部':'All','載入中…':'Loading…','載入中...':'Loading...','週一':'Mon','週二':'Tue','週三':'Wed','週四':'Thu','週五':'Fri','週六':'Sat','週日':'Sun','月度工資日曆':'Monthly Salary Calendar','月收入':'Monthly income','總工時':'Total hours','每日工資':'Daily pay','當日薪資':'Daily pay','本期出糧週期明細':'Current pay-period details','本期出糧週期':'Current pay period','本期':'Current period','上一期':'Previous period','今期總工時':'Current period total hours','上一期總工時':'Previous period total hours','點擊展開':'Click to expand','展開所有以往月份的日曆':'Expand all previous monthly calendars','匯出薪資單':'Export payslip','匯出今期':'Export current period','匯出上一期':'Export previous period','半月工資表':'Semi-monthly salary report','工資管理系統':'Salary Management System','輸入自訂備註':'Enter custom note','自訂備註':'Custom note','確認新增':'Confirm add','批量複製':'Copy in bulk','複製原有日期':'Source date to copy','複製開始日期':'Copy start date','複製結束日期':'Copy end date','單日可留空':'Leave blank for a single day','選擇記錄':'Select record','最新在上':'Newest first','暫無記錄可供複製':'No records available to copy','本期尚無工作薪資記錄':'No salary records for this period','提示':'Tip','請先在頂部控制台新增第一筆資料':'Please add the first record in the top control panel','數據統計與圖表分析':'Statistics and chart analysis','查看圖表':'View charts','收起圖表':'Hide charts','正在載入圖表分析':'Loading chart analysis','展開':'Expand','目前無資料可供匯出':'No data available to export','成功新增記錄':'Record added successfully','找不到要複製的原有日期記錄':'Source record to copy was not found','請選擇有效的複製日期範圍':'Please select a valid copy date range','批量複製完成':'Bulk copy completed','記錄已刪除':'Record deleted','成功新增':'Added successfully','新增失敗':'Failed to add','不能重複新增相同日子的記錄':'A record already exists for this date','系統時間':'System time','歷史以來總收入累計金額':'All-time cumulative income','工資神器':'Salary Manager','薪資管理系統':'Salary Management System','💰 累計薪水':'💰 Cumulative Salary','累計薪水':'Cumulative Salary','系統時間:':'System time:','Copyright © 2026 คาร์ลอส. All Rights Reserved.':'Copyright © 2026 Carlos. All Rights Reserved.'},
 th:{'新增':'เพิ่ม','修改':'แก้ไข','刪除':'ลบ','複製':'คัดลอก','儲存':'บันทึก','取消':'ยกเลิก','搜尋':'ค้นหา','清除':'ล้าง','完成':'เสร็จแล้ว','未完成':'ยังไม่เสร็จ','Paid':'จ่ายแล้ว','休息':'หยุด','其他':'อื่นๆ','工時':'ชั่วโมงทำงาน','時薪':'ค่าจ้างต่อชั่วโมง','收入':'รายได้','總收入':'รายได้รวม','本月':'เดือนนี้','上月':'เดือนก่อน','今日':'วันนี้','日期':'วันที่','備註':'หมายเหตุ','選擇月份':'เลือกเดือน','匯出':'ส่งออก','下載':'ดาวน์โหลด','服務統計':'สถิติบริการ','圖表':'กราฟ','夜間模式':'โหมดกลางคืน','日間模式':'โหมดกลางวัน','確定':'ยืนยัน','返回':'กลับ','複製到日期':'คัดลอกไปยังวันที่','開始日期':'วันที่เริ่มต้น','結束日期':'วันที่สิ้นสุด','月份':'เดือน','年':'ปี','月':'เดือน','日':'วัน','新增記錄':'เพิ่มรายการ','修改記錄':'แก้ไขรายการ','刪除記錄':'ลบรายการ','選擇日期':'เลือกวันที่','開始':'เริ่มต้น','結束':'สิ้นสุด','上一月':'เดือนก่อน','下一月':'เดือนถัดไป','上一期':'งวดก่อน','今期':'งวดปัจจุบัน','匯出今期':'ส่งออกงวดปัจจุบัน','匯出上一期':'ส่งออกงวดก่อน','目前無資料可供匯出。':'ไม่มีข้อมูลสำหรับส่งออก','歷史以來總收入累計金額：':'รายได้สะสมทั้งหมด:','系統時間：':'เวลาระบบ:','圖表載入失敗，請稍後再試':'โหลดกราฟไม่สำเร็จ กรุณาลองอีกครั้ง','新增成功':'เพิ่มสำเร็จ','儲存成功':'บันทึกสำเร็จ','確定刪除？':'ยืนยันการลบหรือไม่?','備註（可選）':'หมายเหตุ (ไม่บังคับ)','自訂備註':'หมายเหตุที่กำหนดเอง','Paid（已出糧）':'จ่ายแล้ว','Off（休息）':'หยุด','其他（自訂）':'อื่นๆ (กำหนดเอง)','複製記錄':'คัดลอกรายการ','單日':'วันเดียว','日期範圍':'ช่วงวันที่','套用':'ใช้','設定':'ตั้งค่า','收入統計':'สถิติรายได้','工作日':'วันทำงาน','總工時':'ชั่วโมงรวม','平均工時':'ชั่วโมงเฉลี่ย','總額':'ยอดรวม','金額':'จำนวนเงิน','操作':'การดำเนินการ','狀態':'สถานะ','全部':'ทั้งหมด','載入中…':'กำลังโหลด…','載入中...':'กำลังโหลด...','週一':'จ.','週二':'อ.','週三':'พ.','週四':'พฤ.','週五':'ศ.','週六':'ส.','週日':'อา.','月度工資日曆':'ปฏิทินเงินเดือนรายเดือน','月收入':'รายได้ต่อเดือน','總工時':'ชั่วโมงรวม','每日工資':'ค่าจ้างรายวัน','當日薪資':'ค่าจ้างรายวัน','本期出糧週期明細':'รายละเอียดงวดจ่ายเงินปัจจุบัน','本期出糧週期':'งวดจ่ายเงินปัจจุบัน','本期':'งวดปัจจุบัน','上一期':'งวดก่อนหน้า','今期總工時':'ชั่วโมงรวมงวดปัจจุบัน','上一期總工時':'ชั่วโมงรวมงวดก่อนหน้า','點擊展開':'คลิกเพื่อขยาย','展開所有以往月份的日曆':'ขยายปฏิทินเดือนก่อนหน้าทั้งหมด','匯出薪資單':'ส่งออกสลิปเงินเดือน','匯出今期':'ส่งออกงวดปัจจุบัน','匯出上一期':'ส่งออกงวดก่อนหน้า','半月工資表':'รายงานเงินเดือนรายครึ่งเดือน','工資管理系統':'ระบบจัดการเงินเดือน','輸入自訂備註':'ป้อนหมายเหตุที่กำหนดเอง','自訂備註':'หมายเหตุที่กำหนดเอง','確認新增':'ยืนยันการเพิ่ม','批量複製':'คัดลอกหลายรายการ','複製原有日期':'วันที่ต้นทาง','複製開始日期':'วันที่เริ่มคัดลอก','複製結束日期':'วันที่สิ้นสุดการคัดลอก','單日可留空':'เว้นว่างได้หากเป็นวันเดียว','選擇記錄':'เลือกรายการ','最新在上':'รายการล่าสุดอยู่ด้านบน','暫無記錄可供複製':'ไม่มีรายการให้คัดลอก','本期尚無工作薪資記錄':'ไม่มีบันทึกค่าจ้างในงวดนี้','提示':'คำแนะนำ','請先在頂部控制台新增第一筆資料':'กรุณาเพิ่มรายการแรกที่แผงควบคุมด้านบน','數據統計與圖表分析':'สถิติและการวิเคราะห์กราฟ','查看圖表':'ดูกราฟ','收起圖表':'ซ่อนกราฟ','正在載入圖表分析':'กำลังโหลดการวิเคราะห์กราฟ','展開':'ขยาย','目前無資料可供匯出':'ไม่มีข้อมูลสำหรับส่งออก','成功新增記錄':'เพิ่มรายการสำเร็จ','找不到要複製的原有日期記錄':'ไม่พบรายการต้นทางที่จะคัดลอก','請選擇有效的複製日期範圍':'กรุณาเลือกช่วงวันที่คัดลอกที่ถูกต้อง','批量複製完成':'คัดลอกหลายรายการเสร็จแล้ว','記錄已刪除':'ลบรายการแล้ว','成功新增':'เพิ่มสำเร็จ','新增失敗':'เพิ่มไม่สำเร็จ','不能重複新增相同日子的記錄':'มีรายการของวันที่นี้อยู่แล้ว ไม่สามารถเพิ่มซ้ำได้','系統時間':'เวลาระบบ','歷史以來總收入累計金額':'รายได้สะสมทั้งหมด','工資神器':'ตัวจัดการเงินเดือน','薪資管理系統':'ระบบจัดการเงินเดือน','💰 累計薪水':'💰 เงินเดือนสะสม','累計薪水':'เงินเดือนสะสม','系統時間:':'เวลาระบบ:','Copyright © 2026 คาร์ลอส. All Rights Reserved.':'Copyright © 2026 Carlos. All Rights Reserved.'}
 };

 Object.assign(maps.en, {'總金💰':'Total 💰'});
 Object.assign(maps.th, {'總金💰':'ยอดรวม 💰'});
 // Expanded common UI vocabulary so labels, mixed text, and dynamically rendered controls are localized.
 Object.assign(maps.en, {'新增資料': 'Add data', '新增工資': 'Add salary', '工作記錄': 'Work records', '工資記錄': 'Salary records', '工資日曆': 'Salary calendar', '出糧日期': 'Payday', '出糧日': 'Payday', '月薪': 'Monthly salary', '日薪': 'Daily pay', '時薪': 'Hourly pay', '工資': 'Salary', '薪金': 'Wages', '金額': 'Amount', '工作時數': 'Working hours', '總時數': 'Total hours', '收入合計': 'Total income', '本月總收入': "This month's total income", '上月總收入': "Last month's total income", '本期收入': 'Current period income', '上期收入': 'Previous period income', '備註': 'Notes', '狀態': 'Status', '已付款': 'Paid', '未付款': 'Unpaid', '休假': 'Day off', '自訂': 'Custom', '其他': 'Other', '選擇': 'Select', '請選擇': 'Please select', '請輸入': 'Please enter', '新增日期': 'Add date', '修改日期': 'Edit date', '刪除記錄': 'Delete record', '確認刪除': 'Confirm deletion', '確定要刪除嗎？': 'Are you sure you want to delete this?', '複製日期': 'Copy date', '複製記錄至其他日期': 'Copy records to another date', '開始日期': 'Start date', '結束日期': 'End date', '單日複製': 'Copy one day', '批量複製': 'Bulk copy', '複製成功': 'Copied successfully', '複製失敗': 'Copy failed', '儲存成功': 'Saved successfully', '儲存失敗': 'Save failed', '更新成功': 'Updated successfully', '刪除成功': 'Deleted successfully', '載入失敗': 'Failed to load', '資料載入中': 'Loading data', '沒有資料': 'No data', '暫無資料': 'No data available', '服務統計': 'Service statistics', '收入統計': 'Income statistics', '工時統計': 'Hours statistics', '每月統計': 'Monthly statistics', '圖表分析': 'Chart analysis', '顯示圖表': 'Show charts', '隱藏圖表': 'Hide charts', '展開全部': 'Expand all', '收起全部': 'Collapse all', '匯出 Excel': 'Export Excel', '匯出 CSV': 'Export CSV', '匯出 PDF': 'Export PDF', '下載圖片': 'Download image', '日間': 'Light', '夜間': 'Dark', '主題': 'Theme', '設定': 'Settings', '返回主頁': 'Back to home', '上一頁': 'Previous', '下一頁': 'Next', '上個月': 'Previous month', '下個月': 'Next month', '星期日': 'Sunday', '星期一': 'Monday', '星期二': 'Tuesday', '星期三': 'Wednesday', '星期四': 'Thursday', '星期五': 'Friday', '星期六': 'Saturday', '總計': 'Total', '平均': 'Average', '累計': 'Cumulative', '歷史累計收入': 'Historical cumulative income', '系統時間': 'System time', '載入圖表失敗，請稍後再試': 'Chart loading failed. Please try again later.'});
 Object.assign(maps.th, {'新增資料': 'เพิ่มข้อมูล', '新增工資': 'เพิ่มค่าจ้าง', '工作記錄': 'บันทึกการทำงาน', '工資記錄': 'บันทึกเงินเดือน', '工資日曆': 'ปฏิทินเงินเดือน', '出糧日期': 'วันจ่ายเงิน', '出糧日': 'วันจ่ายเงิน', '月薪': 'เงินเดือนรายเดือน', '日薪': 'ค่าจ้างรายวัน', '時薪': 'ค่าจ้างต่อชั่วโมง', '工資': 'เงินเดือน', '薪金': 'ค่าจ้าง', '金額': 'จำนวนเงิน', '工作時數': 'ชั่วโมงทำงาน', '總時數': 'ชั่วโมงรวม', '收入合計': 'รายได้รวม', '本月總收入': 'รายได้รวมเดือนนี้', '上月總收入': 'รายได้รวมเดือนก่อน', '本期收入': 'รายได้งวดปัจจุบัน', '上期收入': 'รายได้งวดก่อน', '備註': 'หมายเหตุ', '狀態': 'สถานะ', '已付款': 'จ่ายแล้ว', '未付款': 'ยังไม่จ่าย', '休假': 'วันหยุด', '自訂': 'กำหนดเอง', '其他': 'อื่นๆ', '選擇': 'เลือก', '請選擇': 'กรุณาเลือก', '請輸入': 'กรุณากรอก', '新增日期': 'เพิ่มวันที่', '修改日期': 'แก้ไขวันที่', '刪除記錄': 'ลบบันทึก', '確認刪除': 'ยืนยันการลบ', '確定要刪除嗎？': 'คุณแน่ใจหรือไม่ว่าต้องการลบ?', '複製日期': 'คัดลอกวันที่', '複製記錄至其他日期': 'คัดลอกบันทึกไปยังวันอื่น', '開始日期': 'วันที่เริ่มต้น', '結束日期': 'วันที่สิ้นสุด', '單日複製': 'คัดลอกวันเดียว', '批量複製': 'คัดลอกหลายรายการ', '複製成功': 'คัดลอกสำเร็จ', '複製失敗': 'คัดลอกไม่สำเร็จ', '儲存成功': 'บันทึกสำเร็จ', '儲存失敗': 'บันทึกไม่สำเร็จ', '更新成功': 'อัปเดตสำเร็จ', '刪除成功': 'ลบสำเร็จ', '載入失敗': 'โหลดไม่สำเร็จ', '資料載入中': 'กำลังโหลดข้อมูล', '沒有資料': 'ไม่มีข้อมูล', '暫無資料': 'ไม่มีข้อมูล', '服務統計': 'สถิติบริการ', '收入統計': 'สถิติรายได้', '工時統計': 'สถิติชั่วโมงทำงาน', '每月統計': 'สถิติรายเดือน', '圖表分析': 'วิเคราะห์กราฟ', '顯示圖表': 'แสดงกราฟ', '隱藏圖表': 'ซ่อนกราฟ', '展開全部': 'ขยายทั้งหมด', '收起全部': 'ย่อทั้งหมด', '匯出 Excel': 'ส่งออก Excel', '匯出 CSV': 'ส่งออก CSV', '匯出 PDF': 'ส่งออก PDF', '下載圖片': 'ดาวน์โหลดรูปภาพ', '日間': 'กลางวัน', '夜間': 'กลางคืน', '主題': 'ธีม', '設定': 'ตั้งค่า', '返回主頁': 'กลับหน้าหลัก', '上一頁': 'ก่อนหน้า', '下一頁': 'ถัดไป', '上個月': 'เดือนก่อน', '下個月': 'เดือนถัดไป', '星期日': 'วันอาทิตย์', '星期一': 'วันจันทร์', '星期二': 'วันอังคาร', '星期三': 'วันพุธ', '星期四': 'วันพฤหัสบดี', '星期五': 'วันศุกร์', '星期六': 'วันเสาร์', '總計': 'รวม', '平均': 'เฉลี่ย', '累計': 'สะสม', '歷史累計收入': 'รายได้สะสมย้อนหลัง', '系統時間': 'เวลาระบบ', '載入圖表失敗，請稍後再試': 'โหลดกราฟไม่สำเร็จ กรุณาลองอีกครั้ง'});
 // cargo date-field language optimization: labels and date controls in Add/Copy/Edit tabs.
 Object.assign(maps.en, {
   '工作日期':'Work Date', '複製開始日期':'Copy Start Date',
   '複製結束日期（單日可留空）':'Copy End Date (leave blank for one day)',
   '📋 複製原有日期':'📋 Source Date to Copy', '日期':'Date'
 });
 Object.assign(maps.th, {
   '工作日期':'วันที่ทำงาน', '複製開始日期':'วันที่เริ่มคัดลอก',
   '複製結束日期（單日可留空）':'วันที่สิ้นสุดการคัดลอก (เว้นว่างหากคัดลอกวันเดียว)',
   '📋 複製原有日期':'📋 วันที่ต้นทางสำหรับคัดลอก', '日期':'วันที่'
 });

 // Targeted fixes for labels previously remaining in Chinese.
 Object.assign(maps.en, {'總金💰': 'Total💰', '總金': 'Total', '更新': 'Update', '更新記錄': 'Update record', '更新成功': 'Updated successfully', '更新失敗': 'Update failed', '修改/刪除': 'Edit / Delete', '月度工資日曆': 'Monthly Salary Calendar', '月度工資日曆 (最近)': 'Monthly Salary Calendar (Recent)', '(最近)': '(Recent)', '最近': 'Recent', '星期日': 'Sunday', '星期天': 'Sunday', '星期一': 'Monday', '星期二': 'Tuesday', '星期三': 'Wednesday', '星期四': 'Thursday', '星期五': 'Friday', '星期六': 'Saturday', '星期': 'Day', '每日薪資明細表': 'Daily Salary Details', '日期': 'Date', '工時': 'Hours', '當日薪資': 'Daily Pay', '備註': 'Notes', '本期出糧週期明細': 'Current Pay-period Details', '出糧週期': 'Pay Period', '週期開始': 'Period Start', '週期結束': 'Period End', '總工時': 'Total Hours', '總薪資': 'Total Pay', '本期': 'Current Period', '上期': 'Previous Period', '新增': 'Add', '儲存': 'Save', '取消': 'Cancel', '確定': 'Confirm', '選擇日期': 'Select Date', '日期範圍': 'Date Range', '匯出': 'Export', '下載': 'Download', '操作': 'Actions', '金額': 'Amount', '總計': 'Total', '工資': 'Salary', '薪資': 'Pay', '目前無資料': 'No data available', '沒有資料': 'No data'});
 Object.assign(maps.th, {'總金💰': 'ยอดรวม💰', '總金': 'ยอดรวม', '更新': 'อัปเดต', '更新記錄': 'อัปเดตรายการ', '更新成功': 'อัปเดตสำเร็จ', '更新失敗': 'อัปเดตไม่สำเร็จ', '修改/刪除': 'แก้ไข / ลบ', '月度工資日曆': 'ปฏิทินเงินเดือนรายเดือน', '月度工資日曆 (最近)': 'ปฏิทินเงินเดือนรายเดือน (ล่าสุด)', '(最近)': '(ล่าสุด)', '最近': 'ล่าสุด', '星期日': 'วันอาทิตย์', '星期天': 'วันอาทิตย์', '星期一': 'วันจันทร์', '星期二': 'วันอังคาร', '星期三': 'วันพุธ', '星期四': 'วันพฤหัสบดี', '星期五': 'วันศุกร์', '星期六': 'วันเสาร์', '星期': 'วัน', '每日薪資明細表': 'รายละเอียดค่าจ้างรายวัน', '日期': 'วันที่', '工時': 'ชั่วโมงทำงาน', '當日薪資': 'ค่าจ้างรายวัน', '備註': 'หมายเหตุ', '本期出糧週期明細': 'รายละเอียดงวดจ่ายเงินปัจจุบัน', '出糧週期': 'งวดจ่ายเงิน', '週期開始': 'เริ่มงวด', '週期結束': 'สิ้นสุดงวด', '總工時': 'ชั่วโมงรวม', '總薪資': 'ค่าจ้างรวม', '本期': 'งวดปัจจุบัน', '上期': 'งวดก่อนหน้า', '新增': 'เพิ่ม', '儲存': 'บันทึก', '取消': 'ยกเลิก', '確定': 'ยืนยัน', '選擇日期': 'เลือกวันที่', '日期範圍': 'ช่วงวันที่', '匯出': 'ส่งออก', '下載': 'ดาวน์โหลด', '操作': 'การดำเนินการ', '金額': 'จำนวนเงิน', '總計': 'รวม', '工資': 'ค่าจ้าง', '薪資': 'ค่าจ้าง', '目前無資料': 'ไม่มีข้อมูล', '沒有資料': 'ไม่มีข้อมูล'});
 let lang='zh-Hant';
 try { const r=await fetch('/get_language',{credentials:'same-origin',cache:'no-store'}); if(r.ok){const d=await r.json();lang=d.language||lang;} } catch(e) { console.warn('Cannot read shared language setting',e); }
 document.documentElement.lang=lang==='en'?'en':lang==='th'?'th':'zh-Hant';
 document.querySelectorAll('input[type="date"]').forEach(el => {
   el.setAttribute('lang', document.documentElement.lang);
   el.setAttribute('aria-label', lang==='en' ? 'Work Date' : lang==='th' ? 'วันที่ทำงาน' : '工作日期');
 });
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
  '✈️ Cathay cargo 薪資管理系統':'✈️ Cathay cargo Salary Management System','Cathay cargo 薪資管理系統':'Cathay cargo Salary Management System',
  '✈️ Cathay cargo Pay管理系統':'✈️ Cathay cargo Pay Management System',
  'Cathay cargo Pay管理系統':'Cathay cargo Pay Management System',
  '⚙️ cargo':'⚙️ cargo',
  '⚙️ CMHK':'⚙️ CMHK',
  '總工時':'Total Hours','小時':'hours','工作日數':'Work Days','天':'days',
  '週一':'Monday','週二':'Tuesday','週三':'Wednesday','週四':'Thursday','週五':'Friday','週六':'Saturday','週日':'Sunday',
  '星期一':'Monday','星期二':'Tuesday','星期三':'Wednesday','星期四':'Thursday','星期五':'Friday','星期六':'Saturday','星期日':'Sunday',
  '更新':'Update','修改/刪除':'Edit / Delete','最近':'Recent','本期出糧週期明細':'Current Pay-period Details',
  '📊 本期出糧週期明細':'📊 Current Pay-period Details','💰 累計薪水':'💰 Cumulative Salary','累計薪水':'Cumulative Salary','薪資管理系統':'Salary Management System','系統時間:':'System time:','📅 工作日數':'📅 Work Days','⏱️ 總工時':'⏱️ Total Hours',
  '📋 歷史所有資料出勤比例':'📋 Attendance Ratio — All Historical Data','歷史所有資料出勤比例':'Attendance Ratio — All Historical Data',
  '📅 本月出勤狀態比例':'📅 This Month’s Attendance Status Ratio','本月出勤狀態比例':'This Month’s Attendance Status Ratio',
  '⏮️ 上一個月出勤狀態比例':'⏮️ Previous Month’s Attendance Status Ratio','上一個月出勤狀態比例':'Previous Month’s Attendance Status Ratio',
  '本月尚無資料':'No data for this month yet','上一個月尚無資料':'No data for the previous month yet'
 };
 const extraTh = {
  '✈️ Cathay cargo 薪資管理系統':'✈️ ระบบจัดการเงินเดือน Cathay cargo','Cathay cargo 薪資管理系統':'ระบบจัดการเงินเดือน Cathay cargo',
  '✈️ Cathay cargo Pay管理系統':'✈️ ระบบจัดการเงินเดือน Cathay cargo Pay',
  'Cathay cargo Pay管理系統':'ระบบจัดการเงินเดือน Cathay cargo Pay',
  '⚙️ cargo':'⚙️ cargo','⚙️ CMHK':'⚙️ CMHK',
  '總工時':'ชั่วโมงรวม','小時':'ชั่วโมง','工作日數':'วันทำงาน','天':'วัน',
  '週一':'จันทร์','週二':'อังคาร','週三':'พุธ','週四':'พฤหัสบดี','週五':'ศุกร์','週六':'เสาร์','週日':'อาทิตย์',
  '星期一':'วันจันทร์','星期二':'วันอังคาร','星期三':'วันพุธ','星期四':'วันพฤหัสบดี','星期五':'วันศุกร์','星期六':'วันเสาร์','星期日':'วันอาทิตย์',
  '更新':'อัปเดต','修改/刪除':'แก้ไข / ลบ','最近':'ล่าสุด','本期出糧週期明細':'รายละเอียดงวดจ่ายเงินปัจจุบัน',
  '📊 本期出糧週期明細':'📊 รายละเอียดงวดจ่ายเงินปัจจุบัน','💰 累計薪水':'💰 เงินเดือนสะสม','累計薪水':'เงินเดือนสะสม','薪資管理系統':'ระบบจัดการเงินเดือน','系統時間:':'เวลาระบบ:','📅 工作日數':'📅 วันทำงาน','⏱️ 總工時':'⏱️ ชั่วโมงรวม',
  '📋 歷史所有資料出勤比例':'📋 สัดส่วนการเข้างานจากข้อมูลย้อนหลังทั้งหมด','歷史所有資料出勤比例':'สัดส่วนการเข้างานจากข้อมูลย้อนหลังทั้งหมด',
  '📅 本月出勤狀態比例':'📅 สัดส่วนสถานะการเข้างานของเดือนนี้','本月出勤狀態比例':'สัดส่วนสถานะการเข้างานของเดือนนี้',
  '⏮️ 上一個月出勤狀態比例':'⏮️ สัดส่วนสถานะการเข้างานของเดือนก่อน','上一個月出勤狀態比例':'สัดส่วนสถานะการเข้างานของเดือนก่อน',
  '本月尚無資料':'เดือนนี้ยังไม่มีข้อมูล','上一個月尚無資料':'เดือนก่อนยังไม่มีข้อมูล'
 };
 Object.assign(maps.en, extraEn);
 Object.assign(maps.th, extraTh);

 // 檔案區域的完整多語言字典（包含動態載入、排序、預覽及上載訊息）。
 Object.assign(maps.en, {
  '檔案':'Files', '檔案清單':'File List', '修改日期':'Modified Date',
  '最新修改在上':'Newest Modified First', '最舊修改在上':'Oldest Modified First',
  '重新整理':'Refresh', '選擇要上載的檔案':'Choose a file to upload', '上載中…':'Uploading…',
  '上載成功：':'Upload successful: ', '上載失敗。':'Upload failed.', '檔案上載失敗，請稍後再試。':'File upload failed. Please try again later.',
  '正在載入檔案…':'Loading files…', '內暫時沒有檔案。匯出 JPG 後會自動儲存到這裡。':' has no files yet. Exported JPG files will be saved here automatically.',
  '請在左邊選擇 JPG 檔案進行預覽。':'Select a JPG file on the left to preview it.',
  '此檔案不是可預覽的圖片格式。':'This file is not a previewable image format.',
  '圖片載入失敗，請重新整理檔案清單。':'Image failed to load. Please refresh the file list.',
  '請先匯出 JPG，檔案會儲存在':'Export a JPG first. Files will be saved in ',
  '無法載入':'Unable to load ', '檔案清單。':' file list.',
  '載入失敗，請按「重新整理」再試。':'Failed to load. Click “Refresh” and try again.',
  '圖片預覽':'Image preview', '按一下放大至全屏預覽':'Click to enlarge to full-screen preview',
  '全屏圖片預覽':'Full-screen image preview', '退出全屏預覽':'Exit full-screen preview',
  '檔案錄全屏預覽':'File full-screen preview', '圖片載入失敗':'Image failed to load'
 });
 Object.assign(maps.th, {
  '檔案':'ไฟล์', '檔案清單':'รายการไฟล์', '修改日期':'วันที่แก้ไข',
  '最新修改在上':'เรียงจากแก้ไขล่าสุด', '最舊修改在上':'เรียงจากแก้ไขเก่าสุด',
  '重新整理':'รีเฟรช', '選擇要上載的檔案':'เลือกไฟล์ที่จะอัปโหลด', '上載中…':'กำลังอัปโหลด…',
  '上載成功：':'อัปโหลดสำเร็จ: ', '上載失敗。':'อัปโหลดไม่สำเร็จ', '檔案上載失敗，請稍後再試。':'อัปโหลดไฟล์ไม่สำเร็จ กรุณาลองอีกครั้งภายหลัง',
  '正在載入檔案…':'กำลังโหลดไฟล์…', '內暫時沒有檔案。匯出 JPG 後會自動儲存到這裡。':' ยังไม่มีไฟล์ ไฟล์ JPG ที่ส่งออกจะถูกบันทึกที่นี่โดยอัตโนมัติ',
  '請在左邊選擇 JPG 檔案進行預覽。':'เลือกไฟล์ JPG ทางด้านซ้ายเพื่อดูตัวอย่าง',
  '此檔案不是可預覽的圖片格式。':'ไฟล์นี้ไม่ใช่รูปแบบภาพที่สามารถดูตัวอย่างได้',
  '圖片載入失敗，請重新整理檔案清單。':'โหลดรูปภาพไม่สำเร็จ กรุณารีเฟรชรายการไฟล์',
  '請先匯出 JPG，檔案會儲存在':'กรุณาส่งออก JPG ก่อน ไฟล์จะถูกบันทึกไว้ที่ ',
  '無法載入':'ไม่สามารถโหลด ', '檔案清單。':' รายการไฟล์ได้',
  '載入失敗，請按「重新整理」再試。':'โหลดไม่สำเร็จ กรุณากด “รีเฟรช” แล้วลองอีกครั้ง',
  '圖片預覽':'ตัวอย่างรูปภาพ', '按一下放大至全屏預覽':'คลิกเพื่อขยายเป็นภาพเต็มหน้าจอ',
  '全屏圖片預覽':'ดูภาพเต็มหน้าจอ', '退出全屏預覽':'ออกจากการดูเต็มหน้าจอ',
  '檔案錄全屏預覽':'ประวัติการจ่ายเงินแบบเต็มหน้าจอ', '圖片載入失敗':'โหลดรูปภาพไม่สำเร็จ'
 });

 translate(document.body);
 // Also translate UI created later by calendar, modal and chart interactions.
 let queued=false; new MutationObserver(()=>{if(queued)return;queued=true;requestAnimationFrame(()=>{queued=false;translate(document.body);});}).observe(document.body,{childList:true,subtree:true,characterData:true});
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
<script>
(function () {
  const btn = document.getElementById('cargoMonthlySpeechBtn');
  const stopBtn = document.getElementById('cargoMonthlySpeechStopBtn');
  if (!btn) return;
  const setSpeechStatus = function (status) {
    btn.dataset.speechStatus = status;
    if (stopBtn) stopBtn.style.display = status === 'playing' ? 'inline-flex' : 'none';
  };
  if (!('speechSynthesis' in window) || !('SpeechSynthesisUtterance' in window)) {
    // Keep the feature silent when speech synthesis is unavailable; do not show alerts.
    btn.addEventListener('click', function () { setSpeechStatus('idle'); });
    return;
  }
  const synth = window.speechSynthesis;
  if (stopBtn) stopBtn.addEventListener('click', function () {
    try { synth.cancel(); } catch (_) {}
    setSpeechStatus('idle');
  });
  let voices = [];
  function loadVoices() { try { voices = synth.getVoices() || []; } catch (_) { voices = []; } }
  loadVoices();
  if (typeof synth.addEventListener === 'function') synth.addEventListener('voiceschanged', loadVoices);
  else if ('onvoiceschanged' in synth) synth.onvoiceschanged = loadVoices;

  btn.addEventListener('click', function () {
    // Keep speechSynthesis.speak() directly inside the user gesture for iOS.
    const langAttr = String(document.documentElement.lang || 'zh-Hant').toLowerCase();
    const lang = langAttr.startsWith('th') ? 'th' : (langAttr === 'en' || langAttr.startsWith('en-')) ? 'en' : 'zh-Hant';
    const hours = Number(btn.dataset.hours || 0);
    const monthPay = Number(btn.dataset.monthPay || 0);
    const total = Number(btn.dataset.total || 0);
    const fmt = n => new Intl.NumberFormat(lang === 'th' ? 'th-TH' : lang === 'en' ? 'en-HK' : 'zh-HK', {maximumFractionDigits: 2}).format(n);
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
    const norm = value => String(value || '').toLowerCase().replace(/_/g, '-');
    const wanted = norm(voiceLang);
    let voice = voices.find(v => norm(v.lang) === wanted);
    if (!voice) voice = voices.find(v => voicePattern.test(norm(v.lang)));
    if (!voice && lang === 'zh-Hant') voice = voices.find(v => /^zh(-|$)/i.test(norm(v.lang)));

    // Do not block iOS when its voice list is empty or does not expose a matching
    // voice. Set the BCP-47 language and let the OS choose its own installed voice.
    // If a matching voice exists, prefer it; otherwise still attempt native synthesis.
    try {
      synth.cancel();
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.lang = voice ? (voice.lang || voiceLang) : voiceLang;
      if (voice) utterance.voice = voice;
      utterance.rate = 0.9;
      setSpeechStatus('idle');
      utterance.onstart = function () { setSpeechStatus('playing'); };
      utterance.onend = function () { setSpeechStatus('idle'); };
      utterance.onerror = function () {
        // Fail silently as requested; always hide the stop button on failure.
        setSpeechStatus('idle');
      };
      synth.speak(utterance);
      // Some iOS versions do not start synthesis unless resumed after cancel().
      if (typeof synth.resume === 'function') synth.resume();
    } catch (err) {
      // Fail silently; never interrupt the user with a speech error notification.
      setSpeechStatus('idle');
    }
  });
})();
</script>

<script>
(function () {
    const popup = document.getElementById('cargoCalendarActionPopup');
    if (!popup) return;
    let selectedDate = '';
    const panel = document.getElementById('controlPanelCollapse');

    function hidePopup() { popup.hidden = true; }
    function updatePopupActions(hasRecord) {
        popup.querySelectorAll('[data-requires-record="1"]').forEach(function (button) {
            button.hidden = !hasRecord;
        });
    }
    function openPanel(tabSelector) {
        if (panel && window.bootstrap && bootstrap.Collapse) {
            bootstrap.Collapse.getOrCreateInstance(panel, {toggle:false}).show();
        } else if (panel) {
            panel.classList.add('show');
        }
        if (tabSelector && window.bootstrap && bootstrap.Tab) {
            const tab = document.querySelector('[data-bs-target="' + tabSelector + '"]');
            if (tab) bootstrap.Tab.getOrCreateInstance(tab).show();
        }
    }
    function setDateAndSubmit(action) {
        const addDate = document.querySelector('#add-tab input[name="date"]');
        const source = document.querySelector('#copy-tab select[name="source_date"]');
        const targetStart = document.querySelector('#copy-tab input[name="target_start"]');
        const targetEnd = document.querySelector('#copy-tab input[name="target_end"]');
        const editSelect = document.getElementById('editSelect');

        if (action === 'add') {
            if (addDate) addDate.value = selectedDate;
            openPanel('#add-tab');
        } else if (action === 'copy') {
            if (source) {
                const opt = Array.from(source.options).find(o => o.value === selectedDate);
                if (opt) source.value = selectedDate;
            }
            if (targetStart) targetStart.value = selectedDate;
            if (targetEnd) targetEnd.value = '';
            openPanel('#copy-tab');
        } else if (action === 'edit' || action === 'delete') {
            if (editSelect) {
                const opt = Array.from(editSelect.options).find(o => o.dataset.date === selectedDate);
                if (opt) {
                    editSelect.value = opt.value;
                    if (typeof window.loadEditData === 'function') window.loadEditData(editSelect);
                    openPanel('#edit-tab');
                    if (action === 'delete') {
                        const idx = document.getElementById('delete_index');
                        if (idx) idx.value = opt.value;
                        const form = document.getElementById('deleteForm');
                        if (form) form.requestSubmit ? form.requestSubmit() : form.submit();
                    }
                } else {
                    openPanel('#add-tab');
                    if (addDate) addDate.value = selectedDate;
                }
            } else {
                openPanel('#add-tab');
                if (addDate) addDate.value = selectedDate;
            }
        }
        hidePopup();
    }

    document.addEventListener('click', function (event) {
        const day = event.target.closest('.cargo-calendar-day');
        if (day) {
            event.preventDefault();
            selectedDate = day.dataset.date || '';
            updatePopupActions(day.dataset.hasRecord === '1');
            const rect = day.getBoundingClientRect();
            popup.hidden = false;
            const width = popup.offsetWidth || 172;
            const height = popup.offsetHeight || 50;
            let left = Math.min(Math.max(8, rect.left + rect.width / 2 - width / 2), window.innerWidth - width - 8);
            let top = rect.bottom + 6;
            if (top + height > window.innerHeight - 8) top = Math.max(8, rect.top - height - 6);
            popup.style.left = left + 'px';
            popup.style.top = top + 'px';
            return;
        }
        const actionButton = event.target.closest('[data-cargo-action]');
        if (actionButton && popup.contains(actionButton)) {
            setDateAndSubmit(actionButton.dataset.cargoAction);
            return;
        }
        if (!popup.contains(event.target)) hidePopup();
    });

    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape') hidePopup();
    });
    window.addEventListener('scroll', hidePopup, {passive:true});

    // Keep the control panel, active tab and viewport position after a form action reloads the page.
    document.querySelectorAll('#controlPanelCollapse form[method="POST"]').forEach(function (form) {
        form.addEventListener('submit', function () {
            try {
                const activeTab = document.querySelector('#controlTab .nav-link.active');
                sessionStorage.setItem('cargoPanelRestore', JSON.stringify({
                    scrollY: window.scrollY,
                    tab: activeTab ? activeTab.getAttribute('data-bs-target') : '#add-tab'
                }));
            } catch (e) {}
        });
    });
    window.addEventListener('DOMContentLoaded', function () {
        try {
            const saved = sessionStorage.getItem('cargoPanelRestore');
            if (!saved) return;
            sessionStorage.removeItem('cargoPanelRestore');
            const state = JSON.parse(saved);
            openPanel(state.tab || '#add-tab');
            requestAnimationFrame(function () {
                window.scrollTo(0, Number(state.scrollY) || 0);
            });
        } catch (e) {}
    });
})();
</script>

<style id="cargoCustomDatePickerStyles">
.cargo-date-trigger{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:6px;
 width:100%;min-width:0;min-height:31px;padding:4px 8px!important;border:1px solid var(--bs-border-color,#777)!important;
 border-radius:6px!important;background:var(--bs-body-bg,#fff)!important;color:var(--bs-body-color,#222)!important;
 font:inherit!important;font-size:12px!important;cursor:pointer;text-align:left;white-space:nowrap}
.cargo-date-trigger .cargo-date-label{overflow:hidden;text-overflow:ellipsis}
.cargo-date-trigger .cargo-date-icon{flex:0 0 auto}
#cargoDatePickerPopup{position:fixed;z-index:20000;width:270px;padding:12px;border-radius:13px;
 background:var(--bs-body-bg,#fff);color:var(--bs-body-color,#222);border:1px solid var(--bs-border-color,rgba(120,160,130,.5));
 box-shadow:0 12px 32px rgba(0,0,0,.3);font:14px system-ui,-apple-system,"Segoe UI",sans-serif;box-sizing:border-box}
#cargoDatePickerPopup[hidden]{display:none!important}
#cargoDatePickerPopup *{box-sizing:border-box}
#cargoDatePickerPopup .dp-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:10px}
#cargoDatePickerPopup .dp-title{font-weight:700;flex:1;text-align:center}
#cargoDatePickerPopup button{border:0;border-radius:8px;cursor:pointer;color:inherit;background:rgba(120,180,130,.18)}
#cargoDatePickerPopup .dp-nav{width:34px;height:32px;font-size:20px}
#cargoDatePickerPopup .dp-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;text-align:center}
#cargoDatePickerPopup .dp-week{font-size:12px;font-weight:700;opacity:.72;padding:5px 0}
#cargoDatePickerPopup .dp-day{height:31px;font-size:13px;background:transparent}
#cargoDatePickerPopup .dp-day:hover{background:rgba(56,239,125,.25)}
#cargoDatePickerPopup .dp-day.dp-selected{background:#198754;color:#fff;font-weight:700}
#cargoDatePickerPopup .dp-day.dp-today{outline:1px solid #38a66a}
#cargoDatePickerPopup .dp-foot{display:flex;justify-content:space-between;margin-top:9px}
#cargoDatePickerPopup .dp-foot button{padding:6px 10px;font-size:12px}
@media(max-width:600px){.cargo-date-trigger{font-size:10px!important;padding:4px 5px!important;min-height:29px}}

/* FINAL 3D hit-area correction:
   Keep page-level/layout wrappers flat. Their preserve-3d context was allowing
   raised descendants to overlap neighbouring controls and intercept clicks.
   Individual buttons, chart elements and table cells retain their own 3D styling. */
body.salary-3d-mode .container,
body.salary-3d-mode main,
body.salary-3d-mode .content,
body.salary-3d-mode .main-content {
  transform: none !important;
  transform-style: flat !important;
  perspective: none !important;
}
body.salary-3d-mode .card,
body.salary-3d-mode .calendar-card,
body.salary-3d-mode .liquid-glass-card,
body.salary-3d-mode .chart-container,
body.salary-3d-mode .chart-card,
body.salary-3d-mode .report-card,
body.salary-3d-mode .table-responsive,
body.salary-3d-mode .sticky-control-panel,
body.salary-3d-mode .top-console {
  position: relative;
  transform: none !important;
  transform-style: flat !important;
  perspective: none !important;
  filter: none !important;
  box-shadow: 0 2px 8px rgba(0,0,0,.16),
              0 5px 14px rgba(0,0,0,.10),
              inset 0 1px 0 rgba(255,255,255,.32) !important;
}
body.salary-3d-mode .card:hover:not(.sticky-control-panel),
body.salary-3d-mode .calendar-card:hover,
body.salary-3d-mode .chart-container:hover,
body.salary-3d-mode .chart-card:hover,
body.salary-3d-mode .report-card:hover,
body.salary-3d-mode .table-responsive:hover,
body.salary-3d-mode .sticky-control-panel:hover,
body.salary-3d-mode .top-console:hover {
  transform: none !important;
  filter: none !important;
  box-shadow: 0 2px 8px rgba(0,0,0,.16),
              0 5px 14px rgba(0,0,0,.10),
              inset 0 1px 0 rgba(255,255,255,.32) !important;
}

/* Restore cargo calendar day-tile lift in 3D mode only.
   Large panels stay seated; individual date cells keep their raised hover effect. */
body.salary-3d-mode .night-calendar tbody td:not([style*="background:transparent"]):hover {
    position: relative !important;
    z-index: 20 !important;
    transform: perspective(600px) translateY(-3px) translateZ(8px) rotateX(4deg) scale(1.025) !important;
    box-shadow: 0 7px 0 rgba(0,0,0,.24),
                0 12px 18px rgba(0,0,0,.24),
                inset 0 1px 0 rgba(255,255,255,.18) !important;
    filter: brightness(1.08) !important;
}
body.salary-3d-mode .night-calendar tbody td.today-glow-cell,
body.salary-3d-mode .night-calendar tbody td[style*="border: 2px solid #38ef7d"] {
    transform: translateZ(8px) !important;
}

/* All cargo calendar date cells keep a raised 3D appearance, not only on hover.
   The hover state adds a little extra lift while preserving clickability. */
body.salary-3d-mode .night-calendar tbody td {
    position: relative !important;
    transform: translateZ(8px) !important;
    transform-style: flat !important;
    box-shadow: 0 4px 0 rgba(0,0,0,.24),
                0 7px 12px rgba(0,0,0,.16),
                inset 0 1px 0 rgba(255,255,255,.10) !important;
}
body.salary-3d-mode .night-calendar tbody td:not([style*="background:transparent"]):hover {
    z-index: 20 !important;
    transform: perspective(600px) translateY(-3px) translateZ(12px) rotateX(4deg) scale(1.025) !important;
    box-shadow: 0 7px 0 rgba(0,0,0,.24),
                0 12px 18px rgba(0,0,0,.24),
                inset 0 1px 0 rgba(255,255,255,.18) !important;
    filter: brightness(1.08) !important;
}

/* cargo calendar: seven equal columns, compact edge-to-edge layout.
   Prevent Bootstrap's .table-responsive from creating horizontal scrolling. */
.cargo-calendar-responsive {
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    overflow-x: clip !important;
    overflow-y: visible !important;
    padding: 0 !important;
    margin-left: 0 !important;
    margin-right: 0 !important;
}
.cargo-calendar-responsive > .night-calendar {
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    table-layout: fixed !important;
    margin: 0 !important;
    border-collapse: separate !important;
    border-spacing: 1px !important;
}
.cargo-calendar-responsive > .night-calendar th,
.cargo-calendar-responsive > .night-calendar td {
    width: 14.285714% !important;
    min-width: 0 !important;
    max-width: 14.285714% !important;
    padding: 1px !important;
    box-sizing: border-box !important;
    overflow-wrap: anywhere !important;
    word-break: break-word !important;
}
.cargo-calendar-responsive > .night-calendar th {
    font-size: clamp(10px, 2.8vw, 13px) !important;
    white-space: nowrap !important;
}
.cargo-calendar-responsive > .night-calendar td {
    height: 42px !important;
    font-size: clamp(10px, 2.8vw, 13px) !important;
}
.cargo-calendar-responsive > .night-calendar td span {
    font-size: clamp(8px, 2.35vw, 11px) !important;
}
@media (max-width: 420px) {
    .cargo-calendar-responsive > .night-calendar {
        border-spacing: 1px !important;
    }
    .cargo-calendar-responsive > .night-calendar th,
    .cargo-calendar-responsive > .night-calendar td {
        padding: 0 !important;
    }
}

/* cargo calendar refinement: remove the enclosing frame and let each day tile
   carry the raised shadow, creating a lighter, less boxed-in appearance. */
.cargo-calendar-responsive,
.cargo-calendar-responsive > .night-calendar,
.cargo-calendar-responsive > .night-calendar.table,
.cargo-calendar-responsive > .night-calendar.table-bordered {
    border: 0 !important;
    outline: 0 !important;
    box-shadow: none !important;
    background: transparent !important;
    border-radius: 0 !important;
}
.cargo-calendar-responsive > .night-calendar {
    border-collapse: separate !important;
    border-spacing: 2px !important;
}
.cargo-calendar-responsive > .night-calendar thead th {
    border: 0 !important;
    border-radius: 7px !important;
    box-shadow: 0 3px 0 rgba(0,0,0,.20),
                0 6px 10px rgba(0,0,0,.14),
                inset 0 1px 0 rgba(255,255,255,.12) !important;
}
.cargo-calendar-responsive > .night-calendar tbody td {
    border: 1px solid rgba(255,255,255,.07) !important;
    border-radius: 8px !important;
    box-shadow: 0 4px 0 rgba(0,0,0,.25),
                0 7px 12px rgba(0,0,0,.18),
                inset 0 1px 0 rgba(255,255,255,.10) !important;
}
.cargo-calendar-responsive > .night-calendar tbody td:hover {
    box-shadow: 0 6px 0 rgba(0,0,0,.25),
                0 10px 16px rgba(0,0,0,.22),
                inset 0 1px 0 rgba(255,255,255,.16) !important;
}
/* Keep the calendar itself visually unframed in both normal and 3D modes. */
body.salary-3d-mode .cargo-calendar-responsive,
body.salary-3d-mode .cargo-calendar-responsive > .night-calendar {
    border: 0 !important;
    outline: 0 !important;
    box-shadow: none !important;
    filter: none !important;
    transform: none !important;
}

/* cargo top control strip: remove the large card frame around the ⚙️ cargo pill.
   This is the cargo page's own .sticky-control-panel, not the outer iframe. */
.sticky-control-panel,
.sticky-control-panel.liquid-glass-card,
.sticky-control-panel.card,
.sticky-control-panel > .card-header {
    border: 0 !important;
    outline: 0 !important;
    border-radius: 0 !important;
    box-shadow: none !important;
    filter: none !important;
    background: transparent !important;
    background-image: none !important;
}
.sticky-control-panel::before,
.sticky-control-panel::after,
.sticky-control-panel > .card-header::before,
.sticky-control-panel > .card-header::after {
    content: none !important;
    display: none !important;
    border: 0 !important;
    box-shadow: none !important;
}
body.salary-3d-mode .sticky-control-panel,
body.salary-3d-mode .sticky-control-panel.liquid-glass-card,
body.salary-3d-mode .sticky-control-panel.card,
body.salary-3d-mode .sticky-control-panel:hover {
    border: 0 !important;
    outline: 0 !important;
    border-radius: 0 !important;
    box-shadow: none !important;
    filter: none !important;
    transform: none !important;
}

/* Tighten the gap between the cargo control pill and the calendar below it. */
body > .sticky-control-panel,
body .sticky-control-panel {
    margin-bottom: 0 !important;
}
/* The calendar section follows the top control panel in the cargo page. */
body > .container.my-4 {
    margin-top: 0 !important;
}
/* Remove only the calendar card's top whitespace; retain side/bottom padding. */
body > .container.my-4 > .card.liquid-glass-card {
    padding-top: 0 !important;
    margin-top: 0 !important;
}

/* Keep the ⚙️ cargo title at exactly the same position in 2D and 3D modes.
   Use the 3D layout's centered 38px header as the shared layout in both modes. */
.sticky-control-panel > .card-header {
    height: auto !important;
    min-height: 38px !important;
    overflow: visible !important;
    padding: 4px 8px !important;
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
}
#cargoControlHeaderTitle {
    position: relative !important;
    top: auto !important;
    left: auto !important;
    right: auto !important;
    width: auto !important;
    max-width: 100% !important;
    margin: 0 !important;
    transform: none !important;
    line-height: 1.35 !important;
    z-index: 2 !important;
    pointer-events: auto !important;
}
#cargoControlHeaderTitle .cargo-glass-pill {
    pointer-events: auto !important;
}

/* 2D mode: place the cargo control pill in its own row ABOVE the calendar card.
   It must not float over or appear inside the calendar area. Keep 3D mode unchanged. */
body:not(.salary-3d-mode) > .sticky-control-panel {
    /* In 2D, keep the control row in normal document flow.
       Sticky top:30px was shifting it over the calendar content. */
    position: static !important;
    top: auto !important;
    transform: none !important;
    height: auto !important;
    min-height: 38px !important;
    margin: 0 !important;
    padding: 0 !important;
    overflow: visible !important;
}
body:not(.salary-3d-mode) > .sticky-control-panel > .card-header {
    position: relative !important;
    top: auto !important;
    transform: none !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    height: 38px !important;
    min-height: 38px !important;
    padding: 4px 8px !important;
    overflow: visible !important;
}
body:not(.salary-3d-mode) #cargoControlHeaderTitle {
    position: relative !important;
    inset: auto !important;
    display: block !important;
    width: auto !important;
    max-width: 100% !important;
    margin: 0 !important;
    transform: none !important;
    line-height: 1.35 !important;
    pointer-events: auto !important;
}
</style>
<script id="cargoCustomDatePickerScript">
(function(){
 'use strict';
 if(window.__cargoCustomDatePickerReady)return;
 window.__cargoCustomDatePickerReady=true;
 const popup=document.createElement('div');popup.id='cargoDatePickerPopup';popup.hidden=true;
 document.body.appendChild(popup);
 // Keep clicks inside the picker from reaching the outside-click closer.
 // Month navigation rebuilds the popup DOM and detaches the clicked arrow.
 popup.addEventListener('click',e=>e.stopPropagation());
 let activeInput=null,viewYear=0,viewMonth=0;
 const weekdays=['日','一','二','三','四','五','六'];
 const pad=n=>String(n).padStart(2,'0');
 const toISO=(y,m,d)=>`${y}-${pad(m+1)}-${pad(d)}`;
 function parse(v){const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(v||'');return m?{y:+m[1],m:+m[2]-1,d:+m[3]}:null;}
 function today(){const d=new Date();return {y:d.getFullYear(),m:d.getMonth(),d:d.getDate()};}
 function display(v){const d=parse(v);return d?`${pad(d.d)}/${pad(d.m+1)}/${d.y}`:'選擇日期';}
 function updateTrigger(input){const t=input._cargoDateTrigger;if(t){const label=t.querySelector('.cargo-date-label');if(label)label.textContent=display(input.value);}}
 function close(){popup.hidden=true;activeInput=null;}
 function commit(value){
   if(!activeInput)return;
   activeInput.value=value;
   activeInput.dispatchEvent(new Event('input',{bubbles:true}));
   activeInput.dispatchEvent(new Event('change',{bubbles:true}));
   updateTrigger(activeInput);close();
 }
 function render(){
   const selected=parse(activeInput.value),first=new Date(viewYear,viewMonth,1).getDay();
   const count=new Date(viewYear,viewMonth+1,0).getDate();
   popup.innerHTML='';
   const head=document.createElement('div');head.className='dp-head';
   const prev=document.createElement('button');prev.type='button';prev.className='dp-nav';prev.textContent='‹';prev.setAttribute('aria-label','上一個月');
   prev.onclick=()=>{viewMonth--;if(viewMonth<0){viewMonth=11;viewYear--;}render();};
   const title=document.createElement('div');title.className='dp-title';
   title.textContent=new Date(viewYear,viewMonth,1).toLocaleDateString(document.documentElement.lang||'zh-Hant',{year:'numeric',month:'long'});
   const next=document.createElement('button');next.type='button';next.className='dp-nav';next.textContent='›';next.setAttribute('aria-label','下一個月');
   next.onclick=()=>{viewMonth++;if(viewMonth>11){viewMonth=0;viewYear++;}render();};
   head.append(prev,title,next);popup.append(head);
   const grid=document.createElement('div');grid.className='dp-grid';
   weekdays.forEach(w=>{const e=document.createElement('div');e.className='dp-week';e.textContent=w;grid.append(e);});
   for(let i=0;i<first;i++){const blank=document.createElement('span');grid.append(blank);}
   for(let day=1;day<=count;day++){
     const b=document.createElement('button');b.type='button';b.className='dp-day';b.textContent=day;
     if(selected&&selected.y===viewYear&&selected.m===viewMonth&&selected.d===day)b.classList.add('dp-selected');
     const t=today();if(t.y===viewYear&&t.m===viewMonth&&t.d===day)b.classList.add('dp-today');
     b.addEventListener('click',()=>commit(toISO(viewYear,viewMonth,day)));
     grid.append(b);
   }
   popup.append(grid);
   const foot=document.createElement('div');foot.className='dp-foot';
   const clear=document.createElement('button');clear.type='button';clear.textContent='清除';clear.onclick=()=>commit('');
   const now=document.createElement('button');now.type='button';now.textContent='今天';now.onclick=()=>{const t=today();commit(toISO(t.y,t.m,t.d));};
   foot.append(clear,now);popup.append(foot);
 }
 function open(input,trigger){
   activeInput=input;const d=parse(input.value)||today();viewYear=d.y;viewMonth=d.m;render();popup.hidden=false;
   const r=trigger.getBoundingClientRect(),w=270,h=popup.offsetHeight;
   popup.style.left=Math.max(8,Math.min(r.left,window.innerWidth-w-8))+'px';
   popup.style.top=(r.bottom+h+8>window.innerHeight?Math.max(8,r.top-h-8):r.bottom+6)+'px';
 }
 document.querySelectorAll('input[type="date"]').forEach(input=>{
   const trigger=document.createElement('button');trigger.type='button';trigger.className='cargo-date-trigger';
   const label=document.createElement('span');label.className='cargo-date-label';label.textContent=display(input.value);
   const icon=document.createElement('span');icon.className='cargo-date-icon';icon.textContent='📅';
   trigger.append(label,icon);input._cargoDateTrigger=trigger;
   input.dataset.cargoDateRequired=input.required?'1':'0';
   input.type='hidden';input.required=false;
   input.insertAdjacentElement('afterend',trigger);
   trigger.addEventListener('click',()=>open(input,trigger));
   input.addEventListener('change',()=>updateTrigger(input));
 });
 // Keep the visible date label synchronized when existing cargo scripts assign .value directly.
 const editDate=document.getElementById('edit_date');
 if(editDate){
   const original=window.loadEditData;
   if(typeof original==='function'){
     window.loadEditData=function(sel){original(sel);updateTrigger(editDate);};
   }
 }
 document.addEventListener('click',e=>{if(!popup.hidden&&!popup.contains(e.target)&&!e.target.closest('.cargo-date-trigger'))close();});
 document.addEventListener('keydown',e=>{if(e.key==='Escape')close();});
 window.addEventListener('resize',close);window.addEventListener('scroll',close,true);
 // Reproduce native required-date validation for hidden ISO date fields.
 document.addEventListener('submit',e=>{
   const form=e.target;
   if(!(form instanceof HTMLFormElement))return;
   const missing=[...form.querySelectorAll('input[type="hidden"][data-cargo-date-required="1"]')].find(i=>!i.value);
   if(missing){e.preventDefault();const trigger=missing._cargoDateTrigger;if(trigger)trigger.click();}
 },true);
 // Date values set by the calendar's add/copy/edit actions are reflected on the trigger.
 ['#add-tab input[name="date"]','#copy-tab input[name="target_start"]','#copy-tab input[name="target_end"]','#edit_date'].forEach(sel=>{
   const input=document.querySelector(sel);if(!input)return;
   const desc=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value');
   if(!desc||!desc.get||!desc.set)return;
   let current=input.value;
   Object.defineProperty(input,'value',{
     configurable:true,enumerable:desc.enumerable,
     get(){return desc.get.call(this);},
     set(v){desc.set.call(this,v);updateTrigger(this);}
   });
 });
})();
</script>

</body>
</html>
"""

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8501, debug=False, use_reloader=False)
