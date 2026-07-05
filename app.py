from flask import Flask, render_template, request, redirect, send_file, session
from flask_sqlalchemy import SQLAlchemy
from collections import defaultdict
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
import os
import io
import calendar

app = Flask(__name__)
app.secret_key = "anything-secret"

db_url = os.getenv("DATABASE_URL")

if db_url:
    db_url = db_url.strip()
if db_url and db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://","postgresql://",1)


app.config["SQLALCHEMY_DATABASE_URI"] = db_url or "sqlite:///instance/database.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_pre_ping": True,
    "pool_recycle": 280
}

db = SQLAlchemy(app)


class Data(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    branch = db.Column(db.String(100))
    date = db.Column(db.String(20))
    period = db.Column(db.String(10))
    time = db.Column(db.String(20))
    teacher = db.Column(db.String(100))
    status = db.Column(db.String(20))
    sub_teacher = db.Column(db.String(100))
    note = db.Column(db.String(200))


with app.app_context():
    db.create_all()


SUNDAY_PERIOD_TIMES = {
    "1": "9:30 ～ 10:50",
    "2": "11:00 ～ 12:20",
    "3": "13:20 ～ 14:40",
    "4": "15:00 ～ 16:20",
    "5": "16:30 ～ 17:50",
    "6": "18:05 ～ 19:25",
    "7": "19:35 ～ 20:55",
}

SATURDAY_PERIOD_TIMES = {
    "1": "9:30 ～ 10:50",
    "1-1": "1限 1/2　9:20 ～ 10:05",
    "1-2": "1限 2/2　10:05 ～ 10:50",
    "2": "11:00 ～ 12:20",
    "2-1": "2限 1/2　10:50 ～ 11:35",
    "2-2": "2限 2/2　11:35 ～ 12:20",
    "3s": "3限（小学生集団）13:20 ～ 14:20",
    "3": "13:20 ～ 14:40",
    "3-1": "3限 1/2　13:10 ～ 13:55",
    "3-2": "3限 2/2　13:55 ～ 14:40",
    "4": "14:50 ～ 16:10",
    "4-1": "4限 1/2　14:50 ～ 15:35",
    "4-2": "4限 2/2　15:35 ～ 16:20",
    "5": "16:30 ～ 17:50",
    "6": "18:05 ～ 19:25",
    "7": "19:35 ～ 20:55",
}

WEEKDAY_PERIOD_TIMES = {
    "1-1": "1限 1/2　9:20 ～ 10:05",
    "1-2": "1限 2/2　10:05 ～ 10:50",
    "2": "11:00 ～ 12:20",
    "2-1": "2限 1/2　10:50 ～ 11:35",
    "2-2": "2限 2/2　11:35 ～ 12:20",
    "3s": "3限（小学生集団）13:20 ～ 14:20",
    "3": "13:20 ～ 14:40",
    "3-1": "3限 1/2　13:10 ～ 13:55",
    "3-2": "3限 2/2　13:55 ～ 14:40",
    "4": "14:50 ～ 16:10",
    "4-1": "4限 1/2　15:25 ～ 16:10",
    "4-2": "4限 2/2　16:10 ～ 16:55",
    "5s": "5限（小学生集団）17:15 ～ 18:15",
    "5": "17:15 ～ 18:35",
    "5-1": "5限 1/2　17:05 ～ 17:50",
    "5-2": "5限 2/2　17:50 ～ 18:35",
    "6": "18:45 ～ 20:05",
    "7": "20:15 ～ 21:35",
}

def get_period_times(date_str):
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    weekday = date_obj.weekday()

    if weekday == 6:
        return SUNDAY_PERIOD_TIMES
    elif weekday == 5:
        return SATURDAY_PERIOD_TIMES
    else:
        return WEEKDAY_PERIOD_TIMES
def calc_minutes(time_text):
    try:
        if "～" not in time_text:
            return ""

        time_part = time_text.split("　")[-1]
        start, end = time_part.split("～")

        start = start.strip()
        end = end.strip()

        start_dt = datetime.strptime(start, "%H:%M")
        end_dt = datetime.strptime(end, "%H:%M")

        minutes = int((end_dt - start_dt).total_seconds() / 60)
        return minutes

    except:
        return ""

# ---------------- カレンダー画面 ----------------
@app.route("/")
def index():
    year = request.args.get("year", type=int)
    month = request.args.get("month", type=int)

    today = datetime.today()

    if not year or not month:
        year = today.year
        month = today.month

    if month < 1:
        month = 12
        year -= 1
    elif month > 12:
        month = 1
        year += 1

    cal = calendar.monthcalendar(year, month)

    start_date = f"{year:04d}-{month:02d}-01"
    last_day = calendar.monthrange(year, month)[1]
    end_date = f"{year:04d}-{month:02d}-{last_day:02d}"

    records = Data.query.filter(
        Data.date >= start_date,
        Data.date <= end_date
    ).all()

    teacher_status_by_date = defaultdict(dict)

    for r in records:
        if not r.teacher:
            continue

        teacher_status_by_date[r.date][r.teacher] = r.status

    count_by_date = {}

    for date_key, teachers in teacher_status_by_date.items():
        blue = 0
        red = 0
        yellow = 0

        for teacher, status in teachers.items():
            if status == "決定":
                blue += 1
            elif status == "未定":
                red += 1
            else:
                yellow += 1

        count_by_date[date_key] = {
            "blue": blue,
            "red": red,
            "yellow": yellow
        }

    return render_template(
        "index.html",
        year=year,
        month=month,
        cal=cal,
        count_by_date=count_by_date,
        today=today.strftime("%Y-%m-%d")
    )

# ---------------- 日付を押した後の入力画面 ----------------
@app.route("/day/<date>", methods=["GET", "POST"])
def day(date):
    period_times = get_period_times(date)
    if request.method == "POST":
        action = request.form.get("action")

        # 一括削除
        if action == "bulk_delete":
            selected_ids = request.form.getlist("selected_ids")
            for id in selected_ids:
                data = Data.query.get(id)
                if data:
                    db.session.delete(data)

            db.session.commit()
            return redirect(f"/day/{date}")

        # 一括更新
        if action == "bulk_update":
            selected_ids = request.form.getlist("selected_ids")

            bulk_teacher = request.form.get("bulk_teacher")
            bulk_status = request.form.get("bulk_status")
            bulk_sub_teacher = request.form.get("bulk_sub_teacher")
            bulk_note = request.form.get("bulk_note")

            for id in selected_ids:
                data = Data.query.get(id)
                if data:
                    if bulk_teacher:
                        data.teacher = bulk_teacher

                    if bulk_status:
                        data.status = bulk_status
                        if bulk_status == "未定":
                            data.sub_teacher = ""

                    if bulk_status == "決定" and bulk_sub_teacher:
                        data.sub_teacher = bulk_sub_teacher

                    if bulk_note:
                        data.note = bulk_note

            db.session.commit()
            return redirect(f"/day/{date}")

        # 通常追加
        if action == "add":
            branch = request.form.get("branch")
            teacher = request.form.get("teacher")
            status = request.form.get("status")
            sub_teacher = request.form.get("sub_teacher")
            note = request.form.get("note")
            periods = request.form.getlist("period")

            if status != "決定":
                sub_teacher = ""

            for p in periods:
                new = Data(
                    branch=branch,
                    date=date,
                    period=p,
                    time=period_times.get(p, ""),
                    teacher=teacher,
                    status=status,
                    sub_teacher=sub_teacher,
                    note=note
                )
                db.session.add(new)

            db.session.commit()
            return redirect(f"/day/{date}")

    records = Data.query.filter_by(date=date).order_by(Data.period).all()

    records_by_period = defaultdict(list)
    for r in records:
        records_by_period[r.period].append(r)

    return render_template(
        "day.html",
        date=date,
        period_times=period_times,
        records_by_period=records_by_period,
        records=records,
        calc_minutes=calc_minutes
    )

@app.route("/edit/<int:id>", methods=["GET", "POST"])
def edit(id):
    data = Data.query.get_or_404(id)

    if request.method == "POST":
        data.branch = request.form.get("branch")
        data.teacher = request.form.get("teacher")
        data.status = request.form.get("status")
        data.sub_teacher = request.form.get("sub_teacher")
        data.note = request.form.get("note")

        if data.status != "決定":
            data.sub_teacher = ""

        db.session.commit()
        return redirect(f"/day/{data.date}")

    return render_template("edit.html", data=data)
# ---------------- 削除 ----------------
@app.route("/delete/<int:id>")
def delete(id):
    data = Data.query.get_or_404(id)
    date = data.date
    db.session.delete(data)
    db.session.commit()
    return redirect(f"/day/{date}")


# ---------------- Excel出力 ----------------
@app.route("/export")
def export():
    records = Data.query.order_by(Data.date, Data.period).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "代講一覧"

    ws.append(["日付", "校舎", "限", "時間", "先生", "状態", "代講先生", "備考"])

    for r in records:
        ws.append([
            r.date,
            r.branch,
            r.period,
            r.time,
            r.teacher,
            r.status,
            r.sub_teacher,
            r.note
        ])

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 4

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center")

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="daikou.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True)


