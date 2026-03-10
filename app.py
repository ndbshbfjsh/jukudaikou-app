from flask import Flask, render_template, request, redirect, send_file, session
from flask_sqlalchemy import SQLAlchemy
from collections import defaultdict
from datetime import datetime
from openpyxl import Workbook
from datetime import datetime, timedelta
import os
import io

app = Flask(__name__)
app.secret_key= "anything-secret"
app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv(
    "DATABASE_URL",
    "sqlite:///database.db"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)



class Data(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    branch = db.Column(db.String(100))  # ←追加
    date = db.Column(db.String(20))
    period = db.Column(db.String(10))
    time = db.Column(db.String(20))
    teacher = db.Column(db.String(100))
    status = db.Column(db.String(20))
    sub_teacher = db.Column(db.String(100))
    note = db.Column(db.String(200))
with app.app_context():
    db.create_all()

    
# ---------------- 一覧 ----------------
@app.route("/toggle_today")
def toggle_today():
    session["today_only"] = not session.get("today_only", False)
    return redirect("/")

@app.route("/")
def index():
    limit_date = (datetime.now() - timedelta(days=35)).strftime("%Y-%m-%d")
    Data.query.filter(Data.date <limit_date).delete()
    db.session.commit()

    sort = request.args.get("sort")
    teacher_search = request.args.get("teacher")
    status_filter=request.args.get("status")
    query = Data.query
    if status_filter == "mitei":
    query = query.filter(Data.status == "未定")

    if teacher_search:
    query = query.filter(Data.teacher.contains(teacher_search))

    if sort == "teacher":
    records = query.order_by(Data.teacher, Data.date, Data.period).all()

    elif sort == "status":
    records = query.order_by(Data.status.desc(), Data.date, Data.period).all()

    else:
        if session.get("today_only"):
        today_str = datetime.now().strftime("%Y-%m-%d")
        records = query.filter_by(date=today_str)\
            .order_by(Data.date, Data.period).all()
        else:
        records = query.order_by(Data.status.desc(), Data.date,Data.period).all()

        grouped = defaultdict(list)

    for r in records:
    key = r.date
    grouped[key].append(r)

    return render_template("index.html", grouped=grouped)


# ---------------- 今日だけ ----------------



# ---------------- 追加 ----------------
@app.route("/add", methods=["POST"])
def add():
    periods = request.form.getlist("period")

    for p in periods:
        new = Data(
            branch=request.form["branch"],
            date=request.form["date"],
            period=p,
            time=request.form["time"],
            teacher=request.form["teacher"],
            status=request.form["status"],
            sub_teacher=request.form.get("sub_teacher"),
            note=request.form["note"]
        )
        db.session.add(new)

    db.session.commit()
    return redirect("/")


# ---------------- 編集 ----------------
@app.route("/edit/<int:id>")
def edit(id):
    target = Data.query.get(id)
    return render_template("edit.html", data=target)


# ---------------- 更新 ----------------
@app.route("/update/<int:id>", methods=["POST"])
def update(id):
    periods=request.form.getlist("period")
    old=Data.query.get(id)
    db.session.delete(old)
    for p in periods:

        data = Data(
        branch=branch,
        date=date,
        period=p,
        time=time,
        teacher=teacher,
        status=status,
        sub_teacher=sub_teacher,
        note=note
    )

        db.session.add(data)

    db.session.commit()
    return redirect("/")


# ---------------- 削除 ----------------
@app.route("/delete/<int:id>")
def delete(id):
    target = Data.query.get(id)
    db.session.delete(target)
    db.session.commit()
    return redirect("/")


# ---------------- Excel ----------------
@app.route("/export")
def export():
    from openpyx1.styles import Alignment, Font
    from openpyx1.utils import get_column_letter
    records = Data.query.order_by(Data.date, Data.period).all()

    wb = Workbook()
    ws = wb.active
    ws.append(["日付", "限", "時間", "先生", "状態", "代講先生", "備考"])

    for r in records:
        ws.append([
            r.date,
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
                    max_length = max(max_length,len(str(cell.value)))
            ws.column_dimension[col_letter].width = max_length + 4
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


