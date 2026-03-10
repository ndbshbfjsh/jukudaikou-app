from flask import Flask, render_template, request, redirect, send_file
from flask_sqlalchemy import SQLAlchemy
from collections import defaultdict
from datetime import datetime
from openpyxl import Workbook
import os

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv(
    "DATABASE_URL",
    "sqlite:///database.db"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)
with app.app_context():
    db.create_all()



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

# ---------------- 一覧 ----------------
@app.route("/")
def index():
    sort = request.args.get("sort")

    if sort == "teacher":
        records = Data.query.order_by(Data.teacher).all()
    elif sort == "status":
        records = Data.query.order_by(Data.status).all()
    else:
        records = Data.query.order_by(Data.date, Data.period).all()

    grouped = defaultdict(list)

    for r in records:
        key = f"{r.date}_{r.period}"
        grouped[key].append(r)

    return render_template("index.html", grouped=grouped)


# ---------------- 今日だけ ----------------
@app.route("/today")
def today():
    today_str = datetime.now().strftime("%Y-%m-%d")
    records = Data.query.filter_by(date=today_str).order_by(Data.period).all()

    grouped = defaultdict(list)
    for r in records:
        key = f"{r.date}_{r.period}"
        grouped[key].append(r)

    return render_template("index.html", grouped=grouped)


# ---------------- 追加 ----------------
@app.route("/add", methods=["POST"])
def add():
    new = Data(
    branch=request.form["branch"],   # ←これ追加
    date=request.form["date"],
    period=request.form["period"],
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
    data = Data.query.get(id)

    data.date = request.form["date"]
    data.period = request.form["period"]
    data.time = request.form["time"]
    data.teacher = request.form["teacher"]
    data.status = request.form["status"]
    data.sub_teacher = request.form.get("sub_teacher")
    data.note = request.form["note"]

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

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="daikou.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


