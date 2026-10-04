#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-приложение для расчёта пеней по ЖКУ — для ТСЖ/УК.
Flask + SQLite + логины + импорт/экспорт Excel + PDF.
"""

import os
import io
from datetime import date, datetime

from flask import (Flask, render_template, request, redirect, url_for,
                    flash, send_file, session, jsonify)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                          login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from fpdf import FPDF

from peni_engine import calculate_fifo, due_date_for_period

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-only-change-before-publication")
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL", f"sqlite:///{os.path.join(BASE_DIR, 'peni.db')}"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Пожалуйста, войдите в систему."


# ==================== МОДЕЛИ ====================

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(200))
    position = db.Column(db.String(200))
    role = db.Column(db.String(20), default="user")  # 'admin' | 'user'
    org_id = db.Column(db.Integer, db.ForeignKey("organization.id"))

    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)

    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)


class Organization(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name_full = db.Column(db.String(300), default="ТСЖ «Образец»")
    name_short = db.Column(db.String(100), default="ТСЖ «Образец»")
    address = db.Column(db.String(300), default="г. Москва, ул. Примерная, д. 1")
    inn = db.Column(db.String(20), default="0000000000")
    ogrn = db.Column(db.String(20), default="0000000000000")
    users = db.relationship("User", backref="organization", lazy=True)


class Calculation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"))
    account_number = db.Column(db.String(100))  # лицевой счёт / ФИО
    calc_date = db.Column(db.Date)
    rate_mode = db.Column(db.String(20))
    total_debt = db.Column(db.Float)
    total_peni = db.Column(db.Float)
    total_all = db.Column(db.Float)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    details_json = db.Column(db.Text)  # сериализованные строки расчёта
    user = db.relationship("User", backref="calculations")


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ==================== АВТОРИЗАЦИЯ ====================

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for("index"))
        flash("Неверный логин или пароль", "error")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# ==================== ОСНОВНЫЕ СТРАНИЦЫ ====================

@app.route("/")
@login_required
def index():
    recent = Calculation.query.filter_by(user_id=current_user.id)\
        .order_by(Calculation.created_at.desc()).limit(20).all()
    return render_template("index.html", recent=recent, org=current_user.organization,
                           today_iso=date.today().isoformat())


def _parse_ru_date(s):
    # Excel/openpyxl may return date/datetime objects directly.
    if isinstance(s, datetime):
        return s.date()
    if isinstance(s, date):
        return s
    text = str(s).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%m.%Y", "%Y-%m"):
        try:
            parsed = datetime.strptime(text, fmt).date()
            if fmt in ("%m.%Y", "%Y-%m"):
                return parsed.replace(day=1)
            return parsed
        except ValueError:
            continue
    raise ValueError(f"Не удалось разобрать дату: {text}")


def _parse_ru_number(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    return float(s) if s else 0.0


@app.route("/calculate", methods=["POST"])
@login_required
def calculate():
    """Расчёт по данным, введённым вручную через форму (JSON)."""
    data = request.get_json()
    account = data.get("account", "")
    calc_date_str = data.get("calc_date")
    rate_mode = data.get("rate_mode", "historical")
    kapremont = data.get("kapremont", False)

    calc_date = _parse_ru_date(calc_date_str) if calc_date_str else date.today()

    charges = []
    for row in data.get("charges", []):
        period = _parse_ru_date(row["period"])
        amount = _parse_ru_number(row["amount"])
        deadline = due_date_for_period(period) if not row.get("deadline") else _parse_ru_date(row["deadline"])
        charges.append({"period": period, "amount": amount, "deadline": deadline})

    payments = []
    for row in data.get("payments", []):
        payments.append({"date": _parse_ru_date(row["date"]), "amount": _parse_ru_number(row["amount"])})

    result = calculate_fifo(charges, payments, calc_date, kapremont=kapremont, rate_mode=rate_mode)

    # сохраняем в историю
    import json
    calc = Calculation(
        user_id=current_user.id, account_number=account, calc_date=calc_date,
        rate_mode=rate_mode, total_debt=result["total_debt"],
        total_peni=result["total_peni"], total_all=result["total_all"],
        details_json=json.dumps([
            {"period": r["period"].isoformat(), "deadline": r["deadline"].isoformat(),
             "original": r["original"], "remaining": r["remaining"], "peni": r["peni"]}
            for r in result["rows"]
        ], ensure_ascii=False)
    )
    db.session.add(calc)
    db.session.commit()

    return jsonify({
        "calc_id": calc.id,
        "total_debt": result["total_debt"],
        "total_peni": result["total_peni"],
        "total_all": result["total_all"],
        "rows": [
            {"period": r["period"].strftime("%m.%Y"), "deadline": r["deadline"].strftime("%d.%m.%Y"),
             "original": r["original"], "remaining": r["remaining"], "peni": r["peni"]}
            for r in result["rows"]
        ]
    })


@app.route("/import_excel", methods=["POST"])
@login_required
def import_excel():
    """Массовый импорт долгов/платежей из Excel-файла."""
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Файл не выбран"}), 400

    wb = openpyxl.load_workbook(file, data_only=True)

    charges = []
    if "Долги" in wb.sheetnames:
        ws = wb["Долги"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            period = _parse_ru_date(row[0])
            amount = _parse_ru_number(row[1])
            # Для пересчёта с историей платежей импортируем полные начисления,
            # а платежи распределяем отдельно. Использовать текущий остаток вместе
            # с историей платежей нельзя — получится двойное уменьшение долга.
            amount_to_use = amount
            if amount_to_use <= 0:
                continue
            deadline = due_date_for_period(period)
            if len(row) > 3 and row[3]:
                try:
                    deadline = _parse_ru_date(row[3])
                except Exception:
                    pass
            charges.append({"period": period, "amount": amount_to_use, "deadline": deadline})

    payments = []
    if "Платежи" in wb.sheetnames:
        ws = wb["Платежи"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            payments.append({"date": _parse_ru_date(row[0]), "amount": _parse_ru_number(row[1])})

    return jsonify({
        "charges": [{"period": c["period"].strftime("%d.%m.%Y"), "amount": c["amount"],
                     "deadline": c["deadline"].strftime("%d.%m.%Y")} for c in charges],
        "payments": [{"date": p["date"].strftime("%d.%m.%Y"), "amount": p["amount"]} for p in payments]
    })


# ==================== ЭКСПОРТ ====================

def _build_pdf(calc, rows, org, user):
    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    font_bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

    pdf = FPDF()
    pdf.add_font("DejaVu", "", font_path)
    pdf.add_font("DejaVu", "B", font_bold)
    pdf.add_page()

    pdf.set_font("DejaVu", "B", 13)
    pdf.cell(0, 8, "РАСЧЁТ ПЕНИ", ln=True, align="C")
    pdf.cell(0, 6, "по задолженности за жилищно-коммунальные услуги", ln=True, align="C")
    pdf.ln(4)

    pdf.set_font("DejaVu", "B", 10)
    pdf.cell(0, 6, org.name_full, ln=True)
    pdf.set_font("DejaVu", "", 9)
    pdf.cell(0, 5, f"Адрес: {org.address}", ln=True)
    pdf.cell(0, 5, f"ИНН: {org.inn}   ОГРН: {org.ogrn}", ln=True)
    pdf.ln(2)
    pdf.cell(0, 5, f"Лицевой счёт / должник: {calc.account_number or '—'}", ln=True)
    pdf.cell(0, 5, f"Дата расчёта: {calc.calc_date.strftime('%d.%m.%Y')}", ln=True)
    mode_label = "историческая ставка ЦБ РФ" if calc.rate_mode == "historical" else "фиксированная ставка 9,5%"
    pdf.cell(0, 5, f"Метод расчёта: {mode_label}, ПП РФ № 474/329", ln=True)
    pdf.ln(4)

    # Таблица
    pdf.set_font("DejaVu", "B", 8.5)
    widths = [22, 24, 24, 26, 18, 26, 28]
    headers = ["Период", "Начислено,₽", "Срок опл.", "Остаток,₽", "Дней", "Пени,₽", "К оплате,₽"]
    for w, h in zip(widths, headers):
        pdf.cell(w, 7, h, border=1, align="C")
    pdf.ln()

    pdf.set_font("DejaVu", "", 8)
    for r in rows:
        if r["remaining"] <= 0 and r["peni"] <= 0:
            continue
        days = (calc.calc_date - r["deadline"]).days if hasattr(r["deadline"], "year") else 0
        vals = [
            r["period"].strftime("%m.%Y") if hasattr(r["period"], "year") else str(r["period"]),
            f"{r['original']:,.2f}",
            r["deadline"].strftime("%d.%m.%Y") if hasattr(r["deadline"], "year") else str(r["deadline"]),
            f"{r['remaining']:,.2f}",
            str(days),
            f"{r['peni']:,.2f}",
            f"{r['remaining'] + r['peni']:,.2f}",
        ]
        for w, v in zip(widths, vals):
            pdf.cell(w, 6, v, border=1, align="R" if v != vals[0] else "C")
        pdf.ln()

    pdf.ln(3)
    pdf.set_font("DejaVu", "B", 10)
    pdf.cell(0, 6, f"Итого основной долг: {calc.total_debt:,.2f} ₽", ln=True)
    pdf.cell(0, 6, f"Итого пени: {calc.total_peni:,.2f} ₽", ln=True)
    pdf.cell(0, 7, f"ВСЕГО К ОПЛАТЕ: {calc.total_all:,.2f} ₽", ln=True)
    pdf.ln(8)

    pdf.set_font("DejaVu", "", 9)
    pdf.cell(0, 6, f"Ответственное лицо: {user.position or ''} {user.full_name or ''}", ln=True)
    pdf.ln(10)
    pdf.cell(60, 6, "Подпись: _____________________", ln=False)
    pdf.cell(0, 6, f"Дата: {date.today().strftime('%d.%m.%Y')}", ln=True)

    pdf.ln(6)
    pdf.set_font("DejaVu", "", 7.5)
    pdf.multi_cell(0, 4,
        "Расчёт выполнен в соответствии с ч. 14, 14.1 ст. 155 Жилищного кодекса РФ с применением "
        "ключевой ставки Банка России и учётом особенностей, установленных постановлениями "
        "Правительства РФ № 474 от 26.03.2022 и № 329 от 18.03.2025 (действуют до 01.01.2027). "
        "Платежи распределены по методу FIFO — в первую очередь гасится наиболее ранняя задолженность.")

    return pdf.output(dest="S")


@app.route("/export/pdf/<int:calc_id>")
@login_required
def export_pdf(calc_id):
    import json
    calc = Calculation.query.get_or_404(calc_id)
    rows_raw = json.loads(calc.details_json)
    rows = [{"period": _parse_ru_date(r["period"]), "deadline": _parse_ru_date(r["deadline"]),
             "original": r["original"], "remaining": r["remaining"], "peni": r["peni"]}
            for r in rows_raw]

    pdf_bytes = _build_pdf(calc, rows, current_user.organization, current_user)
    buf = io.BytesIO(bytes(pdf_bytes))
    buf.seek(0)
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                      download_name=f"raschet_peni_{calc_id}.pdf")


@app.route("/export/excel/<int:calc_id>")
@login_required
def export_excel(calc_id):
    import json
    calc = Calculation.query.get_or_404(calc_id)
    rows_raw = json.loads(calc.details_json)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Расчёт пеней"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F6FEB", end_color="1F6FEB", fill_type="solid")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:G1")
    ws["A1"] = "Расчёт пени по задолженности за ЖКУ"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A3"] = "Лицевой счёт:"
    ws["B3"] = calc.account_number or "—"
    ws["A4"] = "Дата расчёта:"
    ws["B4"] = calc.calc_date.strftime("%d.%m.%Y")

    headers = ["Период", "Начислено,₽", "Срок оплаты", "Остаток,₽", "Дней просрочки", "Пени,₽", "К оплате,₽"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=6, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.border = border
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    row_i = 7
    for r in rows_raw:
        if r["remaining"] <= 0 and r["peni"] <= 0:
            continue
        d = _parse_ru_date(r["deadline"])
        p = _parse_ru_date(r["period"])
        days = (calc.calc_date - d).days
        ws.cell(row=row_i, column=1, value=p.strftime("%m.%Y")).border = border
        ws.cell(row=row_i, column=2, value=r["original"]).border = border
        ws.cell(row=row_i, column=3, value=d.strftime("%d.%m.%Y")).border = border
        ws.cell(row=row_i, column=4, value=r["remaining"]).border = border
        ws.cell(row=row_i, column=5, value=days).border = border
        ws.cell(row=row_i, column=6, value=r["peni"]).border = border
        ws.cell(row=row_i, column=7, value=round(r["remaining"] + r["peni"], 2)).border = border
        row_i += 1

    row_i += 1
    ws.cell(row=row_i, column=1, value="ИТОГО:").font = Font(bold=True)
    ws.cell(row=row_i, column=4, value=calc.total_debt).font = Font(bold=True)
    ws.cell(row=row_i, column=6, value=calc.total_peni).font = Font(bold=True)
    ws.cell(row=row_i, column=7, value=calc.total_all).font = Font(bold=True)

    for col, w in zip("ABCDEFG", [12, 14, 13, 14, 16, 14, 14]):
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      as_attachment=True, download_name=f"raschet_peni_{calc_id}.xlsx")


# ==================== МАССОВЫЙ РАСЧЁТ ПО РЕЕСТРУ ====================

def _norm_header(value):
    return "".join(ch.lower() for ch in str(value or "") if ch.isalnum())


def _find_col(headers, aliases, required=True):
    normalized = {_norm_header(v): i for i, v in enumerate(headers)}
    for alias in aliases:
        key = _norm_header(alias)
        if key in normalized:
            return normalized[key]
    if required:
        raise ValueError("Не найдена обязательная колонка: " + " / ".join(aliases))
    return None


def _read_batch_rows(ws, kind):
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        return []
    account_col = _find_col(header, ["Лицевой счёт", "Лицевой счет", "ЛС", "Счёт", "Счет", "Account"])
    if kind == "charges":
        period_col = _find_col(header, ["Период", "Расчётный период", "Расчетный период"])
        amount_col = _find_col(header, ["Начисление", "Сумма начисления", "Сумма", "Начислено"])
        deadline_col = _find_col(header, ["Срок оплаты", "Дата дедлайна", "Дедлайн"], required=False)
        debt_type_col = _find_col(header, ["Вид долга", "Тип долга", "Услуга"], required=False)
    else:
        period_col = None
        amount_col = _find_col(header, ["Сумма платежа", "Сумма", "Платёж", "Платеж"])
        deadline_col = _find_col(header, ["Дата платежа", "Дата" ])
        debt_type_col = None

    parsed = []
    for line_no, row in enumerate(rows, start=2):
        if not row or account_col >= len(row) or not row[account_col]:
            continue
        account = str(row[account_col]).strip()
        amount_value = row[amount_col] if amount_col < len(row) else None
        amount = _parse_ru_number(amount_value)
        if amount <= 0:
            continue
        if kind == "charges":
            period = _parse_ru_date(row[period_col])
            due = _parse_ru_date(row[deadline_col]) if deadline_col is not None and deadline_col < len(row) and row[deadline_col] else due_date_for_period(period)
            debt_type = str(row[debt_type_col] or "ЖКУ").strip() if debt_type_col is not None and debt_type_col < len(row) else "ЖКУ"
            parsed.append({"account": account, "period": period, "amount": amount, "deadline": due, "debt_type": debt_type, "source_row": line_no})
        else:
            paid_date = _parse_ru_date(row[deadline_col])
            parsed.append({"account": account, "date": paid_date, "amount": amount, "source_row": line_no})
    return parsed


@app.route("/batch_calculate", methods=["POST"])
@login_required
def batch_calculate():
    """Считает все лицевые счета из нормализованного Excel-реестра и возвращает Excel-отчёт."""
    file = request.files.get("file")
    if not file or not file.filename.lower().endswith(".xlsx"):
        return "Загрузите файл .xlsx", 400
    try:
        wb_in = openpyxl.load_workbook(file, data_only=True, read_only=True)
        charge_sheet = next((wb_in[n] for n in ("Начисления", "Долги") if n in wb_in.sheetnames), None)
        payment_sheet = next((wb_in[n] for n in ("Платежи", "Платежи") if n in wb_in.sheetnames), None)
        if charge_sheet is None or payment_sheet is None:
            return "В книге нужны листы «Начисления» (или «Долги») и «Платежи».", 400
        charges = _read_batch_rows(charge_sheet, "charges")
        payments = _read_batch_rows(payment_sheet, "payments")
        if not charges:
            return "На листе начислений нет строк с суммами больше нуля.", 400
        calc_date = _parse_ru_date(request.form.get("calc_date") or date.today())
        rate_mode = request.form.get("rate_mode", "historical")
        if rate_mode not in ("historical", "fixed95"):
            rate_mode = "historical"

        by_account = {}
        for c in charges:
            by_account.setdefault(c["account"], {"charges": [], "payments": [], "types": set()})
            by_account[c["account"]]["charges"].append({"period": c["period"], "amount": c["amount"], "deadline": c["deadline"]})
            by_account[c["account"]]["types"].add(c["debt_type"].strip().lower())
        for p in payments:
            by_account.setdefault(p["account"], {"charges": [], "payments": [], "types": set()})
            by_account[p["account"]]["payments"].append({"date": p["date"], "amount": p["amount"]})

        wb_out = openpyxl.Workbook()
        summary = wb_out.active
        summary.title = "Сводка"
        detail = wb_out.create_sheet("Детализация")
        summary.append(["Массовый расчёт пеней по ЖКУ", "", "", "Дата расчёта", calc_date.strftime("%d.%m.%Y")])
        summary.append(["Организация", current_user.organization.name_full if current_user.organization else "ТСЖ/УК"])
        summary.append(["Метод ставки", "Историческая ставка ЦБ" if rate_mode == "historical" else "Фиксированная 9,5% (только сверка)"])
        summary.append([])
        summary.append(["Лицевой счёт", "Основной долг, ₽", "Пени, ₽", "Итого, ₽", "Периодов", "Платежей", "Примечание"])
        detail.append(["Лицевой счёт", "Период", "Начислено, ₽", "Срок оплаты", "Остаток долга, ₽", "Дней просрочки", "Пени, ₽", "К оплате, ₽"])
        summary_results = []
        for account in sorted(by_account):
            entry = by_account[account]
            if len(entry["types"]) > 1:
                raise ValueError(f"ЛС {account}: смешаны виды долга. На данном этапе укажите один вид долга на лицевой счёт.")
            debt_type = next(iter(entry["types"]), "жку")
            is_kapremont = any(term in debt_type for term in ("капрем", "капитальн"))
            result = calculate_fifo(entry["charges"], entry["payments"], calc_date,
                                   kapremont=is_kapremont, rate_mode=rate_mode)
            summary.append([account, result["total_debt"], result["total_peni"], result["total_all"],
                            len(entry["charges"]), len(entry["payments"]), debt_type])
            summary_results.append(result)
            for r in result["rows"]:
                days = (calc_date - r["deadline"]).days
                detail.append([account, r["period"].strftime("%m.%Y"), r["original"],
                               r["deadline"].strftime("%d.%m.%Y"), r["remaining"], days,
                               r["peni"], round(r["remaining"] + r["peni"], 2)])
            # Persist each account as a separate history entry.
            import json
            calc = Calculation(user_id=current_user.id, account_number=account, calc_date=calc_date,
                               rate_mode=rate_mode, total_debt=result["total_debt"],
                               total_peni=result["total_peni"], total_all=result["total_all"],
                               details_json=json.dumps([
                                   {"period": r["period"].isoformat(), "deadline": r["deadline"].isoformat(),
                                    "original": r["original"], "remaining": r["remaining"], "peni": r["peni"]}
                                   for r in result["rows"]], ensure_ascii=False))
            db.session.add(calc)
        db.session.commit()

        # Totals + formatting.
        end = summary.max_row + 1
        summary.cell(end, 1, "ИТОГО ПО РЕЕСТРУ")
        summary.cell(end, 2, round(sum(r["total_debt"] for r in summary_results), 2))
        summary.cell(end, 3, round(sum(r["total_peni"] for r in summary_results), 2))
        summary.cell(end, 4, round(sum(r["total_all"] for r in summary_results), 2))
        for ws in (summary, detail):
            ws.freeze_panes = "A6" if ws is summary else "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F6FEB")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for col in range(1, ws.max_column + 1):
                letter = openpyxl.utils.get_column_letter(col)
                ws.column_dimensions[letter].width = min(max(14, max((len(str(ws.cell(r, col).value or "")) for r in range(1, ws.max_row + 1)), default=12) + 2), 34)
        for ws in (summary, detail):
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, (int, float)) and cell.column > 1:
                        cell.number_format = '#,##0.00'
        output = io.BytesIO()
        wb_out.save(output)
        output.seek(0)
        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name="massovyi_raschet_peni.xlsx")
    except Exception as exc:
        db.session.rollback()
        return f"Ошибка обработки реестра: {exc}", 400


# ==================== ИНИЦИАЛИЗАЦИЯ БД ====================

def init_db():
    with app.app_context():
        db.create_all()
        if not Organization.query.first():
            org = Organization(
                name_full="Товарищество собственников жилья «Образец»",
                name_short="ТСЖ «Образец»",
                address="г. Москва, ул. Примерная, д. 1",
                inn="7700000000", ogrn="1157700000000"
            )
            db.session.add(org)
            db.session.commit()

            admin = User(username="admin", full_name="Иванов Иван Иванович",
                         position="Главный бухгалтер", role="admin", org_id=org.id)
            admin.set_password("admin123")
            db.session.add(admin)
            db.session.commit()
            print("Создан пользователь: admin / admin123")


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
