#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-приложение для расчёта пеней по ЖКУ — ТСЖ/УК.
Flask + SQLite + вход по логину + импорт/экспорт Excel + PDF + массовый расчёт.

Отчёт массового расчёта содержит пять листов: Сводка, Детализация, Разбивка,
Платежи, Контроль — так, чтобы каждую цифру можно было пересчитать вручную.
"""

import os
import io
import json
from datetime import date, datetime

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, send_file, jsonify)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from fpdf import FPDF

from peni_engine import (calculate_fifo, batch_calculate as engine_batch,
                         check_reconciliation, due_date_for_period,
                         RATE_MODE_LABELS, RATE_MODES)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-only-change-before-publication")
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL", f"sqlite:///{os.path.join(BASE_DIR, 'peni.db')}"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Пожалуйста, войдите в систему."

DEFAULT_RATE_MODE = "fixed95"


# ==================== МОДЕЛИ ====================

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(200))
    position = db.Column(db.String(200))
    role = db.Column(db.String(20), default="user")
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
    account_number = db.Column(db.String(100))
    calc_date = db.Column(db.Date)
    rate_mode = db.Column(db.String(30))
    total_debt = db.Column(db.Float)
    total_peni = db.Column(db.Float)
    total_all = db.Column(db.Float)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    details_json = db.Column(db.Text)
    user = db.relationship("User", backref="calculations")


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ==================== ВСПОМОГАТЕЛЬНОЕ ====================

def _parse_ru_date(s):
    if isinstance(s, datetime):
        return s.date()
    if isinstance(s, date):
        return s
    text = str(s).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y", "%m.%Y", "%Y-%m"):
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


def _round(x, n=2):
    return round(float(x), n)


def serialize_result(result):
    """Сериализует результат расчёта одного лицевого счёта для хранения."""
    rows = []
    for r in result["rows"]:
        rows.append({
            "account": result.get("account", ""),
            "period": r["period"].isoformat(),
            "deadline": r["deadline"].isoformat(),
            "original": r["original"], "remaining": r["remaining"],
            "paid_total": r["paid_total"],
            "advance_paid": r.get("advance_paid", 0.0),
            "days_total": r["days_total"], "days_peni": r["days_peni"],
            "peni": r["peni"], "k_pay": r["k_pay"], "note": r["note"],
            "rates_used": r["rates_used"],
            "segments": [{**s, "start": s["start"].isoformat(), "end": s["end"].isoformat()}
                         for s in r["segments"]],
            "payments": [{**p, "date": p["date"].isoformat()} for p in r["payments"]],
        })
    return {
        "account": result.get("account", ""),
        "rate_mode": result["rate_mode"], "kapremont": result["kapremont"],
        "calc_date": result["calc_date"].isoformat(),
        "total_debt": result["total_debt"], "total_peni": result["total_peni"],
        "total_all": result["total_all"], "control": result["control"],
        "advance_total": round(sum(r.get("advance_paid", 0.0) for r in result["rows"]), 2),
        "overpayment": result.get("overpayment", 0.0),
        "movements": [{"date": m["date"].isoformat(), "amount": m["amount"],
                       "unallocated": m["unallocated"]}
                      for m in result.get("movements", [])],
        "rows": rows,
    }


def deserialize_result(data):
    """Обратное преобразование для экспорта PDF/Excel.

    Совместимо с расчётами, сохранёнными старой версией: если разбивки нет,
    она считается пустой, а недостающие поля заполняются значениями по умолчанию.
    """
    rows = []
    for r in data.get("rows", []):
        original = r.get("original", 0.0)
        remaining = r.get("remaining", 0.0)
        rows.append({
            "period": _parse_ru_date(r["period"]),
            "deadline": _parse_ru_date(r["deadline"]),
            "original": original,
            "remaining": remaining,
            "paid_total": r.get("paid_total", _round(original - remaining)),
            "advance_paid": r.get("advance_paid", 0.0),
            "days_total": r.get("days_total", 0),
            "days_peni": r.get("days_peni", 0),
            "peni": r.get("peni", 0.0),
            "k_pay": r.get("k_pay", _round(remaining + r.get("peni", 0.0))),
            "note": r.get("note", ""),
            "rates_used": r.get("rates_used", []),
            "segments": [{**s, "start": _parse_ru_date(s["start"]), "end": _parse_ru_date(s["end"])}
                         for s in r.get("segments", [])],
            "payments": [{**p, "date": _parse_ru_date(p["date"])} for p in r.get("payments", [])],
        })
    rows.sort(key=lambda x: x["deadline"])
    movements = [{**m, "date": _parse_ru_date(m["date"])} for m in data.get("movements", [])]
    return {**data, "rows": rows, "movements": movements,
            "overpayment": data.get("overpayment", 0.0),
            "advance_total": data.get("advance_total", 0.0),
            "calc_date": _parse_ru_date(data["calc_date"]),
            "rate_mode": data.get("rate_mode", DEFAULT_RATE_MODE),
            "account": data.get("account", "")}


# ==================== СТИЛИ EXCEL ====================

BLUE = "1F6FEB"
GREEN = "1E7B45"
AMBER = "9A6400"
GREY = "F2F5F8"
THIN = Side(style="thin", color="C9D4E0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _header_row(ws, values, row=1, fill=BLUE):
    for col, v in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=v)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def _autosize(ws, max_width=42):
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        longest = 0
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, col).value
            if v is None:
                continue
            longest = max(longest, max(len(part) for part in str(v).split("\n")))
        ws.column_dimensions[letter].width = min(max(11, longest + 2), max_width)


def _money_format(ws, first_col=2):
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, (int, float)) and c.column >= first_col:
                c.number_format = '#,##0.00'
                c.border = BORDER
            elif isinstance(c.value, str):
                c.border = BORDER


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


@app.route("/")
@login_required
def index():
    recent = (Calculation.query.filter_by(user_id=current_user.id)
              .order_by(Calculation.created_at.desc()).limit(20).all())
    return render_template("index.html", recent=recent, org=current_user.organization,
                           today_iso=date.today().isoformat(),
                           rate_modes=RATE_MODES, rate_labels=RATE_MODE_LABELS)


# ==================== РАСЧЁТ (ОДИН ЛИЦЕВОЙ СЧЁТ) ====================

@app.route("/calculate", methods=["POST"])
@login_required
def calculate():
    data = request.get_json()
    account = data.get("account", "")
    calc_date = _parse_ru_date(data["calc_date"]) if data.get("calc_date") else date.today()
    rate_mode = data.get("rate_mode") or DEFAULT_RATE_MODE
    kapremont = bool(data.get("kapremont", False))

    charges, payments = [], []
    for row in data.get("charges", []):
        period = _parse_ru_date(row["period"])
        deadline = _parse_ru_date(row["deadline"]) if row.get("deadline") else due_date_for_period(period)
        charges.append({"period": period, "amount": _parse_ru_number(row["amount"]), "deadline": deadline})
    for row in data.get("payments", []):
        payments.append({"date": _parse_ru_date(row["date"]), "amount": _parse_ru_number(row["amount"])})

    result = calculate_fifo(charges, payments, calc_date, kapremont=kapremont, rate_mode=rate_mode)
    result["account"] = account

    payload = serialize_result(result)
    calc = Calculation(user_id=current_user.id, account_number=account, calc_date=calc_date,
                       rate_mode=rate_mode, total_debt=result["total_debt"],
                       total_peni=result["total_peni"], total_all=result["total_all"],
                       details_json=json.dumps(payload, ensure_ascii=False))
    db.session.add(calc)
    db.session.commit()

    return jsonify({
        "calc_id": calc.id,
        "total_debt": result["total_debt"], "total_peni": result["total_peni"],
        "total_all": result["total_all"],
        "advance_total": result.get("advance_total", 0.0),
        "overpayment": result.get("overpayment", 0.0),
        "control": result["control"],
        "rate_mode_label": RATE_MODE_LABELS.get(rate_mode, rate_mode),
        "rows": [{
            "period": r["period"].strftime("%m.%Y"),
            "deadline": r["deadline"].strftime("%d.%m.%Y"),
            "original": r["original"], "remaining": r["remaining"],
            "paid_total": r["paid_total"],
            "advance_paid": r.get("advance_paid", 0.0),
            "days_total": r["days_total"], "days_peni": r["days_peni"],
            "peni": r["peni"], "k_pay": r["k_pay"], "note": r["note"],
            "segments": [{"start": s["start"].strftime("%d.%m.%Y"),
                          "end": s["end"].strftime("%d.%m.%Y"),
                          "days": s["days"], "base": s["base"],
                          "dividend": s["dividend"], "rate": s["rate"],
                          "peni": round(s["peni"], 2)} for s in r["segments"]],
            "payments": [{"date": p["date"].strftime("%d.%m.%Y"),
                          "paid": p["paid"]} for p in r["payments"]],
        } for r in result["rows"]],
    })


# ==================== ИМПОРТ EXCEL (ОДИН СЧЁТ) ====================

@app.route("/import_excel", methods=["POST"])
@login_required
def import_excel():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Файл не выбран"}), 400
    wb = openpyxl.load_workbook(file, data_only=True)

    charges, payments = [], []
    sheet = next((wb[n] for n in ("Начисления", "Долги") if n in wb.sheetnames), None)
    if sheet is not None:
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, None) or []
        idx = {_norm(h): i for i, h in enumerate(header)}
        period_i = _pick(idx, ["период", "расчетныйпериод"])
        amount_i = _pick(idx, ["начисление", "сумманачисления", "сумма", "начислено"])
        due_i = _pick(idx, ["срокоплаты", "датадедлайна", "дедлайн"], optional=True)
        for row in rows:
            if not row or not row[0]:
                continue
            period = _parse_ru_date(row[period_i] if period_i is not None else row[0])
            amount = _parse_ru_number(row[amount_i] if amount_i is not None else row[1])
            if amount <= 0:
                continue
            due = (_parse_ru_date(row[due_i]) if due_i is not None and due_i < len(row)
                   and row[due_i] else due_date_for_period(period))
            charges.append({"period": period, "amount": amount, "deadline": due})

    psheet = wb["Платежи"] if "Платежи" in wb.sheetnames else None
    if psheet is not None:
        rows = psheet.iter_rows(values_only=True)
        header = next(rows, None) or []
        idx = {_norm(h): i for i, h in enumerate(header)}
        date_i = _pick(idx, ["датаплатежа", "дата"])
        amount_i = _pick(idx, ["суммаплатежа", "сумма", "платеж"])
        for row in rows:
            if not row or not row[0]:
                continue
            d = _parse_ru_date(row[date_i] if date_i is not None else row[0])
            a = _parse_ru_number(row[amount_i] if amount_i is not None else row[1])
            if a > 0:
                payments.append({"date": d, "amount": a})

    return jsonify({
        "charges": [{"period": c["period"].strftime("%d.%m.%Y"), "amount": c["amount"],
                     "deadline": c["deadline"].strftime("%d.%m.%Y")} for c in charges],
        "payments": [{"date": p["date"].strftime("%d.%m.%Y"), "amount": p["amount"]} for p in payments],
    })


def _norm(v):
    """Нормализует заголовок: нижний регистр, только буквы и цифры.
    Буква «ё» приводится к «е», чтобы «Лицевой счёт» и «Лицевой счет» совпадали."""
    text = str(v or "").lower().replace("ё", "е")
    return "".join(ch for ch in text if ch.isalnum())


def _pick(idx, aliases, optional=False):
    for a in aliases:
        if a in idx:
            return idx[a]
    if optional:
        return None
    raise ValueError("Не найдена колонка: " + " / ".join(aliases))


# ==================== МАССОВЫЙ РАСЧЁТ ====================

def _read_sheet(ws, kind):
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        return []
    idx = {_norm(h): i for i, h in enumerate(header)}
    acc_i = _pick(idx, ["лицевойсчет", "лс", "account"])
    if kind == "charges":
        period_i = _pick(idx, ["период", "расчетныйпериод"])
        amount_i = _pick(idx, ["начисление", "сумманачисления", "сумма", "начислено"])
        due_i = _pick(idx, ["срокоплаты", "датадедлайна", "дедлайн"], optional=True)
        type_i = _pick(idx, ["виддолга", "типдолга", "услуга"], optional=True)
    else:
        amount_i = _pick(idx, ["суммаплатежа", "сумма", "платеж"])
        date_i = _pick(idx, ["датаплатежа", "дата"])

    out = []
    for line, row in enumerate(rows, start=2):
        if not row or acc_i >= len(row) or not row[acc_i]:
            continue
        account = str(row[acc_i]).strip()
        amount = _parse_ru_number(row[amount_i] if amount_i < len(row) else None)
        if amount <= 0:
            continue
        if kind == "charges":
            period = _parse_ru_date(row[period_i])
            due = (_parse_ru_date(row[due_i]) if due_i is not None and due_i < len(row)
                   and row[due_i] else due_date_for_period(period))
            dtype = (str(row[type_i]).strip() if type_i is not None and type_i < len(row)
                     and row[type_i] else "ЖКУ")
            out.append({"account": account, "period": period, "amount": amount,
                        "deadline": due, "debt_type": dtype, "source_row": line})
        else:
            # Дата платежа обязательна: без неё нельзя разграничить пеню
            # до оплаты и после, и расчёт невозможно пересчитать.
            raw_date = row[date_i] if date_i < len(row) else None
            if raw_date in (None, ""):
                raise ValueError(
                    f"Лист «Платежи», строка {line}: не указана дата платежа для "
                    f"лицевого счёта {account}. Заполните колонку «Дата платежа» — без неё "
                    f"пеня за период до оплаты и после оплаты не разграничивается.")
            out.append({"account": account, "date": _parse_ru_date(raw_date),
                        "amount": amount, "source_row": line})
    return out


@app.route("/batch_calculate", methods=["POST"])
@login_required
def batch_calculate():
    file = request.files.get("file")
    if not file or not file.filename.lower().endswith(".xlsx"):
        return "Загрузите файл .xlsx", 400
    try:
        wb_in = openpyxl.load_workbook(file, data_only=True, read_only=True)
        charge_sheet = next((wb_in[n] for n in ("Начисления", "Долги") if n in wb_in.sheetnames), None)
        payment_sheet = wb_in["Платежи"] if "Платежи" in wb_in.sheetnames else None
        if charge_sheet is None:
            return "В книге нужен лист «Начисления» (или «Долги»).", 400

        charges = _read_sheet(charge_sheet, "charges")
        payments = _read_sheet(payment_sheet, "payments") if payment_sheet is not None else []
        if not charges:
            return "На листе начислений нет строк с суммами больше нуля.", 400

        calc_date = _parse_ru_date(request.form.get("calc_date") or date.today())
        rate_mode = request.form.get("rate_mode") or DEFAULT_RATE_MODE
        if rate_mode not in RATE_MODES:
            rate_mode = DEFAULT_RATE_MODE
        # strict=1 — блокировать расчёт при превышении платежей над начислениями.
        # По умолчанию превышение считается авансом или переплатой: предупреждаем,
        # но расчёт выполняем.
        strict = request.form.get("strict") in ("1", "on", "true", "yes")

        groups = {}
        for c in charges:
            g = groups.setdefault(c["account"], {"charges": [], "payments": [], "types": set()})
            g["charges"].append({"period": c["period"], "amount": c["amount"], "deadline": c["deadline"]})
            g["types"].add(c["debt_type"].strip().lower())
        for p in payments:
            g = groups.setdefault(p["account"], {"charges": [], "payments": [], "types": set()})
            g["payments"].append({"date": p["date"], "amount": p["amount"]})

        for acc, g in groups.items():
            if len(g["types"]) > 1:
                return (f"Лицевой счёт {acc}: указано несколько видов долга одновременно "
                        f"({', '.join(sorted(g['types']))}). Разделите их на отдельные расчёты."), 400
            dtype = next(iter(g["types"]), "жку")
            g["kapremont"] = any(t in dtype for t in ("капрем", "капитальн"))
            g["strict"] = strict
            g.pop("types")

        # Сверка до расчёта. Превышение платежей над начислениями — это аванс
        # или переплата, расчёт продолжаем. Останавливаем только в строгом режиме
        # или при явной ошибке данных.
        blocked, warnings = [], []
        for acc in sorted(groups):
            rec = check_reconciliation(groups[acc]["charges"], groups[acc]["payments"],
                                      strict=strict)
            if not rec["ok"]:
                blocked.append(f"ЛС {acc}: {rec['message']}")
            elif rec["severity"] == "warning":
                warnings.append(f"ЛС {acc}: {rec['message']}")
        if blocked:
            return ("Расчёт остановлен — данные не согласованы:\n\n"
                    + "\n\n".join(blocked)), 400

        batch = engine_batch(groups, calc_date, rate_mode=rate_mode)
        batch_control_ok = batch["control_ok"]

        # сохраняем историю по каждому счёту
        for acc_res in batch["accounts"]:
            payload = serialize_result(acc_res)
            db.session.add(Calculation(
                user_id=current_user.id, account_number=acc_res["account"],
                calc_date=calc_date, rate_mode=rate_mode,
                total_debt=acc_res["total_debt"], total_peni=acc_res["total_peni"],
                total_all=acc_res["total_all"],
                details_json=json.dumps(payload, ensure_ascii=False)))
        db.session.commit()

        wb_out = openpyxl.Workbook()

        # ---------- Сводка ----------
        s = wb_out.active
        s.title = "Сводка"
        org = current_user.organization
        s.append(["Массовый расчёт пеней по задолженности за ЖКУ"])
        s["A1"].font = Font(bold=True, size=13)
        s.append(["Организация", org.name_full if org else "ТСЖ/УК"])
        s.append(["ИНН / ОГРН", f"{org.inn} / {org.ogrn}" if org else "—"])
        s.append(["Дата расчёта", calc_date.strftime("%d.%m.%Y")])
        s.append(["Метод ставки", RATE_MODE_LABELS.get(rate_mode, rate_mode)])
        s.append(["Порядок гашения платежей", "FIFO: платёж закрывает самый ранний непогашенный период"])
        s.append([""])   # разделительная строка: пустой список openpyxl не добавляет
        head = ["Лицевой счёт", "Начислено, ₽", "Оплачено, ₽", "в т.ч. авансом, ₽",
                "Остаток долга, ₽", "Переплата, ₽",
                "Дней просрочки (всего по периодам)", "Дней с пенями (всего по периодам)",
                "Пени, ₽", "Итого к оплате, ₽", "Ставки в расчёте, %", "Примечание"]
        _header_row(s, head, row=s.max_row + 1)
        for a in batch["accounts"]:
            paid = _round(sum(r["paid_total"] for r in a["rows"]))
            billed = _round(sum(r["original"] for r in a["rows"]))
            days_total = sum(r["days_total"] for r in a["rows"])
            days_peni = sum(r["days_peni"] for r in a["rows"])
            rates = sorted({x for r in a["rows"] for x in r["rates_used"]})
            # Примечания собираем с указанием периода: у разных периодов причины разные.
            note_items = []
            rows_with_notes = [r for r in a["rows"] if r["note"]]
            many = len(rows_with_notes) > 1
            for r in rows_with_notes:
                for part in r["note"].split(";"):
                    part = part.strip()
                    if not part:
                        continue
                    item = f"{r['period'].strftime('%m.%Y')}: {part}" if many else part
                    if item not in note_items:
                        note_items.append(item)
            notes = "; ".join(note_items)
            s.append([a["account"], billed, paid, a["advance_total"], a["total_debt"],
                      a["overpayment"], days_total, days_peni, a["total_peni"],
                      a["total_all"],
                      ", ".join(f"{x:g}" for x in rates) if rates else "—", notes])
        s.append(["ИТОГО ПО РЕЕСТРУ",
                  _round(sum(sum(r["original"] for r in a["rows"]) for a in batch["accounts"])),
                  _round(sum(sum(r["paid_total"] for r in a["rows"]) for a in batch["accounts"])),
                  batch["total_advance"], batch["total_debt"], batch["total_overpayment"],
                  "", "", batch["total_peni"], batch["total_all"], "", ""])
        for c in s[s.max_row]:
            c.font = Font(bold=True)
        _autosize(s)
        _money_format(s, 2)

        # ---------- Детализация ----------
        d = wb_out.create_sheet("Детализация")
        _header_row(d, ["Лицевой счёт", "Период", "Срок оплаты", "Начислено, ₽",
                        "Оплачено по периоду, ₽", "в т.ч. авансом, ₽", "Остаток долга, ₽",
                        "Дней просрочки", "Дней с пенями", "Пени, ₽", "Итого к оплате, ₽",
                        "Примечание"])
        for a in batch["accounts"]:
            for r in a["rows"]:
                d.append([a["account"], r["period"].strftime("%m.%Y"),
                          r["deadline"].strftime("%d.%m.%Y"), r["original"],
                          r["paid_total"], r["advance_paid"], r["remaining"],
                          r["days_total"], r["days_peni"], r["peni"], r["k_pay"],
                          r["note"]])
        d.append(["ИТОГО", "", "", "", "", "", batch["total_debt"], "", "",
                  batch["total_peni"], batch["total_all"], ""])
        for c in d[d.max_row]:
            c.font = Font(bold=True)
        d.freeze_panes = "A2"
        _autosize(d)
        _money_format(d, 4)

        # ---------- Разбивка ----------
        b = wb_out.create_sheet("Разбивка")
        b.append(["Расшифровка начисления пеней по отрезкам — каждую строку можно пересчитать"])
        b["A1"].font = Font(bold=True, size=11)
        _header_row(b, ["Лицевой счёт", "Период", "С какого дня", "По какой день",
                        "Дней в отрезке", "База для пеней, ₽", "Доля ставки",
                        "Ставка, % годовых", "Пени за отрезок, ₽"], row=3)
        for a in batch["accounts"]:
            for r in a["rows"]:
                for seg in r["segments"]:
                    b.append([a["account"], r["period"].strftime("%m.%Y"),
                              seg["start"].strftime("%d.%m.%Y"), seg["end"].strftime("%d.%m.%Y"),
                              seg["days"], seg["base"], f"1/{seg['dividend']}",
                              seg["rate"], _round(seg["peni"])])
        b.freeze_panes = "A4"
        _autosize(b)
        _money_format(b, 5)

        # ---------- Платежи ----------
        # Показываем, в какие периоды ушли деньги: один платёж может закрывать
        # несколько периодов, тогда он занимает несколько строк.
        p = wb_out.create_sheet("Платежи")
        _header_row(p, ["Лицевой счёт", "Дата платежа", "Сумма платежа, ₽",
                        "Тип платежа", "Зачислено в период", "Срок оплаты периода",
                        "Остаток периода после зачисления, ₽"])
        for a in batch["accounts"]:
            for r in a["rows"]:
                for pay in r["payments"]:
                    p.append([a["account"], pay["date"].strftime("%d.%m.%Y"),
                              pay["paid"], pay["kind"], r["period"].strftime("%m.%Y"),
                              r["deadline"].strftime("%d.%m.%Y"), pay["remaining_after"]])
            # Неразнесённый остаток платежей: аванс или переплата.
            for mv in a.get("movements", []):
                if mv["unallocated"] > 0.005:
                    p.append([a["account"], mv["date"].strftime("%d.%m.%Y"),
                              mv["unallocated"], "не разнесено (переплата/аванс)",
                              "—", "—", "—"])
        if p.max_row == 1:
            p.append(["—", "—", 0, "—", "—", "—", "—"])
        p.freeze_panes = "A2"
        _autosize(p)
        _money_format(p, 3)

        # ---------- Сверка ----------
        v = wb_out.create_sheet("Сверка")
        v.append(["Сверка «начислено против поступивших платежей» до расчёта пеней"])
        v["A1"].font = Font(bold=True, size=11)
        v.append(["Превышение платежей над начислениями означает, что в листе «Начисления» "
                  "указан остаток долга, а не полное начисление: суммы вычитаются дважды."])
        _header_row(v, ["Лицевой счёт", "Начислено, ₽", "Оплачено, ₽",
                        "Разница (начислено − оплачено), ₽", "Переплата, ₽",
                        "в т.ч. авансом, ₽", "Статус", "Пояснение"], row=4)
        for a in batch["accounts"]:
            rec = a["reconciliation"]
            if rec["severity"] == "error":
                status = "ОШИБКА ДАННЫХ"
            elif rec["severity"] == "warning":
                status = "аванс/переплата"
            else:
                status = "нарушений нет"
            v.append([a["account"], rec["billed"], rec["paid"], rec["diff"],
                      a["overpayment"], a["advance_total"], status, rec["message"]])
        v.append(["", "", "", "", "", "", "",
                  "все счета согласованы" if batch["reconciliation_ok"]
                  else "есть несогласованные счета"])
        for c in v[v.max_row]:
            c.font = Font(bold=True)
        v.freeze_panes = "A5"
        _autosize(v, max_width=70)
        _money_format(v, 2)

        # ---------- Контроль ----------
        k = wb_out.create_sheet("Контроль")
        _header_row(k, ["Лицевой счёт", "Период", "Пени (итог), ₽",
                        "Сумма отрезков, ₽", "Расхождение, ₽", "Статус"])
        bad = 0
        for a in batch["accounts"]:
            for r in a["rows"]:
                seg_sum = _round(sum(x["peni"] for x in r["segments"]))
                diff = _round(r["peni"] - seg_sum)
                ok = abs(diff) < 0.01
                bad += 0 if ok else 1
                k.append([a["account"], r["period"].strftime("%m.%Y"), r["peni"],
                          seg_sum, diff, "сходится" if ok else "РАСХОЖДЕНИЕ"])
        k.append(["", "", "", "", "", ""])
        k.append(["Контроль по реестру", "", "", "", "",
                  "все строки сходятся" if bad == 0 else f"расхождений: {bad}"])
        for c in k[k.max_row]:
            c.font = Font(bold=True)
        k.freeze_panes = "A2"
        _autosize(k)
        _money_format(k, 3)

        out = io.BytesIO()
        wb_out.save(out)
        out.seek(0)
        resp = send_file(out,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name="raschet_peni_reestr.xlsx")
        if warnings:
            resp.headers["X-Reconciliation-Warning"] = " | ".join(warnings)[:900]
        return resp

    except ValueError as ve:
        db.session.rollback()
        return f"Ошибка в данных реестра: {ve}", 400
    except Exception as exc:
        db.session.rollback()
        return f"Ошибка обработки реестра: {exc}", 400


# ==================== ЭКСПОРТ PDF И EXCEL (ОДИН СЧЁТ) ====================

def _build_pdf(payload, org, user):
    font = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    font_b = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

    pdf = FPDF(orientation="L")
    pdf.add_font("DejaVu", "", font)
    pdf.add_font("DejaVu", "B", font_b)
    pdf.add_page()

    pdf.set_font("DejaVu", "B", 13)
    pdf.cell(0, 8, "РАСЧЁТ ПЕНИ по задолженности за жилищно-коммунальные услуги", ln=True, align="C")
    pdf.set_font("DejaVu", "B", 10)
    pdf.cell(0, 6, org.name_full if org else "", ln=True)
    pdf.set_font("DejaVu", "", 9)
    pdf.cell(0, 5, f"Адрес: {org.address if org else '—'}    ИНН: {org.inn if org else '—'}"
                   f"    ОГРН: {org.ogrn if org else '—'}", ln=True)
    pdf.cell(0, 5, f"Лицевой счёт / должник: {payload['account'] or '—'}"
                   f"    Дата расчёта: {payload['calc_date'].strftime('%d.%m.%Y')}", ln=True)
    pdf.cell(0, 5, "Метод ставки: " + RATE_MODE_LABELS.get(payload["rate_mode"], ""), ln=True)
    pdf.cell(0, 5, "Порядок гашения платежей: FIFO (самый ранний непогашенный период)", ln=True)
    pdf.ln(3)

    widths = [18, 22, 24, 24, 24, 18, 18, 24, 24, 24]
    heads = ["Период", "Срок опл.", "Начислено", "Оплачено", "в т.ч. аванс",
             "Остаток", "Дней с пен.", "Пени", "К оплате"]
    pdf.set_font("DejaVu", "B", 8)
    for w, h in zip(widths, heads):
        pdf.cell(w, 7, h, border=1, align="C")
    pdf.ln()
    pdf.set_font("DejaVu", "", 8)
    for r in payload["rows"]:
        vals = [r["period"].strftime("%m.%Y"), r["deadline"].strftime("%d.%m.%Y"),
                f"{r['original']:,.2f}", f"{r['paid_total']:,.2f}",
                f"{r.get('advance_paid', 0.0):,.2f}", f"{r['remaining']:,.2f}",
                str(r["days_peni"]),
                f"{r['peni']:,.2f}", f"{r['k_pay']:,.2f}"]
        for w, v in zip(widths, vals):
            pdf.cell(w, 6, v, border=1, align="C")
        pdf.ln()
    pdf.ln(2)
    pdf.set_font("DejaVu", "B", 10)
    total_billed = sum(r["original"] for r in payload["rows"])
    total_paid = sum(r["paid_total"] for r in payload["rows"])
    total_adv = sum(r.get("advance_paid", 0.0) for r in payload["rows"])
    pdf.cell(0, 6, f"Начислено: {total_billed:,.2f} ₽    Оплачено: {total_paid:,.2f} ₽    "
                   f"в том числе авансом (до срока оплаты): {total_adv:,.2f} ₽    "
                   f"Остаток долга: {payload['total_debt']:,.2f} ₽", ln=True)
    over = payload.get("overpayment", 0.0)
    if over > 0:
        pdf.set_font("DejaVu", "", 9)
        pdf.cell(0, 6, f"Переплата (оплачено больше начисленного): {over:,.2f} ₽ — "
                       f"по периодам не распределена и в пени не участвует", ln=True)
        pdf.set_font("DejaVu", "B", 10)
    pdf.cell(0, 6, f"Пени: {payload['total_peni']:,.2f} ₽    "
                   f"ВСЕГО К ОПЛАТЕ: {payload['total_all']:,.2f} ₽", ln=True)

    # расшифровка
    if any(r["segments"] for r in payload["rows"]):
        pdf.ln(4)
        pdf.set_font("DejaVu", "B", 9)
        pdf.cell(0, 6, "Расшифровка начисления пеней по отрезкам", ln=True)
        w2 = [20, 26, 26, 18, 24, 18, 22, 24]
        pdf.set_font("DejaVu", "B", 7.5)
        for w, h in zip(w2, ["Период", "С", "По", "Дней", "База", "Доля", "Ставка", "Пени"]):
            pdf.cell(w, 6, h, border=1, align="C")
        pdf.ln()
        pdf.set_font("DejaVu", "", 7.5)
        for r in payload["rows"]:
            for seg in r["segments"]:
                for w, v in zip(w2, [r["period"].strftime("%m.%Y"),
                                     seg["start"].strftime("%d.%m.%Y"),
                                     seg["end"].strftime("%d.%m.%Y"), str(seg["days"]),
                                     f"{seg['base']:,.2f}", f"1/{seg['dividend']}",
                                     f"{seg['rate']:.2f}", f"{seg['peni']:,.2f}"]):
                    pdf.cell(w, 5.5, v, border=1, align="C")
                pdf.ln()

    pdf.ln(4)
    pdf.set_font("DejaVu", "", 8)
    pdf.multi_cell(0, 4,
                   "Расчёт выполнен по ч. 14 и ч. 14.1 ст. 155 ЖК РФ с учётом особенностей, установленных "
                   "постановлениями Правительства РФ от 26.03.2022 № 474 и от 18.03.2025 № 329, а также "
                   "Федерального закона от 24.06.2025 № 177-ФЗ (срок оплаты — до 15-го числа месяца, "
                   "следующего за истекшим). Пени начисляются с 31-го дня просрочки: первые 60 дней — "
                   "1/300 ставки, далее — 1/130 (для взносов на капитальный ремонт — 1/300 на весь период).")
    pdf.ln(3)
    pdf.set_font("DejaVu", "", 9)
    pdf.cell(0, 6, f"Ответственное лицо: {user.position or ''} {user.full_name or ''}", ln=True)
    pdf.ln(6)
    pdf.cell(80, 6, "Подпись: _____________________", ln=False)
    pdf.cell(0, 6, f"Дата: {date.today().strftime('%d.%m.%Y')}", ln=True)
    return pdf.output(dest="S")


@app.route("/export/pdf/<int:calc_id>")
@login_required
def export_pdf(calc_id):
    calc = Calculation.query.get_or_404(calc_id)
    payload = deserialize_result(json.loads(calc.details_json))
    pdf_bytes = _build_pdf(payload, current_user.organization, current_user)
    buf = io.BytesIO(bytes(pdf_bytes))
    buf.seek(0)
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                     download_name=f"raschet_peni_{calc_id}.pdf")


@app.route("/export/excel/<int:calc_id>")
@login_required
def export_excel(calc_id):
    calc = Calculation.query.get_or_404(calc_id)
    payload = deserialize_result(json.loads(calc.details_json))
    org = current_user.organization

    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "Сводка"
    s.append(["Расчёт пеней по задолженности за ЖКУ"])
    s["A1"].font = Font(bold=True, size=13)
    s.append(["Организация", org.name_full if org else "—"])
    s.append(["Лицевой счёт / должник", payload["account"] or "—"])
    s.append(["Дата расчёта", payload["calc_date"].strftime("%d.%m.%Y")])
    s.append(["Метод ставки", RATE_MODE_LABELS.get(payload["rate_mode"], payload["rate_mode"])])
    s.append([])
    _header_row(s, ["Период", "Срок оплаты", "Начислено, ₽", "Оплачено, ₽",
                    "в т.ч. авансом, ₽", "Остаток долга, ₽", "Дней просрочки",
                    "Дней с пенями", "Пени, ₽", "Итого к оплате, ₽", "Примечание"],
                row=s.max_row + 1)
    for r in payload["rows"]:
        s.append([r["period"].strftime("%m.%Y"), r["deadline"].strftime("%d.%m.%Y"),
                  r["original"], r["paid_total"], r.get("advance_paid", 0.0),
                  r["remaining"], r["days_total"],
                  r["days_peni"], r["peni"], r["k_pay"], r["note"]])
    over = payload.get("overpayment", 0.0)
    if over > 0:
        s.append(["Переплата по лицевому счёту", "", "", "", "", "", "", "", "", over,
                  "оплачено больше начисленного — по периодам не распределена"])
    s.append(["ИТОГО", "", "", "", payload["total_debt"], "", "",
              payload["total_peni"], payload["total_all"], ""])
    for c in s[s.max_row]:
        c.font = Font(bold=True)
    _autosize(s)
    _money_format(s, 3)

    b = wb.create_sheet("Разбивка")
    _header_row(b, ["Период", "С какого дня", "По какой день", "Дней", "База, ₽",
                    "Доля ставки", "Ставка, %", "Пени за отрезок, ₽"])
    for r in payload["rows"]:
        for seg in r["segments"]:
            b.append([r["period"].strftime("%m.%Y"), seg["start"].strftime("%d.%m.%Y"),
                      seg["end"].strftime("%d.%m.%Y"), seg["days"], seg["base"],
                      f"1/{seg['dividend']}", seg["rate"], _round(seg["peni"])])
    b.freeze_panes = "A2"
    _autosize(b)
    _money_format(b, 5)

    p = wb.create_sheet("Платежи")
    _header_row(p, ["Период", "Дата платежа", "Сумма, ₽", "Остаток периода после зачисления, ₽"])
    for r in payload["rows"]:
        for pay in r["payments"]:
            p.append([r["period"].strftime("%m.%Y"), pay["date"].strftime("%d.%m.%Y"),
                      pay["paid"], pay["remaining_after"]])
    if p.max_row == 1:
        p.append(["—", "—", 0])
    _autosize(p)
    _money_format(p, 3)

    k = wb.create_sheet("Контроль")
    _header_row(k, ["Период", "Пени (итог), ₽", "Сумма отрезков, ₽", "Расхождение, ₽", "Статус"])
    for r in payload["rows"]:
        seg_sum = _round(sum(x["peni"] for x in r["segments"]))
        diff = _round(r["peni"] - seg_sum)
        k.append([r["period"].strftime("%m.%Y"), r["peni"], seg_sum, diff,
                  "сходится" if abs(diff) < 0.01 else "РАСХОЖДЕНИЕ"])
    _autosize(k)
    _money_format(k, 2)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=f"raschet_peni_{calc_id}.xlsx")


# ==================== ИНИЦИАЛИЗАЦИЯ ====================

def init_db():
    with app.app_context():
        db.create_all()
        if not Organization.query.first():
            org = Organization(
                name_full="Товарищество собственников жилья «Образец»",
                name_short="ТСЖ «Образец»",
                address="г. Москва, ул. Примерная, д. 1",
                inn="7700000000", ogrn="1157700000000")
            db.session.add(org)
            db.session.commit()
            admin = User(username="admin", full_name="Иванов Иван Иванович",
                         position="Главный бухгалтер", role="admin", org_id=org.id)
            admin.set_password("admin123")
            db.session.add(admin)
            db.session.commit()
            print("Создан пользователь: admin / admin123")


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=os.environ.get("FLASK_DEBUG", "0") == "1")
