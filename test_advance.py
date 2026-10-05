#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Проверка авансовых платежей и переплаты."""

from datetime import date
from io import BytesIO
import openpyxl

from app import app, init_db
from peni_engine import calculate_fifo, due_date_for_period

CALC = date(2026, 10, 4)


def one(amount, due, pays, kap=False, mode="fixed95"):
    r = calculate_fifo([{"period": due.replace(day=1), "amount": amount, "deadline": due}],
                       [{"date": d, "amount": a} for d, a in pays], CALC,
                       kapremont=kap, rate_mode=mode)
    return r


print("=== 1. Аванс закрывает период до срока оплаты ===")
# долг 2500, срок 10.02.2026, платёж 01.02.2026 (до срока) на всю сумму
r = one(2500, date(2026, 2, 10), [(date(2026, 2, 1), 2500)])
row = r["rows"][0]
print(f"  остаток {row['remaining']}, пени {row['peni']} (ожидается 0 — оплачено до срока)")
print(f"  авансом: {row['advance_paid']}, примечание: {row['note']}")
assert row["peni"] == 0.0, "за долг, оплаченный до срока, пеня не начисляется"

print()
print("=== 2. Частичный аванс: пеня только на неоплаченную часть ===")
# долг 4000, срок 10.02.2026, аванс 1000 01.02.2026
r = one(4000, date(2026, 2, 10), [(date(2026, 2, 1), 1000)])
row = r["rows"][0]
print(f"  остаток {row['remaining']}, пени {row['peni']}, авансом {row['advance_paid']}")
print(f"  примечание: {row['note']}")
assert "аванс" in row["note"]
# сравним с полной суммой
r_full = one(4000, date(2026, 2, 10), [])
print(f"  для сравнения без аванса: {r_full['rows'][0]['peni']}")
assert row["peni"] < r_full["rows"][0]["peni"]

print()
print("=== 3. Переплата: платёж больше начисленного ===")
r = one(600, date(2026, 2, 10), [(date(2026, 3, 1), 2000)])
print(f"  остаток {r['rows'][0]['remaining']}, переплата {r['overpayment']}, авансом {r['advance_total']}")
assert r["total_debt"] == 0.0
assert r["overpayment"] == 1400.0, r["overpayment"]

print()
print("=== 4. Сверка: warning по умолчанию, error в strict ===")
with app.test_client() as c:
    c.post('/login', data={'username': 'admin', 'password': 'admin123'})

    def book(charges, payments):
        wb = openpyxl.Workbook()
        ws = wb.active; ws.title = "Начисления"
        ws.append(["Лицевой счёт", "Период", "Начисление", "Срок оплаты", "Вид долга"])
        for x in charges: ws.append(list(x))
        ws2 = wb.create_sheet("Платежи")
        ws2.append(["Лицевой счёт", "Дата платежа", "Сумма платежа"])
        for x in payments: ws2.append(list(x))
        b = BytesIO(); wb.save(b); return b.getvalue()

    # оплачено больше, чем начислено
    raw = book([("ЛС-1", "01.02.2026", 600, "10.03.2026", "ЖКУ")],
               [("ЛС-1", "20.03.2026", 2000)])

    r1 = c.post('/batch_calculate', data={'file': (BytesIO(raw), 'r.xlsx'),
               'calc_date': '2026-10-04', 'rate_mode': 'fixed95'},
               content_type='multipart/form-data')
    print(f"  режим «разрешить»: {r1.status_code}")
    out = openpyxl.load_workbook(BytesIO(r1.data), data_only=True)
    for row in out['Сверка'].iter_rows(min_row=5, values_only=True):
        if row and row[0]:
            print("   Сверка:", row[0], "| начислено", row[1], "| оплачено", row[2],
                  "| в т.ч. аванс", row[4] if len(row) > 4 else "?", "|", row[5] if len(row) > 5 else "")
    for row in out['Сводка'].iter_rows(min_row=9, values_only=True):
        if row and row[0] and str(row[0]).startswith("ЛС"):
            print("   Сводка:", row[0], "| начислено", row[1], "| оплачено", row[2],
                  "| аванс", row[3], "| долг", row[4], "| переплата", row[3])

    r2 = c.post('/batch_calculate', data={'file': (BytesIO(raw), 'r.xlsx'),
               'calc_date': '2026-10-04', 'rate_mode': 'fixed95', 'strict': '1'},
               content_type='multipart/form-data')
    print(f"  режим «запретить»: {r2.status_code}")
    print("   ответ:", r2.data.decode()[:160].replace("\n", " "))

print()
print("ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ")
