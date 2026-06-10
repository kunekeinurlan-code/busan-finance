
from flask import Flask, request, jsonify, send_file, send_from_directory
import os, json, io
from datetime import datetime
from collections import defaultdict
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

NAVY  = "1A1A2E"
GOLD  = "D4A017"
WHITE = "FFFFFF"
LIGHT = "F5F5F0"
GREEN = "27AE60"
RED   = "C0392B"
GRAY  = "CCCCCC"

def fill(c): return PatternFill("solid", fgColor=c)
def font(bold=False, color=WHITE, size=11, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)
def border_thin():
    s = Side(style="thin", color=GRAY)
    return Border(left=s, right=s, top=s, bottom=s)

CATEGORIES = {
    "Зарплата / доход":     ["зарплат","salary","оклад","выплат","начислен"],
    "Доход от клиентов":    ["клиент","оплат","консалт","услуг","гонорар","вознагражд"],
    "Переводы входящие":    ["перевод","transfer","пополнен"],
    "Налоги / взносы":      ["налог","ипн","соц","пенсион","опв","снп","осмс"],
    "Аренда":               ["аренд","rent","найм"],
    "Продукты / питание":   ["магазин","супермарк","продукт","еда","food","market","magnum","small"],
    "Кафе / рестораны":     ["кафе","ресторан","cafe","restaurant","coffee","кофе","bar"],
    "Транспорт":            ["такси","uber","yandex","яндекс","bolt","автобус","транспорт","parking"],
    "Связь / интернет":     ["beeline","kcell","activ","altel","tele2","билайн","интернет","связь"],
    "Подписки / сервисы":   ["netflix","spotify","apple","google","youtube","подписк"],
    "Образование":          ["курс","обучен","образован","школ","универс","тренинг","семинар"],
    "Здоровье / медицина":  ["аптек","клиник","больниц","медицин","pharmacy","dental","стомат"],
    "Красота / уход":       ["салон","beauty","spa","cosmetic","косметик"],
    "Банковские расходы":   ["комисси","обслуживан","штраф","пени"],
    "IT / сервисы":         ["it","разработк","программ","software","хостинг"],
    "Инвестиции":           ["инвестиц","депозит","брокер","акци","облигац"],
}

def categorize(desc):
    d = str(desc).lower()
    for cat, kws in CATEGORIES.items():
        if any(k in d for k in kws):
            return cat
    return "Прочие"

def parse_amount(val):
    if val is None: return 0.0
    try:
        s = str(val).replace(" ","").replace("\u202f","").replace(",",".")
        return float(s)
    except: return 0.0

def detect_bank(filename, filepath):
    fn = filename.lower()
    if "kaspi" in fn: return "Kaspi"
    if "halyk" in fn or "народный" in fn: return "Halyk"
    if "freedom" in fn or "ffin" in fn: return "Freedom"
    try:
        df = pd.read_excel(filepath, header=None, nrows=10)
        text = " ".join(df.fillna("").astype(str).values.flatten()).lower()
        if "kaspi" in text: return "Kaspi"
        if "halyk" in text or "народный" in text: return "Halyk"
        if "freedom" in text or "ffin" in text: return "Freedom"
    except: pass
    return "Другой"

def parse_file(filepath, bank):
    transactions = []
    try:
        raw = pd.read_excel(filepath, header=None, nrows=20)
        header_row = None
        for i, row in raw.iterrows():
            vals = [str(v).lower() for v in row.values]
            score = sum(1 for v in vals if any(k in v for k in ["дата","date","сумм","дебет","кредит","списан","зачислен"]))
            if score >= 2:
                header_row = i
                break
        if header_row is None: header_row = 0

        df = pd.read_excel(filepath, skiprows=header_row, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        date_col = next((c for c in df.columns if any(k in c.lower() for k in ["дата","date"])), None)
        if not date_col: return []

        for _, row in df.iterrows():
            try:
                date = pd.to_datetime(str(row.get(date_col,"")), dayfirst=True, errors="coerce")
                if pd.isna(date): continue
                debit = credit = 0.0
                for col in df.columns:
                    v = parse_amount(row.get(col, 0))
                    cl = col.lower()
                    if any(k in cl for k in ["списан","расход","дебет","debit","дт"]):
                        debit = abs(v)
                    elif any(k in cl for k in ["зачислен","приход","кредит","credit","кт"]):
                        credit = abs(v)
                if debit == 0 and credit == 0: continue
                desc = ""
                for col in df.columns:
                    if any(k in col.lower() for k in ["назначен","описан","получатель","контрагент","purpose"]):
                        v = str(row.get(col,"")).strip()
                        if v and v.lower() != "nan":
                            desc = v; break
                transactions.append({
                    "date": date.strftime("%Y-%m-%d"),
                    "month": date.strftime("%Y-%m"),
                    "bank": bank,
                    "description": desc,
                    "debit": debit,
                    "credit": credit,
                    "category": categorize(desc),
                    "type": "expense" if debit > 0 else "income",
                    "amount": debit if debit > 0 else credit,
                })
            except: continue
    except Exception as e:
        print(f"Parse error: {e}")
    return transactions

def generate_excel(transactions):
    wb = openpyxl.Workbook()

    # Sheet 1: Dashboard
    ws = wb.active; ws.title = "Дашборд"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:G1")
    c = ws["A1"]; c.value = "BUSAN FINANCE  |  CASHFLOW ОТЧЁТ"
    c.fill = fill(NAVY); c.font = Font(name="Arial", bold=True, color=GOLD, size=14)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    ws.merge_cells("A2:G2")
    c = ws["A2"]; c.value = f"Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    c.fill = fill(NAVY); c.font = font(color=GOLD, size=9, italic=True)
    c.alignment = Alignment(horizontal="center")

    total_inc = sum(t["credit"] for t in transactions)
    total_exp = sum(t["debit"]  for t in transactions)
    net = total_inc - total_exp
    banks = list(set(t["bank"] for t in transactions))

    kpis = [("Доходы (₸)", f"{total_inc:,.0f}"), ("Расходы (₸)", f"{total_exp:,.0f}"),
            ("Чистый CF (₸)", f"{net:+,.0f}"), ("Транзакций", str(len(transactions))),
            ("Банков", str(len(banks)))]
    for ci, (label, val) in enumerate(kpis, 1):
        c = ws.cell(row=4, column=ci, value=label)
        c.fill = fill(GOLD); c.font = font(bold=True, color=NAVY, size=9); c.alignment = Alignment(horizontal="center")
        c = ws.cell(row=5, column=ci, value=val)
        c.fill = fill(LIGHT); c.font = Font(name="Arial", bold=True, color=NAVY, size=12)
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(ci)].width = 20

    # Sheet 2: Transactions
    ws2 = wb.create_sheet("Транзакции")
    ws2.sheet_view.showGridLines = False
    hdrs = ["Дата","Банк","Тип","Категория","Описание","Расход","Доход"]
    for ci, h in enumerate(hdrs, 1):
        c = ws2.cell(row=1, column=ci, value=h)
        c.fill = fill(NAVY); c.font = font(bold=True, color=GOLD, size=10)
        c.alignment = Alignment(horizontal="center"); c.border = border_thin()
    ws2.column_dimensions["A"].width = 12; ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 10; ws2.column_dimensions["D"].width = 22
    ws2.column_dimensions["E"].width = 45; ws2.column_dimensions["F"].width = 16
    ws2.column_dimensions["G"].width = 16
    for ri, t in enumerate(sorted(transactions, key=lambda x: x["date"], reverse=True), 2):
        bg = LIGHT if ri % 2 == 0 else WHITE
        row_vals = [t["date"], t["bank"],
                    "Расход" if t["type"]=="expense" else "Доход",
                    t["category"], t["description"][:80],
                    t["debit"] if t["debit"] > 0 else "",
                    t["credit"] if t["credit"] > 0 else ""]
        for ci, v in enumerate(row_vals, 1):
            c = ws2.cell(row=ri, column=ci, value=v)
            c.fill = fill(bg); c.border = border_thin()
            if ci == 3:
                c.font = font(color=RED if v=="Расход" else GREEN, size=9, bold=True)
            elif ci in [6,7]:
                c.font = font(color=RED if ci==6 else GREEN, size=9)
                c.number_format = "#,##0"
            else:
                c.font = font(color=NAVY, size=9)
    ws2.auto_filter.ref = f"A1:G{len(transactions)+1}"
    ws2.freeze_panes = "A2"

    # Sheet 3: By category
    ws3 = wb.create_sheet("По категориям")
    ws3.sheet_view.showGridLines = False
    cat_inc = defaultdict(float); cat_exp = defaultdict(float)
    for t in transactions:
        if t["type"] == "income": cat_inc[t["category"]] += t["credit"]
        else: cat_exp[t["category"]] += t["debit"]
    for ci, h in enumerate(["Категория","Доходы (₸)","Расходы (₸)","Разница (₸)"], 1):
        c = ws3.cell(row=1, column=ci, value=h)
        c.fill = fill(NAVY); c.font = font(bold=True, color=GOLD, size=10)
        c.alignment = Alignment(horizontal="center"); c.border = border_thin()
    ws3.column_dimensions["A"].width = 28
    for col in ["B","C","D"]: ws3.column_dimensions[col].width = 18
    all_cats = sorted(set(list(cat_inc.keys()) + list(cat_exp.keys())))
    for ri, cat in enumerate(all_cats, 2):
        bg = LIGHT if ri % 2 == 0 else WHITE
        inc = cat_inc.get(cat, 0); exp = cat_exp.get(cat, 0); diff = inc - exp
        for ci, v in enumerate([cat, inc, exp, diff], 1):
            c = ws3.cell(row=ri, column=ci, value=v if v != 0 else "")
            c.fill = fill(bg); c.border = border_thin()
            if ci == 1: c.font = font(color=NAVY, size=9, bold=True)
            elif ci == 2: c.font = font(color=GREEN, size=9); c.number_format = "#,##0"
            elif ci == 3: c.font = font(color=RED, size=9); c.number_format = "#,##0"
            elif ci == 4:
                c.font = font(color=GREEN if isinstance(diff, float) and diff >= 0 else RED, size=9, bold=True)
                c.number_format = "#,##0"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route("/")
def index():
    return send_file("index.html")

@app.route("/analyze", methods=["POST"])
def analyze():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "Файлы не загружены"}), 400

    all_transactions = []
    file_results = []
    upload_dir = "/tmp/busan_uploads"
    os.makedirs(upload_dir, exist_ok=True)

    for f in files:
        path = os.path.join(upload_dir, f.filename)
        f.save(path)
        bank = detect_bank(f.filename, path)
        txns = parse_file(path, bank)
        all_transactions.extend(txns)
        file_results.append({"name": f.filename, "bank": bank, "count": len(txns)})
        os.remove(path)

    if not all_transactions:
        return jsonify({"error": "Транзакции не найдены. Проверьте формат файлов."}), 400

    total_inc  = sum(t["credit"] for t in all_transactions)
    total_exp  = sum(t["debit"]  for t in all_transactions)
    net        = total_inc - total_exp
    banks      = list(set(t["bank"] for t in all_transactions))

    # По категориям
    cat_exp = defaultdict(float)
    cat_inc = defaultdict(float)
    for t in all_transactions:
        if t["type"] == "expense": cat_exp[t["category"]] += t["debit"]
        else: cat_inc[t["category"]] += t["credit"]

    top_exp = sorted(cat_exp.items(), key=lambda x: -x[1])[:8]
    top_inc = sorted(cat_inc.items(), key=lambda x: -x[1])[:5]

    # По месяцам
    monthly = defaultdict(lambda: {"income": 0.0, "expense": 0.0})
    for t in all_transactions:
        if t["type"] == "income": monthly[t["month"]]["income"] += t["credit"]
        else: monthly[t["month"]]["expense"] += t["debit"]
    monthly_sorted = [{"month": m, **v} for m, v in sorted(monthly.items())]

    # По банкам
    bank_data = defaultdict(lambda: {"income": 0.0, "expense": 0.0, "count": 0})
    for t in all_transactions:
        bank_data[t["bank"]]["income"]  += t["credit"]
        bank_data[t["bank"]]["expense"] += t["debit"]
        bank_data[t["bank"]]["count"]   += 1
    banks_list = [{"bank": b, **v} for b, v in bank_data.items()]

    # Рекомендации
    recommendations = generate_recommendations(all_transactions, total_inc, total_exp, net, cat_exp, monthly_sorted)

    # Сохраняем транзакции для Excel
    import json as _json
    with open("/tmp/busan_txns.json", "w", encoding="utf-8") as f:
        _json.dump(all_transactions, f, ensure_ascii=False)

    return jsonify({
        "files": file_results,
        "total": len(all_transactions),
        "total_income": total_inc,
        "total_expense": total_exp,
        "net_cashflow": net,
        "banks": banks_list,
        "top_expenses": [{"category": k, "amount": v} for k,v in top_exp],
        "top_income": [{"category": k, "amount": v} for k,v in top_inc],
        "monthly": monthly_sorted,
        "recommendations": recommendations,
    })


def generate_recommendations(txns, inc, exp, net, cat_exp, monthly):
    recs = []
    savings_rate = (net / inc * 100) if inc > 0 else 0

    if savings_rate < 10:
        recs.append({
            "type": "danger",
            "icon": "alert",
            "title": "Низкая норма сбережений",
            "text": f"Текущий уровень: {savings_rate:.1f}%. Рекомендуемый минимум — 20%. Проанализируйте топ расходных категорий для оптимизации."
        })
    elif savings_rate >= 30:
        recs.append({
            "type": "success",
            "icon": "check",
            "title": "Хорошая норма сбережений",
            "text": f"Уровень сбережений {savings_rate:.1f}% — выше нормы. Рассмотрите инвестиционные инструменты для роста капитала."
        })

    if net < 0:
        recs.append({
            "type": "danger",
            "icon": "alert",
            "title": "Отрицательный денежный поток",
            "text": f"Расходы превышают доходы на {abs(net):,.0f} ₸. Необходим срочный аудит бюджета."
        })

    top_cat = max(cat_exp.items(), key=lambda x: x[1]) if cat_exp else None
    if top_cat:
        share = top_cat[1] / exp * 100 if exp > 0 else 0
        if share > 30:
            recs.append({
                "type": "warning",
                "icon": "info",
                "title": f"Концентрация расходов: «{top_cat[0]}»",
                "text": f"{share:.1f}% всех расходов приходится на одну категорию ({top_cat[1]:,.0f} ₸). Диверсифицируйте бюджет."
            })

    if len(monthly) >= 3:
        recent = monthly[-3:]
        cf_trend = [m["income"] - m["expense"] for m in recent]
        if all(cf_trend[i] < cf_trend[i-1] for i in range(1, len(cf_trend))):
            recs.append({
                "type": "warning",
                "icon": "trend",
                "title": "Нисходящий тренд денежного потока",
                "text": "Чистый CF снижается последние 3 месяца. Проверьте источники доходов и динамику расходов."
            })
        elif all(cf_trend[i] > cf_trend[i-1] for i in range(1, len(cf_trend))):
            recs.append({
                "type": "success",
                "icon": "trend",
                "title": "Положительный тренд",
                "text": "Денежный поток растёт 3 месяца подряд. Хороший момент для формирования резервного фонда."
            })

    tax_amount = cat_exp.get("Налоги / взносы", 0)
    if tax_amount > 0:
        tax_share = tax_amount / exp * 100 if exp > 0 else 0
        recs.append({
            "type": "info",
            "icon": "info",
            "title": "Налоговая нагрузка",
            "text": f"Уплачено налогов и взносов: {tax_amount:,.0f} ₸ ({tax_share:.1f}% расходов). Проверьте применение налоговых вычетов."
        })

    if not recs:
        recs.append({
            "type": "success",
            "icon": "check",
            "title": "Финансовые показатели в норме",
            "text": "Структура доходов и расходов сбалансирована. Продолжайте вести регулярный учёт."
        })
    return recs


@app.route("/download-excel", methods=["GET"])
def download_excel():
    try:
        import json as _json
        with open("/tmp/busan_txns.json", "r", encoding="utf-8") as f:
            txns = _json.load(f)
        buf = generate_excel(txns)
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        return send_file(buf, as_attachment=True,
                         download_name=f"BUSAN_Cashflow_{ts}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
