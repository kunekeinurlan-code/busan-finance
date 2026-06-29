"""
BUSAN FINANCE — Складской модуль
Парсит stock report и строит аналитику по остаткам
"""
import pandas as pd
import json
import math
from collections import defaultdict

def parse_amount(val):
    if val is None: return 0.0
    if isinstance(val, float) and math.isnan(val): return 0.0
    try:
        return float(str(val).replace(" ","").replace("\xa0","").replace(",","."))
    except: return 0.0

def parse_stock_report(filepath):
    """Парсит выписку остатков из системы склада"""
    items = []
    try:
        engine = "xlrd" if filepath.endswith(".xls") else "openpyxl"
        raw = pd.read_excel(filepath, header=None, engine=engine)

        # Ищем строку заголовка
        header_row = None
        for i, row in raw.iterrows():
            vals = [str(v) for v in row.values]
            if any("Наименование" in v for v in vals) and any("Остаток" in v for v in vals):
                header_row = i
                break
        if header_row is None:
            return []

        # Определяем колонки
        header = [str(v).strip().replace("\n","") for v in raw.iloc[header_row].values]

        # Маппинг колонок
        col_map = {}
        for idx, h in enumerate(header):
            hl = h.lower()
            if "наименован" in hl: col_map["name"] = idx
            elif "артикул" in hl: col_map["sku"] = idx
            elif "код" in hl and "name" not in col_map: col_map["code"] = idx
            elif "остаток" in hl and "sum" not in hl: col_map["qty"] = idx
            elif "доступно" in hl: col_map["available"] = idx
            elif "резерв" in hl: col_map["reserve"] = idx
            elif "себестоимость" in hl and "сумм" not in hl.replace("себестоимости","X"): col_map["cost"] = idx
            elif "суммасебестоимости" in hl.replace(" ","") or ("сумм" in hl and "себест" in hl): col_map["cost_total"] = idx
            elif "цена" in hl and "продаж" in hl: col_map["price"] = idx
            elif "суммапродажи" in hl.replace(" ","") or ("сумм" in hl and "продаж" in hl): col_map["revenue_potential"] = idx
            elif "дней" in hl: col_map["days"] = idx

        current_category = "Без категории"
        for i in range(header_row + 1, len(raw)):
            row = raw.iloc[i].values
            non_null = [v for v in row if str(v) != "nan"]

            if not non_null:
                continue

            # Строка категории — одно значение текст без цифр
            if len(non_null) == 1 and isinstance(non_null[0], str):
                current_category = str(non_null[0]).strip()
                continue

            # Строка товара
            try:
                name_idx = col_map.get("name", 2)
                qty_idx  = col_map.get("qty", 7)
                name = str(row[name_idx]).strip() if name_idx < len(row) else ""
                if name in ("nan", "") or name.lower() == "итого":
                    continue

                qty   = parse_amount(row[qty_idx]) if qty_idx < len(row) else 0
                cost  = parse_amount(row[col_map["cost"]]) if "cost" in col_map else 0
                cost_total = parse_amount(row[col_map["cost_total"]]) if "cost_total" in col_map else cost * qty
                price = parse_amount(row[col_map["price"]]) if "price" in col_map else 0
                rev   = parse_amount(row[col_map["revenue_potential"]]) if "revenue_potential" in col_map else price * qty
                days  = parse_amount(row[col_map["days"]]) if "days" in col_map else 0
                sku   = str(row[col_map["sku"]]).strip() if "sku" in col_map else ""
                if sku == "nan": sku = ""

                if qty == 0 and cost_total == 0:
                    continue

                # Возраст товара
                if days <= 30:   age_group = "До 30 дней"
                elif days <= 90:  age_group = "30–90 дней"
                elif days <= 180: age_group = "90–180 дней"
                elif days <= 365: age_group = "180–365 дней"
                else:             age_group = "Более года (365+)"

                # Статус ликвидности
                if days <= 30:   status = "fresh"
                elif days <= 90:  status = "good"
                elif days <= 180: status = "watch"
                elif days <= 365: status = "risk"
                else:             status = "dead"

                items.append({
                    "sku": sku,
                    "name": name,
                    "category": current_category,
                    "qty": qty,
                    "cost": cost,
                    "cost_total": cost_total,
                    "price": price,
                    "revenue_potential": rev,
                    "days": days,
                    "age_group": age_group,
                    "status": status,
                    "margin": round((price - cost) / price * 100, 1) if price > 0 else 0,
                })
            except:
                continue
    except Exception as e:
        print(f"Stock parse error: {e}")
    return items


def build_stock_analytics(items):
    """Строит аналитику по складу"""
    if not items:
        return {}

    total_items   = len(items)
    total_qty     = sum(i["qty"] for i in items)
    total_cost    = sum(i["cost_total"] for i in items)
    total_revenue = sum(i["revenue_potential"] for i in items)
    avg_margin    = round((total_revenue - total_cost) / total_cost * 100, 1) if total_cost > 0 else 0

    # По возрасту
    age_groups = defaultdict(lambda: {"count":0,"qty":0,"cost":0,"revenue":0})
    for item in items:
        g = age_groups[item["age_group"]]
        g["count"]   += 1
        g["qty"]     += item["qty"]
        g["cost"]    += item["cost_total"]
        g["revenue"] += item["revenue_potential"]

    age_order = ["До 30 дней","30–90 дней","90–180 дней","180–365 дней","Более года (365+)"]
    age_table = []
    for ag in age_order:
        g = age_groups.get(ag, {"count":0,"qty":0,"cost":0,"revenue":0})
        age_table.append({
            "age_group": ag,
            "count":   g["count"],
            "qty":     g["qty"],
            "cost":    round(g["cost"], 2),
            "revenue": round(g["revenue"], 2),
            "cost_share":    round(g["cost"] / total_cost * 100, 1) if total_cost > 0 else 0,
            "count_share":   round(g["count"] / total_items * 100, 1) if total_items > 0 else 0,
        })

    # Замороженный капитал (>180 дней)
    frozen = sum(i["cost_total"] for i in items if i["days"] > 180)
    dead   = sum(i["cost_total"] for i in items if i["days"] > 365)

    # По категориям
    cat_data = defaultdict(lambda: {"count":0,"qty":0,"cost":0,"revenue":0,"days_sum":0})
    for item in items:
        c = cat_data[item["category"]]
        c["count"]    += 1
        c["qty"]      += item["qty"]
        c["cost"]     += item["cost_total"]
        c["revenue"]  += item["revenue_potential"]
        c["days_sum"] += item["days"] * item["qty"]

    categories = sorted([{
        "category": cat,
        "count":    v["count"],
        "qty":      v["qty"],
        "cost":     round(v["cost"], 2),
        "revenue":  round(v["revenue"], 2),
        "avg_days": round(v["days_sum"] / v["qty"], 1) if v["qty"] > 0 else 0,
        "cost_share": round(v["cost"] / total_cost * 100, 1) if total_cost > 0 else 0,
    } for cat, v in cat_data.items()], key=lambda x: -x["cost"])

    # Топ залежалых товаров
    dead_items = sorted(
        [i for i in items if i["days"] > 180],
        key=lambda x: -x["cost_total"]
    )[:20]

    # Топ по потенциалу продаж
    top_revenue = sorted(items, key=lambda x: -x["revenue_potential"])[:15]

    # Рекомендации
    recommendations = []
    dead_count = len([i for i in items if i["status"] == "dead"])
    dead_cost  = sum(i["cost_total"] for i in items if i["status"] == "dead")
    risk_count = len([i for i in items if i["status"] == "risk"])
    risk_cost  = sum(i["cost_total"] for i in items if i["status"] == "risk")

    if dead_cost > 0:
        recommendations.append({
            "type": "danger", "icon": "alert",
            "title": f"Мёртвый остаток: {dead_count} позиций",
            "text": f"Товар лежит более года — {dead_cost:,.0f} ₸ заморожено. Срочная распродажа со скидкой 40–60% освободит оборотные средства."
        })
    if risk_cost > 0:
        recommendations.append({
            "type": "warning", "icon": "alert",
            "title": f"Залежалый товар: {risk_count} позиций",
            "text": f"{risk_cost:,.0f} ₸ (180–365 дней). Провести акции, предложить скидки или договориться о возврате поставщику."
        })
    frozen_share = round(frozen / total_cost * 100) if total_cost > 0 else 0
    if frozen_share > 30:
        recommendations.append({
            "type": "warning", "icon": "info",
            "title": f"Заморожено {frozen_share}% капитала",
            "text": f"{frozen:,.0f} ₸ находится в товаре старше 180 дней. Это деньги которые не работают — пересмотрите закупочную политику."
        })
    fresh_share = round(age_groups.get("До 30 дней",{}).get("cost",0) / total_cost * 100) if total_cost > 0 else 0
    if fresh_share > 40:
        recommendations.append({
            "type": "success", "icon": "check",
            "title": "Хорошая оборачиваемость",
            "text": f"{fresh_share}% склада — свежий товар до 30 дней. Здоровая структура запасов."
        })

    return {
        "total_items":   total_items,
        "total_qty":     total_qty,
        "total_cost":    round(total_cost, 2),
        "total_revenue": round(total_revenue, 2),
        "avg_margin":    avg_margin,
        "frozen_capital": round(frozen, 2),
        "dead_capital":   round(dead, 2),
        "frozen_share":  round(frozen / total_cost * 100, 1) if total_cost > 0 else 0,
        "age_table":     age_table,
        "categories":    categories,
        "dead_items":    dead_items,
        "top_revenue":   top_revenue,
        "recommendations": recommendations,
    }
