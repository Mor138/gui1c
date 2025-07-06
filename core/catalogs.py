# catalogs.py
# -*- coding: utf-8 -*-
"""Единое место для всех констант-справочников Jewelry MES."""

# ─────────────  Юридические лица / контрагенты  ─────────────
ORGANIZATIONS = [
    "ООО «Рост Золото»",
    "ИП Смирнова",
    "ЗАО «Злато»",
]

CLIENTS = [
    "ООО «Алмаз»",
    "Zlato Ltd.",
    "ИП Королёв",
]

CONTRACTS = [
    "Купля-продажа",
    "Комиссия",
    "Давальческий",
]

WAREHOUSES = [
    "Основной склад",
    "Салон",
    "Интернет-склад",
]

BANKS = [
    "АО «Банк 1»",
    "Касса № 1",
    "Р/с ИП Смирнова",
]

MANAGERS = [
    "Петров А.А.",
    "Сидорова И.И.",
]

PRICE_TYPES = ["Розница", "Опт", "Спец."]
CURRENCIES  = ["RUB", "USD"]

# ─────────────  Материалы, пробы, цвета  ─────────────
_METALS = {
    "Золото": {
        "hallmarks": ["375", "585", "750"],
        "colors":    ["красный", "желтый", "белый", "зелёный", "краснобелый", "желтообелый", "чернение"],
    },
    "Серебро": {
        "hallmarks": ["925"],
        "colors":    ["светлый", "тёмный", "чернение"],
    },
}

def metals() -> list[str]:
    return list(_METALS.keys())

def hallmarks(metal: str) -> list[str]:
    return _METALS.get(metal, {}).get("hallmarks", [])

def colors(metal: str) -> list[str]:
    return _METALS.get(metal, {}).get("colors", [])

def metal_by_hallmark(hallmark: str) -> str | None:
    """Возвращает металл по коду пробы."""
    for m, data in _METALS.items():
        if hallmark in data.get("hallmarks", []):
            return m
    return None

INSERTS = ["Нет", "Фианит", "Бриллиант", "Сапфир", "Изумруд"]

# ─────────────  Номенклатура изделий  ─────────────
#  артикул → dict(name, avg_weight_g, base_price)
NOMENCLATURE = {
    "R-1001": {"name": "Кольцо классическое","w": 3.20, "price": 8500,"method":"rubber"},
    "R-2001": {"name": "Серьги","w": 9.80, "price": 21500,"method":"rubber"},
    "R-3001": {"name": "Пирсинг","w": 14.5, "price": 33500,"method":"rubber"},
    "R-4001": {"name": "Кулон сердце","w": 1.95, "price": 5200,"method":"rubber"},
    "3D-1003": {"name":"Кольцо классическое","w":3.2,"price":8500,"method":"3d"},
    "3D-4003": {"name":"Кулон сердце","w":1.95,"price":5200,"method":"3d"},
}

# ─────────────  Ставки НДС  ─────────────
VAT_RATES = {"Без НДС": 0, "10 %": 10, "20 %": 20}
