# logic/normalize_catalogs.py
# ──────────────────────────────────────────────────────────────
# Создание нормализованных справочников на основе грязной базы 1С
# Вызывается из CatalogsPage — наполняет словари для использования в UI

from core.com_bridge import COM1CBridge
from collections import defaultdict

bridge = COM1CBridge("C:/Users/Mor/Desktop/1C/proiz")  # путь к базе можно заменить

# Чистый список номенклатуры с группировкой по типу
NORMALIZED = {
    "Камни": [],
    "Вставки": [],
    "Изделия": []
}

nomenclature = bridge.list_catalog_items("Номенклатура", 10000)

for item in nomenclature:
    name = item.get("Наименование", "")
    insert = item.get("Вставка", "")
    method = item.get("ВариантИзготовления", "")

    # Если присутствует Вставка — заносим как камень или вставку
    if insert:
        NORMALIZED["Вставки"].append({"Название": name, "Вставка": insert})
    elif "камень" in name.lower():
        NORMALIZED["Камни"].append({"Название": name})
    else:
        NORMALIZED["Изделия"].append({"Название": name, "Метод": method})


# Можно добавить сохранение в prod_cache.json, если нужно:
# import json, pathlib
# path = pathlib.Path("data/prod_cache.json")
# with open(path, "w", encoding="utf-8") as f:
#     json.dump(NORMALIZED, f, indent=2, ensure_ascii=False)
