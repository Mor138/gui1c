# logic/normalize_catalogs.py
# ──────────────────────────────────────────────────────────────
# Создание нормализованных справочников на основе грязной базы 1С
# Вызывается из CatalogsPage — наполняет словари для использования в UI


from collections import defaultdict
import config

_NORMALIZED = None

def load_normalized():
    """Возвращает нормализованные справочники, загружая их при первом вызове."""
    global _NORMALIZED
    if _NORMALIZED is not None:
        return _NORMALIZED

    normalized = {"Камни": [], "Вставки": [], "Изделия": []}

    nomenclature = config.BRIDGE.list_catalog_items("Номенклатура", 10000)
    for item in nomenclature:
        name = item.get("Наименование", "")
        insert = item.get("Вставка", "")
        method = item.get("ВариантИзготовления", "")

        if insert:
            normalized["Вставки"].append({"Название": name, "Вставка": insert})
        elif "камень" in name.lower():
            normalized["Камни"].append({"Название": name})
        else:
            normalized["Изделия"].append({"Название": name, "Метод": method})

    _NORMALIZED = normalized
    return _NORMALIZED


# Можно добавить сохранение в prod_cache.json, если нужно:
# import json, pathlib
# path = pathlib.Path("data/prod_cache.json")
# with open(path, "w", encoding="utf-8") as f:
#     json.dump(NORMALIZED, f, indent=2, ensure_ascii=False)
