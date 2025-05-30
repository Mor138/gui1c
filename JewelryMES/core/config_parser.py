# config_parser.py
# ─────────────────────────────────────────────────────────────
import xml.etree.ElementTree as ET
from pathlib import Path

CONFIG_XML = Path("data/Configuration.xml")
DUMP_XML = Path("data/ConfigDumpInfo.xml")

# ─────────────────────── Реальная загрузка значений справочников ───────────────────────
def get_catalog_items(name: str) -> list[str]:
    """Чтение значений конкретного справочника из ConfigDumpInfo.xml"""
    if not DUMP_XML.exists():
        return []

    ns = {'v8': 'http://v8.1c.ru/8.1/data/core'}
    tree = ET.parse(DUMP_XML)
    root = tree.getroot()

    result = set()

    for obj in root.findall(".//v8:CatalogObject", ns):
        if obj.attrib.get("name") == name:
            for el in obj.findall("v8:Items/v8:Item", ns):
                val = el.findtext("v8:Description", default="", namespaces=ns)
                if val:
                    result.add(val)
    return sorted(result)

# ─────────────────────── Найти список всех справочников ───────────────────────
def extract_catalog_names() -> dict:
    """Извлекает список всех справочников с русским именем (синонимом)."""
    if not CONFIG_XML.exists():
        return {}

    ns = {'v8': 'http://v8.1c.ru/8.1/data/core'}
    tree = ET.parse(CONFIG_XML)
    root = tree.getroot()

    catalogs = {}
    for elem in root.findall(".//v8:Catalog", ns):
        name = elem.attrib.get("name") or elem.findtext("v8:Name", default="", namespaces=ns)
        synonym = elem.findtext("v8:Synonym/v8:item/v8:content", default="", namespaces=ns)
        if name and synonym:
            catalogs[name] = synonym
    return catalogs

# ─────────────────────── Пример CLI-отладки ───────────────────────
if __name__ == "__main__":
    all_catalogs = extract_catalog_names()
    for k, v in all_catalogs.items():
        print(f"{k} → {v}")
    print("\nТестовое чтение элементов 'Контрагенты':")
    print(get_catalog_items("Контрагенты"))
