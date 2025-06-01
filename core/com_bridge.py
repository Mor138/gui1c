import win32com.client
import pywintypes

def safe_str(val):
    try:
        return str(getattr(val, "Description", val)) if val is not None else ""
    except:
        return str(val)

def log(msg):
    print("[LOG]", msg)

class COM1CBridge:
    def __init__(self, base_path, usr="Администратор", pwd=""):
        self.connector = win32com.client.Dispatch("V83.COMConnector")
        self.connection = self.connector.Connect(f'File="{base_path}";Usr="{usr}";Pwd="{pwd}"')
        self.catalogs = self.connection.Catalogs
        self.documents = self.connection.Documents
        self.cache_variants()  # Предзагрузка всех вариантов

    def get_articles(self):
        result = {}
        catalog = getattr(self.catalogs, "Номенклатура", None)
        if not catalog:
            return result
        selection = catalog.Select()
        while selection.Next():
            obj = selection.GetObject()
            art = safe_str(getattr(obj, "Артикул", ""))
            result[art] = {
                "name": safe_str(obj.Description),
                "variant": safe_str(getattr(obj, "ВариантИзготовления", "")),
                "size": getattr(obj, "Размер1", ""),
                "w": getattr(obj, "СреднийВес", 0),
            }
        return result

    def get_production_status_variants(self) -> list[str]:
        try:
            enum = getattr(self.connection.Enums, "ВидСтатусПродукции", None)
            if enum is None:
                print("[COM ERROR] Перечисление 'ВидСтатусПродукции' не найдено.")
                return []
            return [str(val.Представление()) for val in enum.GetValues()]
        except Exception as e:
            print(f"[COM ERROR] ВидСтатусПродукции: {e}")
            return []

    def cache_variants(self):
        self._all_variants = []
        catalog = getattr(self.catalogs, "ВариантыИзготовленияНоменклатуры", None)
        if catalog is None:
            return
        selection = catalog.Select()
        while selection.Next():
            item = selection.GetObject()
            self._all_variants.append(str(item.Description))

    def get_variants_by_article(self, article_prefix: str) -> list[str]:
        if not hasattr(self, "_all_variants"):
            self.cache_variants()
        prefix = article_prefix.strip() + "-"
        return [name for name in self._all_variants if name.startswith(prefix)]

    def get_catalog_object_by_description(self, catalog_name, description):
        catalog = getattr(self.catalogs, catalog_name, None)
        if not catalog:
            log(f"Каталог '{catalog_name}' не найден")
            return None
        selection = catalog.Select()
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Description).strip() == str(description).strip():
                log(f"[{catalog_name}] Найден: {description}")
                return obj
        log(f"[{catalog_name}] Не найден: {description}")
        return None

    def get_ref(self, catalog_name, description):
        obj = self.get_catalog_object_by_description(catalog_name, description)
        return obj.Ref if obj else None

    def get_last_order_number(self):
        doc = getattr(self.documents, "ЗаказВПроизводство", None)
        if not doc:
            return "00ЮП-000000"
        selection = doc.Select()
        number = "00ЮП-000000"
        while selection.Next():
            try:
                obj = selection.GetObject()
                number = str(obj.Number)
            except:
                continue
        return number

    def get_next_order_number(self):
        last = self.get_last_order_number()
        try:
            prefix, num = last.split("-")
            next_num = int(num) + 1
            return f"{prefix}-{next_num:06d}"
        except:
            return "00ЮП-000001"

    def create_order(self, fields, items):
        doc = self.documents.ЗаказВПроизводство.CreateDocument()
        catalog_fields_map = {
            "Организация": "Организации",
            "Контрагент": "Контрагенты",
            "ДоговорКонтрагента": "ДоговорыКонтрагентов",
            "Ответственный": "Пользователи",
            "Склад": "Склады",
            "ВидСтатусПродукции": "ВидСтатусПродукции"
        }

        log("Создание заказа. Поля:")
        for k, v in fields.items():
            try:
                log(f"  -> {k} = {v}")
                if k == "Склад":
                    log("    ⚠ Поле 'Склад' не устанавливается вручную. Пропускаем.")
                    continue
                if k in catalog_fields_map:
                    ref = self.get_ref(catalog_fields_map[k], v)
                    if ref:
                        setattr(doc, k, ref)
                        log(f"    Установлено: {k} (Ref: {ref})")
                    else:
                        log(f"    ❌ {k} '{v}' не найден.")
                    continue
                setattr(doc, k, v)
            except Exception as e:
                log(f"    ❌ Ошибка установки поля {k}: {e}")

        log("Добавление строк заказа:")
        for row in items:
            try:
                log(f"  -> строка: {row}")
                new_row = doc.Товары.Add()

                new_row.Номенклатура = self.get_ref("Номенклатура", row.get("Номенклатура"))

                variant = row.get("ВариантИзготовления")
                if variant and variant != "—":
                    new_row.ВариантИзготовления = self.get_ref("ВариантыИзготовленияНоменклатуры", variant)

                new_row.Размер = float(row.get("Размер", 0))
                new_row.Количество = int(row.get("Количество", 1))
                new_row.Вес = float(row.get("Вес", 0))

            except Exception as e:
                log(f"    ❌ Ошибка в строке заказа: {e}")

        try:
            log("Проводим документ...")
            doc.Write()
            log(f"✅ Документ проведён. Номер: {doc.Number}")
            return str(doc.Number)
        except Exception as e:
            log(f"❌ Ошибка при записи документа: {e}")
            return f"Ошибка: {e}"

    def list_orders(self, limit=1000):
        result = []
        doc = getattr(self.documents, "ЗаказВПроизводство", None)
        if doc is None:
            return result
        selection = doc.Select()
        while selection.Next():
            obj = selection.GetObject()
            rows = []
            for line in obj.Товары:
                rows.append({
                    "nomenclature": safe_str(getattr(line, "Номенклатура", "")),
                    "variant": safe_str(getattr(line, "ВариантИзготовления", "")),
                    "status": safe_str(getattr(line, "ВидСтатусПродукции", "")),
                    "size": getattr(line, "Размер", 0),
                    "qty": getattr(line, "Количество", 0),
                    "w": getattr(line, "Вес", 0),
                })
            result.append({
                "num": str(obj.Number),
                "date": str(obj.Date),
                "contragent": safe_str(getattr(obj, "Контрагент", "")),
                "rows": rows
            })
        return result

    def list_catalog_items(self, catalog_name, limit=1000):
        result = []
        catalog = getattr(self.catalogs, catalog_name, None)
        if not catalog:
            log(f"[Catalog Error] {catalog_name}: not found")
            return result
        selection = catalog.Select()
        count = 0
        while selection.Next() and count < limit:
            obj = selection.GetObject()
            result.append({
                "Description": safe_str(obj.Description),
                "Ref": obj.Ref
            })
            count += 1
        return result
