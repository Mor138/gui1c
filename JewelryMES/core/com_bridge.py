import win32com.client


def safe_str(val):
    try:
        return str(getattr(val, "Description", val)) if val is not None else ""
    except:
        return str(val)


class COM1CBridge:
    def __init__(self, base_path, usr="Администратор", pwd=""):
        self.connector = win32com.client.Dispatch("V83.COMConnector")
        self.connection = self.connector.Connect(f'File="{base_path}";Usr="{usr}";Pwd="{pwd}"')
        self.catalogs = self.connection.Catalogs
        self.documents = self.connection.Documents

    def get_articles(self) -> dict:
        result = {}
        catalog = getattr(self.catalogs, "Номенклатура", None)
        if not catalog:
            print("[Catalog Error] Номенклатура: not found")
            return result
        try:
            selection = catalog.Select()
        except Exception:
            print("[get_articles error] 'NoneType' object is not callable")
            return result

        while True:
            try:
                obj = selection.Next()
                if not obj:
                    break
                obj = selection.GetObject()
                art = safe_str(getattr(obj, "Артикул", ""))
                result[art] = {
                    "name": safe_str(obj.Description),
                    "metal": safe_str(getattr(obj, "Металл", "")),
                    "hallmark": safe_str(getattr(obj, "Проба", "")),
                    "color": safe_str(getattr(obj, "ЦветМеталла", "")),
                    "insert": safe_str(getattr(obj, "Камень", "")),
                    "size": getattr(obj, "Размер1", ""),
                    "w": getattr(obj, "СреднийВес", 0),
                }
            except Exception:
                continue
        return result

    def get_manufacturing_options(self) -> dict:
        result = {}
        catalog = getattr(self.catalogs, "ВариантыИзготовленияНоменклатуры", None)
        if not catalog:
            print("[Catalog Error] ВариантыИзготовленияНоменклатуры: not found")
            return result
        try:
            selection = catalog.Select()
        except Exception:
            print("[get_manufacturing_options error] 'NoneType' object is not callable")
            return result

        while True:
            try:
                obj = selection.Next()
                if not obj:
                    break
                obj = selection.GetObject()
                desc = safe_str(obj)
                metal, hallmark, color = "", "", ""

                parts = desc.split()
                for part in parts:
                    if part.isdigit(): hallmark = part
                    elif "бел" in part.lower() or "крас" in part.lower() or "желт" in part.lower():
                        color = part.capitalize()
                    else:
                        metal = part.capitalize()

                result[desc] = {
                    "metal": metal,
                    "hallmark": hallmark,
                    "color": color,
                }
            except Exception:
                continue
        return result

    def list_catalog_items(self, catalog_name: str, limit: int = 100000) -> list:
        result = []
        catalog = getattr(self.catalogs, catalog_name, None)
        if catalog is None:
            print(f"[Catalog Error] {catalog_name}: not found")
            return result
        try:
            selection = catalog.Select()
        except Exception:
            print(f"[Catalog Error] {catalog_name}: 'NoneType' object is not callable")
            return result
        while True:
            try:
                obj = selection.Next()
                if not obj:
                    break
                obj = selection.GetObject()
                result.append({"Description": safe_str(obj.Description)})
            except Exception:
                continue
        return result
        
    def get_inserts(self) -> list:
        result = []
        catalog = getattr(self.catalogs, "складкамни", None)
        if catalog is None:
            print("[Catalog Error] складкамни: not found")
            return result
        try:
            selection = catalog.Select()
            while True:
                obj = selection.Next()
                if not obj:
                    break
                obj = selection.GetObject()
                result.append(str(getattr(obj, "Description", obj)))
        except Exception as e:
            print("[get_inserts error]", e)
        return sorted(set(result))    

    def list_orders(self, limit=5000):
        result = []
        doc = getattr(self.documents, "ЗаказВПроизводство", None)
        if doc is None:
            print("[1C Error: list_orders] Document not found")
            return result
        try:
            selection = doc.Select()
        except Exception:
            print("[list_orders error] 'NoneType' object is not callable")
            return result

        while True:
            try:
                obj = selection.Next()
                if not obj:
                    break
                obj = selection.GetObject()
                rows = []
                for line in obj.Товары:
                    rows.append({
                        "article": safe_str(getattr(line, "Номенклатура", "")),
                        "metal": safe_str(getattr(line, "Металл", "")),
                        "hallmark": safe_str(getattr(line, "Проба", "")),
                        "color": safe_str(getattr(line, "ЦветМеталла", "")),
                        "qty": getattr(line, "Количество", 0),
                        "w": getattr(line, "Вес", 0),
                        "comment": safe_str(getattr(line, "Комментарий", "")),
                    })
                result.append({
                    "num": str(obj.Number),
                    "date": str(obj.Date),
                    "contragent": safe_str(getattr(obj, "Контрагент", "")),
                    "rows": rows,
                })
            except Exception as e:
                print("[Order Parse Error]", e)
        return result
