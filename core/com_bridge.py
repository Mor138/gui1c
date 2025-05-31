import win32com.client
import pywintypes

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

    def get_ref(self, catalog_name: str, description: str):
        catalog = getattr(self.catalogs, catalog_name, None)
        if not catalog:
            return None
        selection = catalog.Select()
        while selection.Next():
            obj = selection.GetObject()
            if safe_str(obj.Description) == description:
                return obj.Ref
        return None

    def create_order(self, fields: dict, items: list) -> str:
        doc = self.documents.ЗаказВПроизводство.CreateDocument()

        for k, v in fields.items():
            if k.lower() == "номер":
                continue  # Номер нельзя устанавливать вручную
            try:
                setattr(doc, k, v)
            except Exception as e:
                print(f"[create_order] Failed to set {k}: {e}")

        for row in items:
            try:
                new_row = doc.Товары.Add()
                new_row.Номенклатура = row.get("Номенклатура")
                new_row.ВариантИзготовления = row.get("ВариантИзготовления")
                new_row.Размер = row.get("Размер")
                new_row.Количество = row.get("Количество", 1)
                new_row.Вес = row.get("Вес", 0)
            except Exception as e:
                print(f"[create_order row] Failed: {e}")

        try:
            doc.Write()  # Проводим без True, чтобы избежать дополнительных действий
            return str(doc.Number)  # Получаем присвоенный номер
        except Exception as e:
            print("[create_order error]", e)
            return "Ошибка"

    def list_orders(self, limit=5000):
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
                    "article": safe_str(getattr(line, "Номенклатура", "")),
                    "variant": safe_str(getattr(line, "ВариантИзготовления", "")),
                    "size": getattr(line, "Размер", 0),
                    "qty": getattr(line, "Количество", 0),
                    "w": getattr(line, "Вес", 0),
                })
            result.append({
                "num": str(obj.Number),
                "date": str(obj.Date),
                "contragent": safe_str(getattr(obj, "Контрагент", "")),
                "rows": rows,
            })
        return result

    def get_last_order_number(self) -> str:
        doc = getattr(self.documents, "ЗаказВПроизводство", None)
        if not doc:
            return "00ЮП-000000"
        selection = doc.Select()
        number = "00ЮП-000000"
        while selection.Next():
            try:
                obj = selection.GetObject()
                number = str(obj.Number)
            except Exception as e:
                print("[get_last_order_number error]", e)
                continue
        return number

    def get_next_order_number(self) -> str:
        last = self.get_last_order_number()
        try:
            prefix, num = last.split("-")
            next_num = int(num) + 1
            return f"{prefix}-{next_num:06d}"
        except:
            return "00ЮП-000001"

    def list_catalog_items(self, catalog_name: str, limit: int = 100000) -> list:
        result = []
        catalog = getattr(self.catalogs, catalog_name, None)
        if catalog is None:
            return result
        selection = catalog.Select()
        while selection.Next():
            obj = selection.GetObject()
            result.append({"Description": safe_str(obj.Description)})
        return result
