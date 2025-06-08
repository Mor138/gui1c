# com_bridge.py • взаимодействие с 1С через COM
import win32com.client
import pywintypes
import os
import tempfile
from typing import Any, Dict, List
from win32com.client import VARIANT
from pythoncom import VT_BOOL

# ---------------------------
# Маппинг описаний в системные имена перечисления
# ---------------------------
PRODUCTION_STATUS_MAP = {
    "Собств металл, собств камни": "СобствМеталлСобствКамни",
    "Собств металл, дав камни":    "СобствМеталлДавКамни",
    "Дав металл, собств камни":    "ДавМеталлСобствКамни",
    "Дав металл, дав камни":       "ДавМеталлДавКамни",
}

# ---------------------------
# Безопасное преобразование
# ---------------------------
def safe_str(val: Any) -> str:
    try:
        if val is None:
            return ""
        if hasattr(val, "GetPresentation"):
            try:
                return str(val.GetPresentation())  # ← скобки были обязательны!
            except Exception:
                pass
        for attr in ("Presentation", "Description", "Name", "Имя"):
            if hasattr(val, attr):
                return str(getattr(val, attr))
        return str(val)
    except Exception as e:
        return f"<error: {e}>"

def log(msg: str) -> None:
    print("[LOG]", msg)
    
   

class COM1CBridge:
    PRODUCTION_STATUSES = [
        "СобствМеталлСобствКамни",
        "СобствМеталлДавКамни",
        "ДавМеталлСобствКамни",
        "ДавМеталлДавКамни"
    ]
    
    def __init__(self, base_path, usr="Администратор", pwd=""):
        self.connector = win32com.client.Dispatch("V83.COMConnector")
        self.connection = self.connector.Connect(
            f'File="{base_path}";Usr="{usr}";Pwd="{pwd}"'
        )
        self.catalogs = self.connection.Catalogs
        self.documents = self.connection.Documents
        self.enums = self.connection.Enums
        
    def print_order_preview_pdf(self, number: str) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[Печать] Заказ №{number} не найден")
            return False
        try:
            form = obj.GetForm("ФормаДокумента")
            temp_dir = tempfile.gettempdir()
            pdf_path = os.path.join(temp_dir, f"Заказ_{number}.pdf")
            form.PrintFormToFile("Заказ в производство с фото", pdf_path)

            if os.path.exists(pdf_path):
                log(f"📄 PDF сформирован: {pdf_path}")
                os.startfile(pdf_path)  # Открытие в системе по умолчанию
                return True
            else:
                log(f"❌ Не удалось сохранить PDF")
                return False
        except Exception as e:
            log(f"❌ Ошибка при формировании PDF: {e}")
            return False     

    def _find_document_by_number(self, doc_name: str, number: str):
        doc = getattr(self.documents, doc_name, None)
        if not doc:
            log(f"[ERROR] Документ '{doc_name}' не найден")
            return None
        selection = doc.Select()
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Number).strip() == number.strip():
                return obj
        return None

    def undo_posting(self, number: str) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[UndoPosting] Документ №{number} не найден")
            return False
        try:
            obj.UndoPosting()
            obj.Write()
            log(f"✔ Проведение снято для заказа №{number}")
            return True
        except Exception as e:
            log(f"❌ UndoPosting error: {e}")
            return False

    def delete_order_by_number(self, number: str) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[Удаление] Заказ №{number} не найден")
            return False
        try:
            if getattr(obj, "Проведен", False):
                log(f"⚠ Документ проведён, снимаем проведение...")
                self.undo_posting(number)
            obj.Delete()
            log(f"🗑 Документ удалён полностью")
            return True
        except Exception as e:
            log(f"❌ Ошибка при удалении: {e}")
            return False

    def post_order(self, number: str) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[Проведение] Заказ №{number} не найден")
            return False
        try:
            obj.Проведен = True  # <- ЯВНО УСТАНАВЛИВАЕМ ФЛАГ ПРОВЕДЕНИЯ
            obj.Write()          # <- обычный Write без параметров
            if getattr(obj, "Проведен", False):
                log(f"[Проведение] Заказ №{number} успешно проведён через флаг Проведен")
                return True
            else:
                log(f"[Проведение] Не удалось провести заказ №{number}")
                return False
        except Exception as e:
            log(f"[Проведение] Ошибка при установке Проведен: {e}")
            return False


    def mark_order_for_deletion(self, number: str) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[Пометка] Документ №{number} не найден")
            return False
        try:
            if getattr(obj, "Проведен", False):
                log("⚠ Документ проведён. Снимаем проведение перед пометкой")
                obj.UndoPosting()
                obj.Write()
            obj.DeletionMark = VARIANT(VT_BOOL, True)
            obj.Write()
            if getattr(obj, "DeletionMark", False):
                log(f"🗑 Документ №{number} помечен на удаление")
                return True
            else:
                log(f"❌ Не удалось установить пометку на удаление")
                return False
        except Exception as e:
            log(f"❌ Ошибка при установке пометки: {e}")
            return False
            
    def unmark_order_deletion(self, number: str) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[Снятие пометки] Документ №{number} не найден")
            return False
        try:
            obj.DeletionMark = VARIANT(VT_BOOL, False)
            obj.Write()
            if not getattr(obj, "DeletionMark", True):
                log(f"✅ Пометка на удаление снята с документа №{number}")
                return True
            else:
                log(f"❌ Не удалось снять пометку")
                return False
        except Exception as e:
            log(f"❌ Ошибка при снятии пометки: {e}")
            return False        

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

    def get_size_ref(self, size_value):
        catalog = getattr(self.catalogs, "Размеры", None)
        if catalog is None:
            log("❌ Каталог 'Размеры' не найден")
            return None
        selection = catalog.Select()
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Description).strip().replace(",", ".") == str(size_value).strip().replace(",", "."):
                return obj.Ref
        log(f"❌ Размер '{size_value}' не найден в справочнике 'Размеры'")
        return None

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


    def get_ref(self, catalog_name, description):
        obj = self.get_catalog_object_by_description(catalog_name, description)
        return obj.Ref if hasattr(obj, "Ref") else obj

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
            
    def to_string(self, value):
        """Возвращает строковое представление значения через 1С Application"""
        try:
            return str(self.connection.String(value))
        except Exception as e:
            log(f"[to_string] Ошибка получения строки: {e}")
            return "[??]"     
            
    def get_catalog_object_by_description(self, catalog_name, description):
        if catalog_name == "ВидыСтатусыПродукции":
            predefined = {
                "Собств металл, собств камни": "СобствМеталлСобствКамни",
                "Собств металл, дав камни":    "СобствМеталлДавКамни",
                "Дав металл, собств камни":    "ДавМеталлСобствКамни",
                "Дав металл, дав камни":       "ДавМеталлДавКамни"
            }
            internal = predefined.get(description.strip())
            if internal:
                enum = getattr(self.enums, "ВидыСтатусыПродукции", None)
                if enum:
                    try:
                        val = getattr(enum, internal)
                        log(f"[{catalog_name}] Найден (Enum): {description} → {internal}")
                        return val
                    except Exception as e:
                        log(f"[Enum Error] {catalog_name}.{internal}: {e}")
            log(f"[{catalog_name}] Не найден по описанию: {description}")
            return None

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
            

    def get_production_status_variants(self) -> list[str]:
        return list(PRODUCTION_STATUS_MAP.keys()) 
        
    def print_order_with_photo(self, number: str):
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[Печать] Заказ №{number} не найден")
            return False
        try:
            form = obj.GetForm("ФормаДокумента")
            form.Open()  # Можно убрать, если не нужен показ формы
            form.PrintForm("Заказ в производство с фото")
            log(f"🖨 Печать формы 'Заказ в производство с фото' запущена")
            return True
        except Exception as e:
            log(f"❌ Ошибка печати: {e}")
            return False    

    def update_order(self, number: str, fields: dict, items: list) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number)
        if not obj:
            log(f"[Обновление] Заказ №{number} не найден")
            return False

        # Установка полей
        for k, v in fields.items():
            try:
                if k == "ВидСтатусПродукции":
                    ref = self.get_ref("ВидыСтатусыПродукции", v)
                    if ref:
                        setattr(obj, k, ref)
                    continue
                if k in ["Организация", "Контрагент", "ДоговорКонтрагента", "Ответственный", "Склад"]:
                    ref = self.get_ref(k + "ы" if not k.endswith("т") else k + "а", v)
                    if ref:
                        setattr(obj, k, ref)
                    continue
                setattr(obj, k, v)
            except Exception as e:
                log(f"[Обновление] Ошибка установки поля {k}: {e}")

        # Очищаем старые строки
        while len(obj.Товары) > 0:
            obj.Товары.Delete(0)

        # Добавляем новые строки
        for row in items:
            try:
                new_row = obj.Товары.Add()
                new_row.Номенклатура = self.get_ref("Номенклатура", row.get("Номенклатура"))
                variant = row.get("ВариантИзготовления")
                if variant and variant != "—":
                    new_row.ВариантИзготовления = self.get_ref("ВариантыИзготовленияНоменклатуры", variant)
                size_val = row.get("Размер", 0)
                new_row.Размер = self.get_size_ref(size_val)
                new_row.Количество = int(row.get("Количество", 1))
                new_row.Вес = float(row.get("Вес", 0))
                new_row.Примечание = row.get("Примечание", "")
            except Exception as e:
                log(f"[Обновление] Ошибка в строке заказа: {e}")

        try:
            obj.Write()
            log(f"✔ Обновлён заказ №{number}")
            return True
        except Exception as e:
            log(f"[Обновление] Ошибка при записи: {e}")
            return False        
   

    def create_order(self, fields, items):
        doc = self.documents.ЗаказВПроизводство.CreateDocument()
        catalog_fields_map = {
            "Организация": "Организации",
            "Контрагент": "Контрагенты",
            "ДоговорКонтрагента": "ДоговорыКонтрагентов",
            "Ответственный": "Пользователи",
            "Склад": "Склады",
            "ВидСтатусПродукции": "ВидыСтатусыПродукции"
        }

        # Справочники
        catalog_fields_map = {
            "Организация": "Организации",
            "Контрагент": "Контрагенты",
            "ДоговорКонтрагента": "ДоговорыКонтрагентов",
            "Ответственный": "Пользователи",
            "Склад": "Склады",
        }

        # ← ВАЖНО: внутри метода, чтобы можно было использовать self

        log("Создание заказа. Поля:")
        for k, v in fields.items():
            try:
                log(f"  -> {k} = {v}")
                if k == "Склад":
                    log("    ⚠ Поле 'Склад' не устанавливается вручную. Пропускаем.")
                    continue

                if k == "ВидСтатусПродукции":
                    # Если передано внутреннее имя — преобразуем в описание
                    reverse_map = {v: k for k, v in PRODUCTION_STATUS_MAP.items()}
                    if v in reverse_map:
                        v = reverse_map[v]
                    
                    ref = self.get_ref("ВидыСтатусыПродукции", v)
                    log(f"[ref] {k}: {v} => {ref}")
                    if ref:
                        setattr(doc, k, ref)
                        log(f"    Установлено: {k} (Ref: {ref})")
                    else:
                        log(f"    ❌ {k} '{v}' не найден.")
                    continue

                if k in catalog_fields_map:
                    ref = self.get_ref(catalog_fields_map[k], v)
                    log(f"[ref] {k}: {v} => {ref}")
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

                size_val = row.get("Размер", 0)
                size_ref = self.get_size_ref(size_val)
                if size_ref:
                    new_row.Размер = size_ref
                else:
                    log(f"    ❌ Размер '{size_val}' не найден, пропущен")

                new_row.Количество = int(row.get("Количество", 1))
                new_row.Вес = float(row.get("Вес", 0))
                log(f"    -> Примечание: {row.get('Примечание')}")
                new_row.Примечание = row.get("Примечание", "")

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

    # ------------------------------------------------------------------
    def create_wax_job(self, job: dict) -> str:
        """Создаёт документ "НарядВосковыеИзделия" на основании наряда."""
        try:
            doc = self.documents.НарядВосковыеИзделия.CreateDocument()
        except Exception as e:
            log(f"[1C] Не удалось создать документ НарядВосковыеИзделия: {e}")
            return ""

        try:
            doc.Дата = self.connection.ТекущаяДата()
            if hasattr(doc, "Комментарий"):
                doc.Комментарий = f"WX:{job.get('wax_job')} партия:{job.get('batch_code')}"
            if job.get("assigned_to") and hasattr(doc, "Ответственный"):
                ref = self.get_ref("Пользователи", job["assigned_to"])
                if ref:
                    doc.Ответственный = ref
            if hasattr(doc, "Количество"):
                doc.Количество = job.get("qty", 0)
            weight = job.get("weight_wax") or job.get("weight")
            if hasattr(doc, "Вес") and weight is not None:
                doc.Вес = float(weight)
            doc.Write()
            log(f"✅ Создан документ НарядВосковыеИзделия №{doc.Number}")
            return str(doc.Number)
        except Exception as e:
            log(f"❌ Ошибка создания НарядаВосковыеИзделия: {e}")
            return ""

    def list_orders(self):
        result = []
        orders = self.connection.Documents.ЗаказВПроизводство.Select()
        while orders.Next():
            doc = orders.GetObject()
            rows = []
            for row in doc.Товары:
                rows.append({
                    "nomenclature": safe_str(row.Номенклатура),
                    "size": safe_str(row.Размер),
                    "qty": row.Количество,
                    "w": row.Вес,
                    "variant": safe_str(row.ВариантИзготовления),
                    "note": row.Примечание,
                })
            result.append({
                "num": doc.Номер,
                "date": str(doc.Дата),
                "org": safe_str(doc.Организация),
                "contragent": safe_str(doc.Контрагент),
                "contract": safe_str(doc.ДоговорКонтрагента),
                "comment": safe_str(doc.Комментарий),
                "prod_status": self.to_string(doc.ВидСтатусПродукции),
                "posted": doc.Проведен,
                "deleted": doc.ПометкаУдаления,
                "qty": sum([r["qty"] for r in rows]),
                "weight": sum([r["w"] for r in rows]),
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

    # ------------------------------------------------------------------
    def create_task_from_order(self, order: dict) -> str:
        """Создаёт документ 'ЗаданиеНаПроизводство' на основании заказа."""
        try:
            doc = self.documents.ЗаданиеНаПроизводство.CreateDocument()
        except Exception as e:
            log(f"[1C] Не удалось создать ЗаданиеНаПроизводство: {e}")
            return ""

        try:
            doc.Дата = self.connection.ТекущаяДата()
            base = self._find_document_by_number("ЗаказВПроизводство", order.get("num", ""))
            if base:
                doc.ДокументОснование = base

            if order.get("assigned_to"):
                ref = self.get_ref("Пользователи", order["assigned_to"])
                if ref:
                    doc.РабочийЦентр = ref

            # Используем "восковка" в качестве участка и операции по умолчанию
            section = self.get_ref("ПроизводственныеУчастки", "восковка")
            if section:
                doc.ПроизводственныйУчасток = section
            op = self.get_ref("ТехОперации", "работа с восковыми изделиями")
            if op:
                doc.ТехОперация = op

            for row in order.get("rows", []):
                r = doc.Товары.Add()
                r.Номенклатура = self.get_ref("Номенклатура", row.get("article"))
                var = row.get("variant")
                if var:
                    r.ВариантИзготовления = self.get_ref("ВариантыИзготовленияНоменклатуры", var)
                r.Размер = self.get_size_ref(row.get("size"))
                r.Количество = row.get("qty", 0)
                r.Вес = row.get("weight", 0)
                r.Проба = str(row.get("hallmark", ""))
                r.ЦветМеталла = self.get_ref("ЦветаМеталлов", row.get("color"))
                r.АртикулГП = row.get("article")

            doc.Write()
            doc.Провести()
            log(f"✅ Создано ЗаданиеНаПроизводство №{doc.Number}")
            return str(doc.Number)
        except Exception as e:
            log(f"❌ Ошибка создания Задания: {e}")
            return ""

    # ------------------------------------------------------------------
    def create_wax_job_from_task(self, task_number: str) -> str:
        """Создаёт 'НарядВосковыеИзделия' на основании задания."""
        task = self._find_document_by_number("ЗаданиеНаПроизводство", task_number)
        if not task:
            log(f"❌ Задание №{task_number} не найдено")
            return ""

        try:
            doc = self.documents.НарядВосковыеИзделия.CreateDocument()
        except Exception as e:
            log(f"[1C] Не удалось создать НарядВосковыеИзделия: {e}")
            return ""

        try:
            doc.Дата = self.connection.ТекущаяДата()
            doc.ДокументОснование = task
            if hasattr(task, "ТехОперация"):
                doc.ТехОперация = task.ТехОперация
            if hasattr(task, "ПроизводственныйУчасток"):
                doc.ПроизводственныйУчасток = task.ПроизводственныйУчасток
            if hasattr(task, "РабочийЦентр"):
                doc.Ответственный = task.РабочийЦентр

            for row in task.Товары:
                r = doc.ТоварыВыдано.Add()
                r.Номенклатура = row.Номенклатура
                r.Размер = row.Размер
                r.Количество = row.Количество
                r.ВариантИзготовления = row.ВариантИзготовления
                r.Проба = row.Проба
                r.ЦветМеталла = row.ЦветМеталла
                r.Вес = row.Вес
                r.АртикулГП = row.АртикулГП

            doc.Write()
            doc.Провести()
            log(f"✅ Создан НарядВосковыеИзделия №{doc.Number}")
            return str(doc.Number)
        except Exception as e:
            log(f"❌ Ошибка создания Наряда: {e}")
            return ""
