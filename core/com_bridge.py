# com_bridge.py • взаимодействие с 1С через COM v0.8
# -*- coding: utf-8 -*-
import win32com.client
import pywintypes
import os
import tempfile
from typing import Any, Dict, List
from win32com.client import VARIANT
from pythoncom import VT_BOOL
from collections import defaultdict
from datetime import datetime, timedelta

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
        

        
    def get_wax_job_lines(self, doc_num: str) -> list[dict]:
        """Возвращает табличную часть 'ТоварыВыдано' по номеру наряда"""
        result = []
        doc_manager = getattr(self.connection.Documents, "НарядВосковыеИзделия", None)
        if not doc_manager:
            return result

        selection = doc_manager.Select()
        while selection.Next():
            doc = selection.GetObject()
            if str(doc.Number) == doc_num:
                for row in doc.ТоварыВыдано:
                    result.append({
                        "norm": self.safe(row, "ВидНорматива"),
                        "nomen": self.safe(row, "Номенклатура"),
                        "size": self.safe(row, "Размер"),
                        "sample": self.safe(row, "Проба"),
                        "color": self.safe(row, "ЦветМеталла"),
                        "qty": row.Количество,
                        "weight": row.Вес if hasattr(row, "Вес") else "",
                    })
                break
        return result    
        
    def safe(self, obj, attr):
        try:
            val = getattr(obj, attr, None)
            if val is None:
                return ""
            if hasattr(val, "GetPresentation"):
                return str(val.GetPresentation())
            if hasattr(val, "Presentation"):
                return str(val.Presentation)
            if hasattr(val, "Description"):
                return str(val.Description)
            return str(val)
        except Exception as e:
            return "<err>"    
        
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
        
    def _find_doc(self, doc_name: str, num: str):
        """Находит документ по имени и номеру"""
        try:
            docs = getattr(self.connection.Documents, doc_name)
        except AttributeError:
            raise Exception(f"Документ '{doc_name}' не найден")

        selection = docs.Select()
        while selection.Next():
            doc = selection.GetObject()
            if str(doc.Number).strip() == str(num).strip():
                return doc
        raise Exception(f"Документ '{doc_name}' с номером {num} не найден")    

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

    def get_ref_by_description(self, catalog_name: str, description: str):
        """Возвращает ссылку на элемент каталога по его описанию."""
        obj = self.get_catalog_object_by_description(catalog_name, description)
        if obj and hasattr(obj, "Ref"):
            return obj.Ref
        log(f"[get_ref_by_description] Не найден элемент '{description}' в каталоге '{catalog_name}'")
        return None

    def get_enum_by_description(self, enum_name: str, description: str):
        """Возвращает элемент перечисления по его представлению"""
        enum = getattr(self.enums, enum_name, None)
        if not enum:
            log(f"Перечисление '{enum_name}' не найдено")
            return None
        if description is None:
            return None

        desc = str(description).strip().lower()
        try:
            for attr in dir(enum):
                if attr.startswith("_"):
                    continue
                try:
                    val = getattr(enum, attr)
                except Exception:
                    continue
                pres = ""
                try:
                    if hasattr(val, "GetPresentation"):
                        pres = str(val.GetPresentation())
                    elif hasattr(val, "Presentation"):
                        pres = str(val.Presentation)
                    else:
                        pres = str(val)
                except Exception:
                    pres = str(val)

                if pres.strip().lower() == desc or attr.lower() == desc:
                    return val
        except Exception as e:
            log(f"[Enum] Ошибка поиска {description} в {enum_name}: {e}")

        log(f"[{enum_name}] Не найдено значение: {description}")
        return None

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
            
    def get_doc_object_by_number(self, doc_type: str, number: str):
        try:
            catalog = getattr(self.connection.Documents, doc_type)
            selection = catalog.Select()
            while selection.Next():
                doc = selection.GetObject()
                if str(doc.Number).strip() == str(number).strip():
                    log(f"[get_doc_object_by_number] ✅ Найден объект документа {doc_type} №{number}")
                    return doc
            log(f"[get_doc_object_by_number] ❌ Документ {doc_type} №{number} не найден")
        except Exception as e:
            log(f"[get_doc_object_by_number] ❌ Ошибка: {e}")
        return None

    def get_doc_ref(self, doc_name: str, number: str):
        """Возвращает ссылку на документ по номеру."""
        docs = getattr(self.connection.Documents, doc_name, None)
        if docs is None:
            log(f"[get_doc_ref] Документ '{doc_name}' не найден")
            return None

        selection = docs.Select()
        while selection.Next():
            doc = selection.GetObject()
            if str(doc.Number).strip() == str(number).strip():
                log(f"[get_doc_ref] ✅ Найден документ {doc_name} №{number}")
                return doc.Ref

        log(f"[get_doc_ref] ❌ Документ {doc_name} №{number} не найден")
        return None
    # ------------------------------------------------------------------
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
                "Ref": doc.Ref,  # ← вот правильное место для Ref
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

    def list_catalog_items(self, catalog_name: str, limit: int = 1000) -> list[dict]:
        """Возвращает список элементов справочника"""
        result = []
        try:
            catalog = getattr(self.connection.Catalogs, catalog_name, None)
            if catalog is None:
                self._log(f"[Catalog Error] Справочник '{catalog_name}' не найден")
                return result

            selection = catalog.Select()
            count = 0
            while selection.Next() and count < limit:
                obj = selection.GetObject()
                item = {
                    "Ref": str(obj.Ref),
                    "Code": str(obj.Code),
                    "Description": str(obj.Description)
                }
                result.append(item)
                count += 1
            return result
        except Exception as e:
            self._log(f"[Catalog Exception] {catalog_name}: {e}")
            return []
            
    def log_catalog_contents(self, catalog_name: str, limit: int = 1000):
        """Логирует все элементы указанного справочника по имени"""
        print(f"[Catalog Dump] Содержимое справочника '{catalog_name}':")
        items = self.list_catalog_items(catalog_name, limit)
        if not items:
            print("📭 Пусто или справочник не найден")
            return
        for item in items:
            print(f" - {item.get('Description', '?')} (Код: {item.get('Code', '?')}, Ref: {item.get('Ref', '?')})")       
        
    def list_tasks(self) -> list[dict]:
        result = []
        doc_manager = getattr(self.connection.Documents, "ЗаданиеНаПроизводство", None)
        if doc_manager is None:
            return result
        selection = doc_manager.Select()
        while selection.Next():
            doc = selection.GetObject()
            result.append({
                "ref": str(doc.Ref),
                "num": str(doc.Number),
                "date": str(doc.Date.strftime("%d.%m.%Y")),
                "employee": self.safe(doc, "Ответственный"),
                "tech_op": self.safe(doc, "ТехОперация"),
                "section": self.safe(doc, "ПроизводственныйУчасток"),
                "posted": getattr(doc, "Проведен", False),
                "deleted": getattr(doc, "ПометкаУдаления", False),
            })
        return result
        
    def detect_method_from_items(self, items: list[dict]) -> str:
        """Автоматически определяет метод производства по названию номенклатуры"""
        for row in items:
            name = row.get("Номенклатура", "").lower()
            if "д" in name or "d" in name:
                return "3D печать"
        return "Резина"    
        
    def find_production_task_ref_by_method(self, method: str) -> str | None:
        """Возвращает ссылку на первое задание по указанному методу."""
        method_ref = self.get_ref_by_description("ВариантыИзготовленияНоменклатуры", method)
        if not method_ref:
            print_catalog_contents("ВариантыИзготовленияНоменклатуры")
        if method_ref is None:
            log(f"[find_production_task_ref_by_method] Не найден вариант {method}")
            return None

        doc_manager = getattr(self.connection.Documents, "ЗаданиеНаПроизводство", None)
        if doc_manager is None:
            log("[find_production_task_ref_by_method] Документ 'ЗаданиеНаПроизводство' не найден")
            return None

        tasks = doc_manager.Select()
        while tasks.Next():
            obj = tasks.GetObject()
            for row in obj.Товары:
                if row.ВариантИзготовления == method_ref:
                    return str(obj.Ref)
        return None
        

    def list_wax_jobs(self) -> list[dict]:
        result = []
        doc_manager = getattr(self.connection.Documents, "НарядВосковыеИзделия", None)
        if doc_manager is None:
            return result

        selection = doc_manager.Select()
        while selection.Next():
            doc = selection.GetObject()
            result.append({
                "ref": str(doc.Ref),
                "num": str(doc.Number),
                "date": str(doc.Date.strftime("%d.%m.%Y")),
                "status": self.safe(doc, "Статус"),
                "employee": self.safe(doc, "Сотрудник"),
                "organization": self.safe(doc, "Организация"),
                "warehouse": self.safe(doc, "Склад"),
                "section": self.safe(doc, "ПроизводственныйУчасток"),
                "tech_op": self.safe(doc, "ТехОперация"),
                "based_on": self.safe(getattr(doc, "ДокументОснование", None), "Number")
            })
        return result
        
    def get_ref_by_description(self, catalog_name: str, description: str):
        catalog = getattr(self.connection.Catalogs, catalog_name, None)
        if catalog is None:
            log(f"[get_ref_by_description] Каталог '{catalog_name}' не найден")
            return None

        selection = catalog.Select()
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Description).strip().lower() == description.strip().lower():
                return obj.Ref
        log(f"[get_ref_by_description] Не найден элемент '{description}' в каталоге '{catalog_name}'")
        return None    
        
        
        
    def create_production_task(self, order_ref, rows: list[dict]) -> dict:
        doc_manager = getattr(self.connection.Documents, "ЗаданиеНаПроизводство", None)
        if doc_manager is None:
            log("❌ Документ 'ЗаданиеНаПроизводство' не найден")
            return {}

        if not order_ref:
            log("❌ order_ref = None. Задание не может быть создано.")
            return {}

        try:
            # 🎯 Получаем объект документа-заказа
            if hasattr(order_ref, "GetObject"):
                base_doc = order_ref.GetObject()
            elif hasattr(order_ref, "Ref"):
                base_doc = self.connection.GetObject(order_ref.Ref)
            elif isinstance(order_ref, str):
                base_doc = self._find_document_by_number("ЗаказВПроизводство", order_ref)
                if base_doc is None:
                    log(f"[create_production_task] ❌ Не удалось найти заказ №{order_ref}")
                    return {}
            else:
                log("❌ order_ref — неизвестного типа")
                return {}

            base_doc_ref = base_doc.Ref  # ссылка на заказ

            doc = doc_manager.CreateDocument()
            doc.Дата = datetime.now()
            doc.КонечнаяДатаЗадания = datetime.now() + timedelta(days=1)
            doc.ДокументОснование = base_doc_ref

            if hasattr(doc, "Заказ"):
                try:
                    doc.Заказ = base_doc_ref
                except Exception as e:
                    log(f"[create_production_task] ⚠ Не удалось установить поле 'Заказ': {e}")

            if hasattr(doc, "ЗаказВПроизводство"):
                try:
                    doc.ЗаказВПроизводство = base_doc_ref
                except Exception as e:
                    log(f"[create_production_task] ⚠ Не удалось установить поле 'ЗаказВПроизводство': {e}")

            if hasattr(base_doc, "Организация") and hasattr(doc, "Организация"):
                try:
                    doc.Организация = base_doc.Организация
                except Exception as e:
                    log(f"[create_production_task] ⚠ Не удалось установить организацию: {e}")

            if hasattr(base_doc, "Склад") and hasattr(doc, "Склад"):
                try:
                    doc.Склад = base_doc.Склад
                except Exception as e:
                    log(f"[create_production_task] ⚠ Не удалось установить склад: {e}")

            doc.ПроизводственныйУчасток = self.get_ref("ПроизводственныеУчастки", "задание на производство")
            doc.ТехОперация = self.get_ref("ТехОперации", "работа с восковыми изделиями")
            employee_name = rows[0].get("employee", "Администратор")
            doc.РабочийЦентр = self.get_ref("ФизическиеЛица", employee_name)
            doc.Ответственный = self.get_ref("Пользователи", "Администратор")
            doc.Комментарий = "Создано автоматически из GUI"

            date_start = datetime.now()
            date_end = date_start + timedelta(days=1)

            for row in rows:
                try:
                    item = doc.Продукция.Add()
                    item.Номенклатура = self.get_ref("Номенклатура", row.get("name", ""))
                    item.Размер = self.get_ref("Размеры", row.get("size", ""))
                    item.Проба = self.get_ref("Пробы", row.get("assay", ""))
                    item.ЦветМеталла = self.get_ref("ЦветаМеталла", row.get("color", ""))
                    item.ХарактеристикаВставок = self.get_ref("ХарактеристикиВставок", row.get("insert", ""))
                    item.ВариантИзготовления = self.get_ref("ВариантыИзготовленияНоменклатуры", row.get("method", ""))
                    item.Количество = row.get("qty", 0)
                    item.Вес = float(row.get("weight", 0) or 0)
                    item.ДатаНачала = date_start
                    item.ДатаОкончания = date_end
                    item.РабочийЦентр = self.get_ref("ФизическиеЛица", employee_name)
                    item.Заказ = base_doc_ref
                    item.КонечнаяПродукция = item.Номенклатура
                    item.ВариантИзготовленияПродукции = item.ВариантИзготовления
                    if hasattr(item, "АртикулГП"):
                        item.АртикулГП = row.get("article", "")
                except Exception as e:
                    log(f"❌ Ошибка в строке 'Продукция': {e}")

            for row in rows:
                try:
                    z = doc.ЗаданияНаВыполнениеТехОперации.Add()
                    z.Вес = float(row.get("weight", 0) or 0)
                    z.Заказ = base_doc_ref
                    z.ТехОперация = self.get_ref("ТехОперации", "работа с восковыми изделиями")
                    z.РабочийЦентр = self.get_ref("ФизическиеЛица", employee_name)
                    z.Номенклатура = self.get_ref("Номенклатура", row.get("name", ""))
                    z.ВариантИзготовления = self.get_ref("ВариантыИзготовленияНоменклатуры", row.get("method", ""))
                    z.Размер = self.get_ref("Размеры", row.get("size", ""))
                    z.Проба = self.get_ref("Пробы", row.get("assay", ""))
                    z.ЦветМеталла = self.get_ref("ЦветаМеталла", row.get("color", ""))
                    z.ХарактеристикаВставок = self.get_ref("ХарактеристикиВставок", row.get("insert", ""))
                    z.Количество = row.get("qty", 0)
                    z.ДатаНачала = date_start
                    z.ДатаОкончания = date_end
                    z.КонечнаяПродукция = z.Номенклатура
                    z.ВариантИзготовленияПродукции = z.ВариантИзготовления
                except Exception as e:
                    log(f"❌ Ошибка в строке 'ЗаданияНаВыполнениеТехОперации': {e}")

            doc.Write()
            log(f"✅ Задание создано: №{doc.Номер}")
            return {
                "Ref": doc.Ref,
                "Номер": str(doc.Номер),
                "Дата": str(doc.Дата)
            }

        except Exception as e:
            log(f"❌ Ошибка при создании задания: {e}")
            return {}
            
    def calculate_batches(self, order_lines: list[dict]) -> list[dict]:
        from collections import defaultdict

        rows_by_batch = defaultdict(lambda: {"qty": 0, "total_w": 0.0})
        for row in order_lines:
            key = (row.get("metal"), row.get("assay"), row.get("color"))
            rows_by_batch[key]["qty"] += row.get("qty", 0)
            rows_by_batch[key]["total_w"] += row.get("weight", 0.0)

        result = []
        for (metal, hallmark, color), data in rows_by_batch.items():
            result.append({
                "batch_barcode": "AUTO",
                "metal": metal,
                "hallmark": hallmark,
                "color": color,
                "qty": data["qty"],
                "total_w": round(data["total_w"], 3)
            })
        return result      
            
    def get_catalog_ref(self, catalog_name, description):
        try:
            catalog = getattr(self.connection.Catalogs, catalog_name, None)
            if not catalog:
                log(f"Каталог '{catalog_name}' не найден")
                return None
            selection = catalog.Select()
            while selection.Next():
                item = selection.GetObject()
                if safe_str(item.Description) == description or safe_str(item) == description:
                    log(f"[{catalog_name}] Найден: {description}")
                    return item.Ref
            log(f"[{catalog_name}] Не найден: {description}")
        except Exception as e:
            log(f"[{catalog_name}] Ошибка: {e}")
        return None
        
    def get_wax_job_rows(self, num: str) -> list[dict]:
        doc = self._find_doc("НарядВосковыеИзделия", num)
        rows = []

        for r in doc.Выдано:  # <-- важно: корректное имя табличной части
            rows.append({
                "Номенклатура": self.safe(r, "Номенклатура"),
                "Размер": self.safe(r, "Размер"),
                "Проба": self.safe(r, "Проба"),
                "Цвет": self.safe(r, "ЦветМеталла"),
                "Количество": r.Количество,
                "Вес": r.Вес,
                "Партия": self.safe(r, "Партия"),
                "Номер ёлки": r.НомерЕлки if hasattr(r, "НомерЕлки") else "",
                "Состав набора": r.СоставНабора if hasattr(r, "СоставНабора") else ""
            })

        return rows   
           
        
    def get_order_lines(self, doc_number: str) -> list[dict]:
        doc = self._find_document_by_number("ЗаказВПроизводство", doc_number)
        if not doc:
            log(f"❌ Заказ №{doc_number} не найден")
            return []

        rows = []
        for r in doc.Товары:
            rows.append({
                "name": safe_str(r.Номенклатура),
                "size": safe_str(getattr(r, "Размер", "")),
                "insert": safe_str(getattr(r, "ХарактеристикаВставок", "")),
                "assay": safe_str(getattr(r, "Проба", "")),
                "color": safe_str(getattr(r, "ЦветМеталла", "")),
                "method": safe_str(getattr(r, "ВариантИзготовления", "")),
                "qty": getattr(r, "Количество", 0),
                "weight": float(getattr(r, "Вес", 0)),  # ← добавлено
                "article": safe_str(getattr(r.Номенклатура, "Артикул", ""))
            })
        return rows   
        
    def get_task_lines(self, doc_num: str) -> list[dict]:
        """Возвращает табличную часть 'Продукция' по номеру задания"""
        result = []
        doc_manager = getattr(self.connection.Documents, "ЗаданиеНаПроизводство", None)
        if not doc_manager:
            return result

        selection = doc_manager.Select()
        while selection.Next():
            doc = selection.GetObject()
            if str(doc.Number) == doc_num:
                for row in doc.Продукция:
                    result.append({
                        "nomen": self.safe(row, "Номенклатура"),
                        "size": self.safe(row, "Размер"),
                        "sample": self.safe(row, "Проба"),
                        "color": self.safe(row, "ЦветМеталла"),
                        "qty": row.Количество,
                        "weight": row.Вес if hasattr(row, "Вес") else ""
                    })
                break
        return result

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
            from datetime import datetime
            doc.Date = datetime.now()
            doc.ДокументОснование = task
            doc.ТехОперация = getattr(task, "ТехОперация", None)
            doc.ПроизводственныйУчасток = getattr(task, "ПроизводственныйУчасток", None)
            doc.Ответственный = getattr(task, "РабочийЦентр", None)

            for row in task.Продукция:
                r = doc.ТоварыВыдано.Add()
                r.Номенклатура = row.Номенклатура
                r.Размер = row.Размер
                r.Количество = row.Количество
                r.ВариантИзготовления = row.ВариантИзготовления
                r.Проба = row.Проба
                r.ЦветМеталла = row.ЦветМеталла
                r.Вес = row.Вес

            doc.Write()
            doc.Провести()
            log(f"✅ Создан НарядВосковыеИзделия №{doc.Number}")
            return str(doc.Number)
        except Exception as e:
            log(f"❌ Ошибка создания Наряда: {e}")
            return ""
            
    def _get_object_from_ref(self, ref):
        try:
            return self.connection.GetObject(ref)
        except Exception as e:
            log(f"[get_object_from_ref] ❌ Ошибка получения объекта по ссылке: {e}")
            return None

    def get_object_from_ref(self, ref):
        try:
            obj = ref.GetObject()
            log(f"[get_object_from_ref] ✅ Получен объект из ссылки")
            return obj
        except Exception as e:
            log(f"[get_object_from_ref] ❌ Ошибка получения объекта по ссылке: {e}")
            return None
    # ------------------------------------------------------------------
    def create_multiple_wax_jobs_from_task(self, task_ref, method_to_employee: dict) -> list[str]:
        """Создаёт по заданию два наряда: для 3D и для резины."""
        result = []

        log(f"[create_jobs] type(task_ref) = {type(task_ref)}")
        log(
            f"[create_jobs] hasattr 'Организация': {hasattr(task_ref, 'Организация')} "
            f"hasattr 'Organization': {hasattr(task_ref, 'Organization')}"
        )

        try:
            organization = getattr(task_ref, "Организация", None)
            log(f"[create_jobs] Организация задания: {safe_str(organization)}")
        except Exception as e:
            log(f"❌ Ошибка получения Организации из задания: {e}")
            return []

        try:
            if isinstance(task_ref, str):                     # передали строку-Ref
                task = self.connection.GetObject(task_ref)

            elif hasattr(task_ref, "Продукция"):             # уже полноценный объект
                task = task_ref

            elif hasattr(task_ref, "GetObject"):             # DocumentRef → поднимаем
                task = task_ref.GetObject()

            else:
                log("[create_jobs] ❌ Неизвестный тип ссылки на задание")
                return []
        except Exception as e:
            log(f"[create_jobs] ❌ Ошибка получения объекта задания: {e}")
            return []

        rows_by_method = defaultdict(list)
        for row in task.Продукция:
            name = safe_str(row.Номенклатура).lower()
            method = "3D печать" if "д" in name or "d" in name else "Резина"
            rows_by_method[method].append(row)

        for method, rows in rows_by_method.items():
            try:
                job = self.documents.НарядВосковыеИзделия.CreateDocument()
                job.Дата = datetime.now()

                if organization:
                    job.Организация = organization
                wh = getattr(task, "Склад", None)
                if wh:
                    job.Склад = wh
                job.ПроизводственныйУчасток = task.ПроизводственныйУчасток
                job.ЗаданиеНаПроизводство = task
                job.ТехОперация = self.get_ref("ТехОперации", "3D" if method == "3D печать" else "Пресс-форма")

                master_name = method_to_employee.get(method)
                job.Сотрудник = self.get_ref("ФизическиеЛица", master_name)
                job.Комментарий = f"Создан автоматически для {method}"

                for r in rows:
                    row = job.Товары.Add()
                    row.Номенклатура = r.Номенклатура
                    row.Количество = r.Количество
                    row.Размер = r.Размер
                    row.ВариантИзготовления = r.ВариантИзготовления
                    row.Проба = r.Проба
                    row.ЦветМеталла = r.ЦветМеталла
                    if hasattr(r, "ХарактеристикаВставок"):
                        row.ХарактеристикаВставок = r.ХарактеристикаВставок
                    row.Заказ = r.Заказ
                    row.КонечнаяПродукция = r.КонечнаяПродукция
                    if hasattr(r, "Вес"):
                        row.Вес = r.Вес

                job.Write()
                result.append(str(job.Номер))
                log(f"[create_job] ✅ Наряд для {method} создан: №{job.Номер}")
            except Exception as e:
                log(f"[create_job] ❌ Ошибка для {method}: {e}")
        return result

    def _find_task_by_number(self, number: str):
        doc_manager = getattr(self.connection.Documents, "ЗаданиеНаПроизводство", None)
        if doc_manager is None:
            log("❌ Документ 'ЗаданиеНаПроизводство' не найден")
            return None
        selection = doc_manager.Select()
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Number).strip() == number.strip():
                return obj
        return None

    def post_task(self, number: str) -> bool:
        obj = self._find_task_by_number(number)
        if not obj:
            log(f"[Проведение] ❌ Задание №{number} не найдено")
            return False
        try:
            obj.Проведен = True
            obj.Write()
            log(f"[Проведение] ✅ Задание №{number} проведено")
            return True
        except Exception as e:
            log(f"❌ Ошибка при проведении задания №{number}: {e}")
            return False

    def undo_post_task(self, number: str) -> bool:
        obj = self._find_task_by_number(number)
        if not obj:
            log(f"[Снятие проведения] ❌ Задание №{number} не найдено")
            return False
        try:
            obj.Проведен = False
            obj.Write()
            log(f"[Снятие проведения] ✅ Задание №{number} отменено")
            return True
        except Exception as e:
            log(f"❌ Ошибка при снятии проведения задания №{number}: {e}")
            return False

    def mark_task_for_deletion(self, number: str) -> bool:
        obj = self._find_task_by_number(number)
        if not obj:
            log(f"[Пометка удаления] ❌ Задание №{number} не найдено")
            return False
        try:
            obj.DeletionMark = True
            obj.Write()
            log(f"[Пометка удаления] ✅ Задание №{number} помечено на удаление")
            return True
        except Exception as e:
            log(f"❌ Ошибка при пометке на удаление задания №{number}: {e}")
            return False

    def unmark_task_deletion(self, number: str) -> bool:
        obj = self._find_task_by_number(number)
        if not obj:
            log(f"[Снятие пометки] ❌ Задание №{number} не найдено")
            return False
        try:
            obj.DeletionMark = False
            obj.Write()
            log(f"[Снятие пометки] ✅ Задание №{number} восстановлено")
            return True
        except Exception as e:
            log(f"❌ Ошибка при снятии пометки задания №{number}: {e}")
            return False

    def delete_task(self, number: str) -> bool:
        obj = self._find_task_by_number(number)
        if not obj:
            log(f"[Удаление] ❌ Задание №{number} не найдено")
            return False
        try:
            # Обязательно снять проведение, если стоит
            if getattr(obj, "Проведен", False):
                obj.Проведен = False
                obj.Write()

            # Пометить на удаление (иначе 1С не даст удалить)
            obj.DeletionMark = True
            obj.Write()

            # Теперь можно удалить
            obj.Delete()
            log(f"[Удаление] ✅ Задание №{number} удалено")
            return True
        except Exception as e:
            log(f"❌ Ошибка при удалении задания №{number}: {e}")
            return False
            
            