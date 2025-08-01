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

from .logger import logger
import config

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
    """Выводит сообщение через модуль logging с определением уровня."""
    text = str(msg)
    lower = text.lower()
    if "❌" in text or "error" in lower or "ошибка" in lower:
        logger.error(text)
    elif "⚠" in text or "warning" in lower:
        logger.warning(text)
    else:
        logger.info(text)



from .orders_bridge import OrdersBridge
from .wax_bridge import WaxBridge

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
        # Кэш ссылок на элементы справочников
        self._ref_cache: dict[str, dict[str, Any]] = {}

        # Разделённые мосты для разных страниц
        self.orders_bridge = OrdersBridge(self)
        self.wax_bridge = WaxBridge(self)
        

        
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
                        "weight": round(float(row.Вес), config.WEIGHT_DECIMALS) if hasattr(row, "Вес") else "",
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
        
    def print_order_preview_pdf(self, number: str, date: str | None = None) -> bool:
        """Формирует PDF через OrdersBridge."""
        return self.orders_bridge.print_order_preview_pdf(number, date)

    def _find_document_by_number(self, doc_name: str, number: str, date: str | None = None):
        doc = getattr(self.documents, doc_name, None)
        if not doc:
            log(f"[ERROR] Документ '{doc_name}' не найден")
            return None
        selection = doc.Select()
        found = None
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Number).strip() == number.strip():
                if date:
                    if str(obj.Date)[:10] == str(date)[:10]:
                        return obj
                elif not found or obj.Date > found.Date:
                    found = obj
        return found
        
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

    def undo_posting(self, number: str, date: str | None = None) -> bool:
        """Снимает проведение заказа через OrdersBridge."""
        return self.orders_bridge.undo_posting(number, date)

    def delete_order_by_number(self, number: str, date: str | None = None) -> bool:
        """Удаляет заказ через OrdersBridge."""
        return self.orders_bridge.delete_order_by_number(number, date)

    def post_order(self, number: str, date: str | None = None) -> bool:
        """Проводит заказ через OrdersBridge."""
        return self.orders_bridge.post_order(number, date)
            

    def mark_order_for_deletion(self, number: str, date: str | None = None) -> bool:
        """Помечает заказ на удаление через OrdersBridge."""
        return self.orders_bridge.mark_order_for_deletion(number, date)
            
    def unmark_order_deletion(self, number: str, date: str | None = None) -> bool:
        """Снимает пометку удаления заказа через OrdersBridge."""
        return self.orders_bridge.unmark_order_deletion(number, date)

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
        """Возвращает ссылку на размер с учётом кеша."""
        desc = str(size_value).strip().replace(",", ".")
        ref = self.get_ref_by_description("Размеры", desc)
        if not ref:
            log(f"❌ Размер '{size_value}' не найден в справочнике 'Размеры'")
        return ref

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
        """Возвращает ссылку на элемент справочника или перечисления."""
        return self.get_ref_by_description(catalog_name, description)


    def get_enum_by_description(self, enum_name: str, description: str):
        """Возвращает элемент перечисления по его представлению"""
        enum = getattr(self.enums, enum_name, None)
        if not enum:
            log(f"Перечисление '{enum_name}' не найдено")
            return None
        if description is None:
            return None

        desc = str(description).strip()
        # Попытка прямого доступа по имени атрибута
        try:
            if hasattr(enum, desc):
                return getattr(enum, desc)
        except Exception:
            pass

        low_desc = desc.lower()
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

                if pres.strip().lower() == low_desc or attr.lower() == low_desc:
                    return val
        except Exception as e:
            log(f"[Enum] Ошибка поиска {description} в {enum_name}: {e}")

        log(f"[{enum_name}] Не найдено значение: {description}")
        return None

    def list_enum_values(self, enum_name: str) -> list[str]:
        """Возвращает список представлений элементов перечисления."""
        enum = getattr(self.enums, enum_name, None)
        if not enum:
            log(f"Перечисление '{enum_name}' не найдено")
            return []
        values: list[str] = []
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
                if pres:
                    values.append(pres)
        except Exception as e:
            log(f"[list_enum_values] Ошибка чтения {enum_name}: {e}")
        return values

    def get_last_order_number(self):
        """Возвращает последний номер заказа через OrdersBridge."""
        return self.orders_bridge.get_last_order_number()

    def get_next_order_number(self):
        """Возвращает следующий номер заказа через OrdersBridge."""
        return self.orders_bridge.get_next_order_number()

    # ------------------------------------------------------------------
    def get_last_task_number(self):
        """Возвращает номер последнего документа 'ЗаданиеНаПроизводство'."""
        doc = getattr(self.documents, "ЗаданиеНаПроизводство", None)
        if not doc:
            return "ТП-000000"
        selection = doc.Select()
        number = "ТП-000000"
        while selection.Next():
            try:
                obj = selection.GetObject()
                number = str(obj.Number)
            except Exception:
                continue
        return number

    def get_next_task_number(self):
        """Возвращает следующий номер задания на производство."""
        last = self.get_last_task_number()
        try:
            prefix, num = last.split("-")
            next_num = int(num) + 1
            return f"{prefix}-{next_num:06d}"
        except Exception:
            return "ТП-000001"
            
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

    def update_order(self, number: str, fields: dict, items: list, date: str | None = None) -> bool:
        obj = self._find_document_by_number("ЗаказВПроизводство", number, date)
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
        # Справочники (поле 'ВидСтатусПродукции' обрабатывается отдельно)
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

                if k == "ВидСтатусПродукции":
                    # Если передано внутреннее имя — преобразуем в описание
                    reverse_map = {val: key for key, val in PRODUCTION_STATUS_MAP.items()}
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
            
    def get_doc_object_by_number(self, doc_type: str, number: str, date: str | None = None):
        try:
            catalog = getattr(self.connection.Documents, doc_type)
            selection = catalog.Select()
            while selection.Next():
                doc = selection.GetObject()
                if str(doc.Number).strip() == str(number).strip():
                    if date is None or str(doc.Date)[:10] == str(date)[:10]:
                        log(f"[get_doc_object_by_number] ✅ Найден объект документа {doc_type} №{number}")
                        return doc
            log(f"[get_doc_object_by_number] ❌ Документ {doc_type} №{number} не найден")
        except Exception as e:
            log(f"[get_doc_object_by_number] ❌ Ошибка: {e}")
        return None

    def get_doc_ref(self, doc_name: str, number: str, date: str | None = None):
        """Возвращает ссылку на документ по номеру."""
        docs = getattr(self.connection.Documents, doc_name, None)
        if docs is None:
            log(f"[get_doc_ref] Документ '{doc_name}' не найден")
            return None

        selection = docs.Select()
        while selection.Next():
            doc = selection.GetObject()
            if str(doc.Number).strip() == str(number).strip():
                if date is None or str(doc.Date)[:10] == str(date)[:10]:
                    log(f"[get_doc_ref] ✅ Найден документ {doc_name} №{number}")
                    return doc.Ref

        log(f"[get_doc_ref] ❌ Документ {doc_name} №{number} не найден")
        return None

    def list_documents(self, doc_name: str) -> list:
        """Возвращает список объектов документов указанного типа."""
        result: list = []
        docs = getattr(self.connection.Documents, doc_name, None)
        if docs is None:
            log(f"[list_documents] Документ '{doc_name}' не найден")
            return result

        selection = docs.Select()
        while selection.Next():
            try:
                obj = selection.GetObject()
                result.append(obj)
            except Exception as e:
                log(f"[list_documents] ❌ Ошибка: {e}")
        return result
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
                log(f"[Catalog Error] Справочник '{catalog_name}' не найден")
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
            log(f"[Catalog Exception] {catalog_name}: {e}")
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
            self.log_catalog_contents("ВариантыИзготовленияНоменклатуры")
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
        """Возвращает список нарядов на восковые изделия из 1С."""
        result = []
        docs = self.connection.Documents.НарядВосковыеИзделия.Select()
        while docs.Next():
            obj = docs.GetObject()


            rows = getattr(obj, "ТоварыВыдано", [])
            rows = getattr(obj, "Выдано", [])
            weight = 0.0
            qty = 0
            for r in rows:
                weight += getattr(r, "Вес", 0)
                qty += getattr(r, "Количество", 0)

            result.append({
                "Номер": str(obj.Номер),
                "Дата": obj.Дата.strftime("%d.%m.%Y"),
                "Сотрудник": safe_str(obj.Сотрудник.Description),
                "Вес": weight,
                "Кол-во": qty,
                "Проведен": obj.Проведен,
                "ПометкаУдаления": getattr(obj, "ПометкаУдаления", False),
                "Закрыт": "✅" if getattr(obj, "Проведен", False) else "—",
                "Сотрудник": safe_str(obj.Сотрудник),
                "ТехОперация": safe_str(obj.ТехОперация),
                "Комментарий": safe_str(obj.Комментарий),
                "Склад": safe_str(obj.Склад),
                "ПроизводственныйУчасток": safe_str(obj.ПроизводственныйУчасток),
                "Организация": safe_str(obj.Организация),
                "Задание": safe_str(getattr(obj, "ЗаданиеНаПроизводство", "")) or "—",
                "Ответственный": safe_str(obj.Ответственный),
                "Вес": round(weight, 2),
                "Кол-во": qty,
            })
        return result

    def find_wax_jobs_by_task(self, task_ref) -> list:
        """Возвращает наряды, связанные с заданием, через WaxBridge."""
        return self.wax_bridge.find_wax_jobs_by_task(task_ref)

    def close_wax_jobs(self, job_refs: list) -> list[str]:
        """Закрывает наряды через WaxBridge."""
        return self.wax_bridge.close_wax_jobs(job_refs)
        
    def get_ref_by_description(self, catalog_name: str, description: str):
        """Возвращает ссылку на элемент каталога по описанию с кешированием."""
        if catalog_name == "ВидыСтатусыПродукции":
            predefined = {
                "Собств металл, собств камни": "СобствМеталлСобствКамни",
                "Собств металл, дав камни":    "СобствМеталлДавКамни",
                "Дав металл, собств камни":    "ДавМеталлСобствКамни",
                "Дав металл, дав камни":       "ДавМеталлДавКамни",
            }
            internal = predefined.get(str(description).strip())
            if internal:
                enum = getattr(self.enums, "ВидыСтатусыПродукции", None)
                if enum is not None:
                    try:
                        return getattr(enum, internal)
                    except Exception as e:
                        log(f"[Enum Error] {catalog_name}.{internal}: {e}")
            log(f"[{catalog_name}] Не найден по описанию: {description}")
            return None

        desc = str(description).strip().lower()
        if catalog_name == "Размеры":
            desc = desc.replace(",", ".")

        cache = self._ref_cache.get(catalog_name)
        if cache is None:
            catalog = getattr(self.connection.Catalogs, catalog_name, None)
            if catalog is None:
                log(f"[get_ref_by_description] Каталог '{catalog_name}' не найден")
                return None
            cache = {}
            selection = catalog.Select()
            while selection.Next():
                obj = selection.GetObject()
                key = str(obj.Description).strip().lower()
                if catalog_name == "Размеры":
                    key = key.replace(",", ".")
                cache[key] = obj.Ref
            self._ref_cache[catalog_name] = cache

        ref = cache.get(desc)
        if ref is None:
            log(f"[get_ref_by_description] Не найден элемент '{description}' в каталоге '{catalog_name}'")
        return ref
        
        
        
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
            operation = rows[0].get("operation", "работа с восковыми изделиями")
            doc.ТехОперация = self.get_ref("ТехОперации", operation)
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
                    z.ТехОперация = self.get_ref("ТехОперации", row.get("operation", operation))
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
                "total_w": round(data["total_w"], config.WEIGHT_DECIMALS)
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
                "Вес": round(float(r.Вес), config.WEIGHT_DECIMALS),
                "Партия": self.safe(r, "Партия"),
                "Номер ёлки": r.НомерЕлки if hasattr(r, "НомерЕлки") else "",
                "Состав набора": r.СоставНабора if hasattr(r, "СоставНабора") else ""
            })

        return rows   
           
        
    def get_order_lines(self, doc_number: str, date: str | None = None) -> list[dict]:
        doc = self._find_document_by_number("ЗаказВПроизводство", doc_number, date)
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
                        "article": safe_str(getattr(row.Номенклатура, "Артикул", "")),
                        "size": self.safe(row, "Размер"),
                        "sample": self.safe(row, "Проба"),
                        "color": self.safe(row, "ЦветМеталла"),
                        "qty": row.Количество,
                        "weight": row.Вес if hasattr(row, "Вес") else ""
                    })
                break
        return result

    def get_wax_job_lines_by_ref(self, ref):
        """Возвращает табличную часть 'ТоварыВыдано' по ссылке на наряд"""
        try:
            doc = self.get_object_from_ref(ref)
            result = []
            for row in doc.ТоварыВыдано:
                result.append({
                    "nomen": self.safe(row, "Номенклатура"),
                    "size": self.safe(row, "Размер"),
                    "sample": self.safe(row, "Проба"),
                    "color": self.safe(row, "ЦветМеталла"),
                    "qty": row.Количество,
                    "weight": round(float(row.Вес), config.WEIGHT_DECIMALS) if hasattr(row, "Вес") else "",
                })
            return result
        except Exception as e:
            log(f"[get_wax_job_lines_by_ref] ❌ Ошибка: {e}")
            return []

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

            # Получаем заказ в производстве — строго через GetObject
            order_ref = (
                getattr(task, "ЗаказВПроизводство", None)
                or getattr(task, "ДокументОснование", None)
            )
            order_obj = None
            if order_ref and hasattr(order_ref, "GetObject"):
                try:
                    order_obj = order_ref.GetObject()
                    log("[create_wax_job_from_task] ✅ Получен заказ-основание")
                except Exception as exc:
                    log(f"[create_wax_job_from_task] ⚠ Ошибка получения заказа: {exc}")

            # Подстановка организации и склада ТОЛЬКО из заказа
            if order_obj:
                try:
                    org = getattr(order_obj, "Организация", None)
                    if org:
                        doc.Организация = org if hasattr(org, "Ref") else org
                        log(f"[create_wax_job_from_task] ✅ Установлена организация: {safe_str(org)}")
                except Exception as e:
                    log(f"[create_wax_job_from_task] ⚠ Не удалось установить организацию: {e}")

                try:
                    wh = getattr(order_obj, "Склад", None)
                    if wh:
                        doc.Склад = wh if hasattr(wh, "Ref") else wh
                        log(f"[create_wax_job_from_task] ✅ Установлен склад: {safe_str(wh)}")
                except Exception as e:
                    log(f"[create_wax_job_from_task] ⚠ Не удалось установить склад: {e}")

            # --- Табличная часть
            for row in task.Продукция:
                r = doc.ТоварыВыдано.Add()
                r.Номенклатура = row.Номенклатура
                r.Размер = row.Размер
                r.Количество = row.Количество
                r.ВариантИзготовления = row.ВариантИзготовления
                r.Проба = row.Проба
                r.ЦветМеталла = row.ЦветМеталла
                r.Вес = row.Вес

            doc.Проведен = True
            doc.Write()
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
            if hasattr(ref, "GetObject"):
                obj = ref.GetObject()
            else:
                obj = self.connection.GetObject(ref)
            log(f"[get_object_from_ref] ✅ Получен объект из ссылки")
            return obj
        except Exception as e:
            log(f"[get_object_from_ref] ❌ Ошибка получения объекта по ссылке: {e}")
            return None

    def get_object_property(self, obj, prop_name: str):
        """Возвращает значение свойства 1С-объекта."""
        try:
            target = obj
            if hasattr(target, "GetObject"):
                target = target.GetObject()
            return getattr(target, prop_name, None)
        except Exception as e:
            log(f"[get_object_property] Ошибка получения {prop_name}: {e}")
            return None
    # ------------------------------------------------------------------

    def create_wax_jobs_from_task(
        self,
        task_ref,
        master_3d: str,
        master_form: str,
        warehouse: str | None = None,
        norm_type: str = "Номенклатура",
    ) -> list[str]:
        """Создаёт два наряда из одного задания по артикулу."""
        mapping = {"3D печать": master_3d, "Пресс-форма": master_form}
        result: list[str] = []

        # Получаем объект и ссылку задания
        try:
            if isinstance(task_ref, str):
                task = self.connection.GetObject(task_ref)
            elif hasattr(task_ref, "Продукция"):
                task = task_ref
            elif hasattr(task_ref, "GetObject"):
                task = task_ref.GetObject()
            else:
                log("[create_wax_jobs_from_task] ❌ Неверный тип ссылки")
                return []
            task_ref_link = task.Ref
        except Exception as exc:
            log(f"[create_wax_jobs_from_task] ❌ Ошибка доступа к заданию: {exc}")
            return []

        # Получаем данные для шапки наряда
        org = getattr(task, "Организация", None)
        wh = getattr(task, "Склад", None)
        responsible = getattr(task, "Ответственный", None)

        # Попытка получить значения из связанного заказа, если их нет в задании
        order_ref = (
            getattr(task, "ЗаказВПроизводство", None)
            or getattr(task, "ДокументОснование", None)
        )
        if (org is None or wh is None) and order_ref and hasattr(order_ref, "GetObject"):
            try:
                order_obj = order_ref.GetObject()
                org = org or getattr(order_obj, "Организация", None)
                wh = wh or getattr(order_obj, "Склад", None)
                responsible = responsible or getattr(order_obj, "Ответственный", None)
            except Exception as e:
                log(
                    f"[create_wax_jobs_from_task] ⚠ Не удалось получить данные из заказа: {e}"
                )

        section = getattr(task, "ПроизводственныйУчасток", None)
        if warehouse:
            wh = self.get_ref_by_description("Склады", warehouse) or wh

        rows_by_method = {"3D печать": [], "Пресс-форма": []}
        for row in task.Продукция:
            art = safe_str(getattr(row.Номенклатура, "Артикул", "")).lower()
            method = "3D печать" if "д" in art or "d" in art else "Пресс-форма"
            rows_by_method[method].append(row)

        for method, rows in rows_by_method.items():
            if not rows:
                continue
            try:
                job = self.documents.НарядВосковыеИзделия.CreateDocument()

                # Заполнение шапки наряда
                job.Дата = datetime.now()
                job.ДокументОснование = task_ref_link
                job.ЗаданиеНаПроизводство = task_ref_link
                if org is not None:
                    try:
                        job.Организация = org
                    except Exception as exc:
                        log(
                            f"[create_wax_jobs_from_task] ⚠ Не удалось установить организацию: {exc}"
                        )
                if wh is not None:
                    try:
                        job.Склад = wh
                    except Exception as exc:
                        log(
                            f"[create_wax_jobs_from_task] ⚠ Не удалось установить склад: {exc}"
                        )
                if section:
                    job.ПроизводственныйУчасток = section
                if responsible:
                    job.Ответственный = responsible
                job.ТехОперация = self.get_ref("ТехОперации", method)
                job.Сотрудник = self.get_ref("ФизическиеЛица", mapping.get(method, ""))
                job.Комментарий = f"Создан автоматически для {method}"

                # Заполнение табличной части вручную
                for r in rows:
                    row = job.ТоварыВыдано.Add()
                    row.Номенклатура = r.Номенклатура
                    row.Количество = r.Количество
                    row.Размер = r.Размер
                    row.Проба = r.Проба
                    row.ЦветМеталла = r.ЦветМеталла
                    if hasattr(r, "ХарактеристикаВставок"):
                        row.ХарактеристикаВставок = r.ХарактеристикаВставок
                    if hasattr(r, "Вес"):
                        row.Вес = r.Вес

                # ---- Подстановка вида норматива для всех строк
                enum_norm = self.get_enum_by_description(
                    "ВидыНормативовНоменклатуры", norm_type
                )
                if enum_norm:
                    for row in job.ТоварыВыдано:
                        row.ВидНорматива = enum_norm


                job.Проведен = True
                job.Write()
                result.append(str(job.Номер))
                log(f"[create_wax_jobs_from_task] ✅ Создан наряд {method}: №{job.Номер}")
            except Exception as exc:
                log(f"[create_wax_jobs_from_task] ❌ Ошибка для {method}: {exc}")
        return result

    def _find_task_by_number(self, number: str):
        """Ищет задание по номеру через WaxBridge."""
        return self.wax_bridge._find_task_by_number(number)

    def post_task(self, number: str) -> bool:
        """Проводит задание через WaxBridge."""
        return self.wax_bridge.post_task(number)

    def undo_post_task(self, number: str) -> bool:
        """Снимает проведение задания через WaxBridge."""
        return self.wax_bridge.undo_post_task(number)

    def mark_task_for_deletion(self, number: str) -> bool:
        """Помечает задание на удаление через WaxBridge."""
        return self.wax_bridge.mark_task_for_deletion(number)

    def unmark_task_deletion(self, number: str) -> bool:
        """Снимает пометку удаления задания через WaxBridge."""
        return self.wax_bridge.unmark_task_deletion(number)

    def delete_task(self, number: str) -> bool:
        """Удаляет задание через WaxBridge."""
        return self.wax_bridge.delete_task(number)

    # -------------------------------------------------------------
    # Методы для работы с нарядами
    # -------------------------------------------------------------

    def _find_wax_job_by_number(self, number: str):
        """Ищет наряд по номеру через WaxBridge."""
        return self.wax_bridge._find_wax_job_by_number(number)

    def post_wax_job(self, number: str) -> bool:
        """Проводит наряд через WaxBridge."""
        return self.wax_bridge.post_wax_job(number)

    def undo_post_wax_job(self, number: str) -> bool:
        """Отменяет проведение наряда через WaxBridge."""
        return self.wax_bridge.undo_post_wax_job(number)

    def mark_wax_job_for_deletion(self, number: str) -> bool:
        """Помечает наряд на удаление через WaxBridge."""
        return self.wax_bridge.mark_wax_job_for_deletion(number)

    def unmark_wax_job_deletion(self, number: str) -> bool:
        """Снимает пометку удаления наряда через WaxBridge."""
        return self.wax_bridge.unmark_wax_job_deletion(number)

    def delete_wax_job(self, number: str) -> bool:
        """Удаляет наряд через WaxBridge."""
        return self.wax_bridge.delete_wax_job(number)
            
            