# -*- coding: utf-8 -*-
from __future__ import annotations
import os
import tempfile
from typing import Any
from win32com.client import VARIANT
from pythoncom import VT_BOOL
from .com_bridge import safe_str, log


class OrdersBridge:
    """Часть COM-моста, относящаяся к документу 'ЗаказВПроизводство'."""

    def __init__(self, bridge: 'COM1CBridge'):
        self.bridge = bridge

    def print_order_preview_pdf(self, number: str, date: str | None = None) -> bool:
        obj = self.bridge._find_document_by_number("ЗаказВПроизводство", number, date)
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
                os.startfile(pdf_path)
                return True
            log("❌ Не удалось сохранить PDF")
            return False
        except Exception as e:
            log(f"❌ Ошибка при формировании PDF: {e}")
            return False

    def get_last_order_number(self) -> str:
        doc = getattr(self.bridge.documents, "ЗаказВПроизводство", None)
        if not doc:
            return "00ЮП-000000"
        selection = doc.Select()
        number = "00ЮП-000000"
        while selection.Next():
            try:
                obj = selection.GetObject()
                number = str(obj.Number)
            except Exception:
                continue
        return number

    def get_next_order_number(self) -> str:
        last = self.get_last_order_number()
        try:
            prefix, num = last.split("-")
            next_num = int(num) + 1
            return f"{prefix}-{next_num:06d}"
        except Exception:
            return "00ЮП-000001"

    # -------------------------------------------------------------
    # Методы работы с документом "ЗаказВПроизводство"
    # -------------------------------------------------------------

    def undo_posting(self, number: str, date: str | None = None) -> bool:
        obj = self.bridge._find_document_by_number("ЗаказВПроизводство", number, date)
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

    def delete_order_by_number(self, number: str, date: str | None = None) -> bool:
        obj = self.bridge._find_document_by_number("ЗаказВПроизводство", number, date)
        if not obj:
            log(f"[Удаление] Заказ №{number} не найден")
            return False
        try:
            if getattr(obj, "Проведен", False):
                log("⚠ Документ проведён, снимаем проведение...")
                self.undo_posting(number)
            obj.Delete()
            log("🗑 Документ удалён полностью")
            return True
        except Exception as e:
            log(f"❌ Ошибка при удалении: {e}")
            return False

    def post_order(self, number: str, date: str | None = None) -> bool:
        obj = self.bridge._find_document_by_number("ЗаказВПроизводство", number, date)
        if not obj:
            log(f"[Проведение] Заказ №{number} не найден")
            return False
        try:
            obj.Проведен = True
            obj.Write()
            if getattr(obj, "Проведен", False):
                log(f"[Проведение] Заказ №{number} успешно проведён через флаг Проведен")
                return True
            log(f"[Проведение] Не удалось провести заказ №{number}")
            return False
        except Exception as e:
            log(f"[Проведение] Ошибка при установке Проведен: {e}")
            return False

    def mark_order_for_deletion(self, number: str, date: str | None = None) -> bool:
        obj = self.bridge._find_document_by_number("ЗаказВПроизводство", number, date)
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
            log("❌ Не удалось установить пометку на удаление")
            return False
        except Exception as e:
            log(f"❌ Ошибка при установке пометки: {e}")
            return False

    def unmark_order_deletion(self, number: str, date: str | None = None) -> bool:
        obj = self.bridge._find_document_by_number("ЗаказВПроизводство", number, date)
        if not obj:
            log(f"[Снятие пометки] Документ №{number} не найден")
            return False
        try:
            obj.DeletionMark = VARIANT(VT_BOOL, False)
            obj.Write()
            if not getattr(obj, "DeletionMark", True):
                log(f"✅ Пометка на удаление снята с документа №{number}")
                return True
            log("❌ Не удалось снять пометку")
            return False
        except Exception as e:
            log(f"❌ Ошибка при снятии пометки: {e}")
            return False

    def update_order(self, number: str, fields: dict, items: list, date: str | None = None) -> bool:
        obj = self.bridge._find_document_by_number("ЗаказВПроизводство", number, date)
        if not obj:
            log(f"[Обновление] Заказ №{number} не найден")
            return False

        for k, v in fields.items():
            try:
                if k == "ВидСтатусПродукции":
                    ref = self.bridge.get_ref("ВидыСтатусыПродукции", v)
                    if ref:
                        setattr(obj, k, ref)
                    continue
                if k in ["Организация", "Контрагент", "ДоговорКонтрагента", "Ответственный", "Склад"]:
                    ref = self.bridge.get_ref(k + "ы" if not k.endswith("т") else k + "а", v)
                    if ref:
                        setattr(obj, k, ref)
                    continue
                setattr(obj, k, v)
            except Exception as e:
                log(f"[Обновление] Ошибка установки поля {k}: {e}")

        while len(obj.Товары) > 0:
            obj.Товары.Delete(0)

        for row in items:
            try:
                new_row = obj.Товары.Add()
                new_row.Номенклатура = self.bridge.get_ref("Номенклатура", row.get("Номенклатура"))
                variant = row.get("ВариантИзготовления")
                if variant and variant != "—":
                    new_row.ВариантИзготовления = self.bridge.get_ref("ВариантыИзготовленияНоменклатуры", variant)
                size_val = row.get("Размер", 0)
                new_row.Размер = self.bridge.get_size_ref(size_val)
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

    def create_order(self, fields: dict, items: list) -> str:
        doc = self.bridge.documents.ЗаказВПроизводство.CreateDocument()
        catalog_fields_map = {
            "Организация": "Организации",
            "Контрагент": "Контрагенты",
            "ДоговорКонтрагента": "ДоговорыКонтрагентов",
            "Ответственный": "Пользователи",
            "Склад": "Склады",
        }

        log("Создание заказа. Поля:")
        for k, v in fields.items():
            try:
                log(f"  -> {k} = {v}")

                if k == "ВидСтатусПродукции":
                    reverse_map = {val: key for key, val in PRODUCTION_STATUS_MAP.items()}
                    if v in reverse_map:
                        v = reverse_map[v]

                    ref = self.bridge.get_ref("ВидыСтатусыПродукции", v)
                    log(f"[ref] {k}: {v} => {ref}")
                    if ref:
                        setattr(doc, k, ref)
                        log(f"    Установлено: {k} (Ref: {ref})")
                    else:
                        log(f"    ❌ {k} '{v}' не найден.")
                    continue

                if k in catalog_fields_map:
                    ref = self.bridge.get_ref(catalog_fields_map[k], v)
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
                new_row.Номенклатура = self.bridge.get_ref("Номенклатура", row.get("Номенклатура"))

                variant = row.get("ВариантИзготовления")
                if variant and variant != "—":
                    new_row.ВариантИзготовления = self.bridge.get_ref("ВариантыИзготовленияНоменклатуры", variant)

                size_val = row.get("Размер", 0)
                size_ref = self.bridge.get_size_ref(size_val)
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

    def list_orders(self) -> list[dict]:
        result = []
        orders = self.bridge.connection.Documents.ЗаказВПроизводство.Select()
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
                "Ref": doc.Ref,
                "num": doc.Номер,
                "date": str(doc.Дата),
                "org": safe_str(doc.Организация),
                "contragent": safe_str(doc.Контрагент),
                "contract": safe_str(doc.ДоговорКонтрагента),
                "comment": safe_str(doc.Комментарий),
                "prod_status": self.bridge.to_string(doc.ВидСтатусПродукции),
                "posted": doc.Проведен,
                "deleted": doc.ПометкаУдаления,
                "qty": sum([r["qty"] for r in rows]),
                "weight": sum([r["w"] for r in rows]),
                "rows": rows,
            })
        return result

    def get_order_lines(self, doc_number: str, date: str | None = None) -> list[dict]:
        doc = self.bridge._find_document_by_number("ЗаказВПроизводство", doc_number, date)
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
                "weight": float(getattr(r, "Вес", 0)),
                "article": safe_str(getattr(r.Номенклатура, "Артикул", "")),
            })
        return rows

