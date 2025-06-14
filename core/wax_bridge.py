# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Any
from .com_bridge import log, safe_str


class WaxBridge:
    """Часть COM-моста для работы с нарядами и заданиями."""

    def __init__(self, bridge: 'COM1CBridge'):
        self.bridge = bridge

    # -------------------------------------------------------------
    # Методы работы с заданиями и нарядами
    # -------------------------------------------------------------

    def _find_task_by_number(self, number: str):
        doc_manager = getattr(self.bridge.connection.Documents, "ЗаданиеНаПроизводство", None)
        if doc_manager is None:
            log("❌ Документ 'ЗаданиеНаПроизводство' не найден")
            return None
        selection = doc_manager.Select()
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Number).strip() == str(number).strip():
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
            if getattr(obj, "Проведен", False):
                obj.Проведен = False
                obj.Write()

            obj.DeletionMark = True
            obj.Write()
            obj.Delete()
            log(f"[Удаление] ✅ Задание №{number} удалено")
            return True
        except Exception as e:
            log(f"❌ Ошибка при удалении задания №{number}: {e}")
            return False

    def list_tasks(self) -> list[dict]:
        result = []
        doc_manager = getattr(self.bridge.connection.Documents, "ЗаданиеНаПроизводство", None)
        if doc_manager is None:
            return result
        selection = doc_manager.Select()
        while selection.Next():
            doc = selection.GetObject()
            result.append({
                "ref": str(doc.Ref),
                "num": str(doc.Number),
                "date": str(doc.Date.strftime("%d.%m.%Y")),
                "employee": self.bridge.safe(doc, "Ответственный"),
                "tech_op": self.bridge.safe(doc, "ТехОперация"),
                "section": self.bridge.safe(doc, "ПроизводственныйУчасток"),
                "posted": getattr(doc, "Проведен", False),
                "deleted": getattr(doc, "ПометкаУдаления", False),
            })
        return result

    def get_task_lines(self, doc_num: str) -> list[dict]:
        result = []
        doc_manager = getattr(self.bridge.connection.Documents, "ЗаданиеНаПроизводство", None)
        if not doc_manager:
            return result

        selection = doc_manager.Select()
        while selection.Next():
            doc = selection.GetObject()
            if str(doc.Number) == str(doc_num):
                for row in doc.Продукция:
                    result.append({
                        "nomen": self.bridge.safe(row, "Номенклатура"),
                        "article": safe_str(getattr(row.Номенклатура, "Артикул", "")),
                        "size": self.bridge.safe(row, "Размер"),
                        "sample": self.bridge.safe(row, "Проба"),
                        "color": self.bridge.safe(row, "ЦветМеталла"),
                        "qty": row.Количество,
                        "weight": row.Вес if hasattr(row, "Вес") else "",
                    })
                break
        return result

    def list_wax_jobs(self) -> list[dict]:
        result = []
        docs = self.bridge.connection.Documents.НарядВосковыеИзделия.Select()
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
        """Возвращает наряды, связанные с указанным заданием."""
        result: list = []

        if hasattr(task_ref, "Ref"):
            task_ref = task_ref.Ref

        task_ref = str(task_ref)

        docs = self.bridge.connection.Documents.НарядВосковыеИзделия.Select()
        if hasattr(docs, "Count"):
            log(f"[find_wax_jobs_by_task] всего найдено {docs.Count()} нарядов")

        while docs.Next():
            job = docs.GetObject()
            if not job or not hasattr(job, "ЗаданиеНаПроизводство"):
                continue
            try:
                if str(job.ЗаданиеНаПроизводство) == task_ref:
                    result.append(job.Ref)
            except Exception as e:
                log(f"[find_wax_jobs_by_task] ❌ Ошибка: {e}")

        log(
            f"[find_wax_jobs_by_task] найдено {len(result)} "
            f"нарядов для задания {task_ref}"
        )
        return result

    def close_wax_jobs(self, job_refs: list) -> list[str]:
        closed: list[str] = []
        for ref in job_refs:
            try:
                doc = self.bridge.get_object_from_ref(ref)
                try:
                    doc.ТоварыПринято.ЗаполнитьПоВыданному()
                except Exception as exc:
                    log(f"[close_wax_jobs] ⚠ Заполнение: {exc}")
                try:
                    doc.Закрыт = True
                except Exception:
                    pass
                doc.Провести()
                closed.append(str(doc.Номер))
                log(f"[close_wax_jobs] ✅ {doc.Номер}")
            except Exception as e:
                log(f"[close_wax_jobs] ❌ {e}")
        return closed

    def get_wax_job_lines(self, doc_num: str) -> list[dict]:
        result = []
        doc_manager = getattr(self.bridge.connection.Documents, "НарядВосковыеИзделия", None)
        if not doc_manager:
            return result

        selection = doc_manager.Select()
        while selection.Next():
            doc = selection.GetObject()
            if str(doc.Number) == str(doc_num):
                for row in doc.ТоварыВыдано:
                    result.append({
                        "norm": self.bridge.safe(row, "ВидНорматива"),
                        "nomen": self.bridge.safe(row, "Номенклатура"),
                        "size": self.bridge.safe(row, "Размер"),
                        "sample": self.bridge.safe(row, "Проба"),
                        "color": self.bridge.safe(row, "ЦветМеталла"),
                        "qty": row.Количество,
                        "weight": round(float(row.Вес), self.bridge.config.WEIGHT_DECIMALS) if hasattr(row, "Вес") else "",
                    })
                break
        return result

    def get_wax_job_rows(self, num: str) -> list[dict]:
        doc = self.bridge._find_doc("НарядВосковыеИзделия", num)
        rows = []
        for r in doc.Выдано:
            rows.append({
                "Номенклатура": self.bridge.safe(r, "Номенклатура"),
                "Размер": self.bridge.safe(r, "Размер"),
                "Проба": self.bridge.safe(r, "Проба"),
                "Цвет": self.bridge.safe(r, "ЦветМеталла"),
                "Количество": r.Количество,
                "Вес": round(float(r.Вес), self.bridge.config.WEIGHT_DECIMALS),
                "Партия": self.bridge.safe(r, "Партия"),
                "Номер ёлки": r.НомерЕлки if hasattr(r, "НомерЕлки") else "",
                "Состав набора": r.СоставНабора if hasattr(r, "СоставНабора") else "",
            })
        return rows

    def create_wax_job_from_task(self, task_number: str) -> str:
        task = self._find_task_by_number(task_number)
        if not task:
            log(f"❌ Задание №{task_number} не найдено")
            return ""

        try:
            doc = self.bridge.connection.Documents.НарядВосковыеИзделия.CreateDocument()
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

            order_ref = getattr(task, "ЗаказВПроизводство", None) or getattr(task, "ДокументОснование", None)
            order_obj = None
            if order_ref and hasattr(order_ref, "GetObject"):
                try:
                    order_obj = order_ref.GetObject()
                    log("[create_wax_job_from_task] ✅ Получен заказ-основание")
                except Exception as exc:
                    log(f"[create_wax_job_from_task] ⚠ Ошибка получения заказа: {exc}")

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

    def create_wax_jobs_from_task(
        self,
        task_ref,
        master_3d: str,
        master_form: str,
        warehouse: str | None = None,
        norm_type: str = "Номенклатура",
    ) -> list[str]:
        mapping = {"3D печать": master_3d, "Пресс-форма": master_form}
        result: list[str] = []

        try:
            if isinstance(task_ref, str):
                task = self.bridge.connection.GetObject(task_ref)
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

        org = getattr(task, "Организация", None)
        wh = getattr(task, "Склад", None)
        responsible = getattr(task, "Ответственный", None)

        order_ref = getattr(task, "ЗаказВПроизводство", None) or getattr(task, "ДокументОснование", None)
        if (org is None or wh is None) and order_ref and hasattr(order_ref, "GetObject"):
            try:
                order_obj = order_ref.GetObject()
                org = org or getattr(order_obj, "Организация", None)
                wh = wh or getattr(order_obj, "Склад", None)
                responsible = responsible or getattr(order_obj, "Ответственный", None)
            except Exception as e:
                log(f"[create_wax_jobs_from_task] ⚠ Не удалось получить данные из заказа: {e}")

        section = getattr(task, "ПроизводственныйУчасток", None)
        if warehouse:
            wh = self.bridge.get_ref_by_description("Склады", warehouse) or wh

        rows_by_method = {"3D печать": [], "Пресс-форма": []}
        for row in task.Продукция:
            art = safe_str(getattr(row.Номенклатура, "Артикул", "")).lower()
            method = "3D печать" if "д" in art or "d" in art else "Пресс-форма"
            rows_by_method[method].append(row)

        for method, rows in rows_by_method.items():
            if not rows:
                continue
            try:
                job = self.bridge.connection.Documents.НарядВосковыеИзделия.CreateDocument()

                job.Дата = datetime.now()
                job.ДокументОснование = task_ref_link
                job.ЗаданиеНаПроизводство = task_ref_link
                if org is not None:
                    try:
                        job.Организация = org
                    except Exception as exc:
                        log(f"[create_wax_jobs_from_task] ⚠ Не удалось установить организацию: {exc}")
                if wh is not None:
                    try:
                        job.Склад = wh
                    except Exception as exc:
                        log(f"[create_wax_jobs_from_task] ⚠ Не удалось установить склад: {exc}")
                if section:
                    job.ПроизводственныйУчасток = section
                if responsible:
                    job.Ответственный = responsible
                job.ТехОперация = self.bridge.get_ref("ТехОперации", method)
                job.Сотрудник = self.bridge.get_ref("ФизическиеЛица", mapping.get(method, ""))
                job.Комментарий = f"Создан автоматически для {method}"

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

                enum_norm = self.bridge.get_enum_by_description(
                    "ВидыНормативовНоменклатуры", norm_type
                )
                if enum_norm:
                    for row in job.ТоварыВыдано:
                        row.ВидНорматива = enum_norm

                for row in job.ТоварыВыдано:
                    nomenclature = getattr(row, "Номенклатура", None)
                    if nomenclature is not None:
                        type_enum = self.bridge.get_object_property(nomenclature, "ТипНоменклатуры")
                        enum = self.bridge.get_enum_by_description(
                            "ВидыНормативовНоменклатуры",
                            "Номенклатура" if safe_str(type_enum) == "Продукция" else "Комплектующее",
                        )
                        if enum:
                            row.ВидНорматива = enum

                job.Write()
                result.append(str(job.Номер))
                log(f"[create_wax_jobs_from_task] ✅ Создан наряд {method}: №{job.Номер}")
            except Exception as exc:
                log(f"[create_wax_jobs_from_task] ❌ Ошибка для {method}: {exc}")
        return result
