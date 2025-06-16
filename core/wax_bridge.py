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

    # -------------------------------------------------------------
    # Дополнительные операции с нарядами
    # -------------------------------------------------------------

    def _find_wax_job_by_number(self, number: str):
        doc_manager = getattr(self.bridge.connection.Documents, "НарядВосковыеИзделия", None)
        if doc_manager is None:
            log("❌ Документ 'НарядВосковыеИзделия' не найден")
            return None
        selection = doc_manager.Select()
        while selection.Next():
            obj = selection.GetObject()
            if str(obj.Number).strip() == str(number).strip():
                return obj
        return None

    def post_wax_job(self, number: str) -> bool:
        obj = self._find_wax_job_by_number(number)
        if not obj:
            log(f"[Проведение] ❌ Наряд №{number} не найден")
            return False
        try:
            obj.Проведен = True
            obj.Write()
            log(f"[Проведение] ✅ Наряд №{number} проведён")
            return True
        except Exception as e:
            log(f"❌ Ошибка при проведении наряда №{number}: {e}")
            return False

    def undo_post_wax_job(self, number: str) -> bool:
        obj = self._find_wax_job_by_number(number)
        if not obj:
            log(f"[Снятие проведения] ❌ Наряд №{number} не найден")
            return False
        try:
            obj.Проведен = False
            obj.Write()
            log(f"[Снятие проведения] ✅ Наряд №{number} отменён")
            return True
        except Exception as e:
            log(f"❌ Ошибка при отмене проведения наряда №{number}: {e}")
            return False

    def mark_wax_job_for_deletion(self, number: str) -> bool:
        obj = self._find_wax_job_by_number(number)
        if not obj:
            log(f"[Пометка удаления] ❌ Наряд №{number} не найден")
            return False
        try:
            obj.DeletionMark = True
            obj.Write()
            log(f"[Пометка удаления] ✅ Наряд №{number} помечен на удаление")
            return True
        except Exception as e:
            log(f"❌ Ошибка при пометке наряда №{number}: {e}")
            return False

    def unmark_wax_job_deletion(self, number: str) -> bool:
        obj = self._find_wax_job_by_number(number)
        if not obj:
            log(f"[Снятие пометки] ❌ Наряд №{number} не найден")
            return False
        try:
            obj.DeletionMark = False
            obj.Write()
            log(f"[Снятие пометки] ✅ Наряд №{number} восстановлен")
            return True
        except Exception as e:
            log(f"❌ Ошибка при снятии пометки наряда №{number}: {e}")
            return False

    def delete_wax_job(self, number: str) -> bool:
        obj = self._find_wax_job_by_number(number)
        if not obj:
            log(f"[Удаление] ❌ Наряд №{number} не найден")
            return False
        try:
            if getattr(obj, "Проведен", False):
                obj.Проведен = False
                obj.Write()
            obj.DeletionMark = True
            obj.Write()
            obj.Delete()
            log(f"[Удаление] ✅ Наряд №{number} удалён")
            return True
        except Exception as e:
            log(f"❌ Ошибка при удалении наряда №{number}: {e}")
            return False

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
        """Возвращает наряды, связанные с указанным заданием."""
        result: list = []

        task_ref_str = self.bridge.to_string(task_ref)
        jobs = self.bridge.list_documents("НарядВосковыеИзделия")

        for job in jobs:
            try:
                job_task_ref = getattr(job, "ЗаданиеНаПроизводство", None)
                if job_task_ref and self.bridge.to_string(job_task_ref) == task_ref_str:
                    result.append(job.Ref)
            except Exception as e:
                log(f"[find_wax_jobs_by_task] ❌ Ошибка: {e}")

        log(f"[find_wax_jobs_by_task] ✅ найдено {len(result)} нарядов для задания {task_ref_str}")
        return result

    def close_wax_jobs(self, job_refs: list) -> list[str]:
        from datetime import datetime
        closed: list[str] = []

        for ref in job_refs:
            try:
                doc = self.bridge.get_object_from_ref(ref)
                if not doc:
                    log("[close_wax_jobs] ❌ Не удалось получить документ по ссылке")
                    continue

                issued_table = getattr(doc, "ТоварыВыдано", None)
                accepted_table = getattr(doc, "ТоварыПринято", None)
                if not accepted_table:
                    log(f"[close_wax_jobs] ⚠ Не найдена табличная часть 'Принято' для {doc.Номер}")
                    continue

                # Попробуем штатно заполнить
                filled = False
                try:
                    accepted_table.Заполнить()
                    accepted_table.ЗаполнитьПоВыданному()
                    filled = True
                    log(f"[close_wax_jobs] ✅ Табличная часть Принято заполнена встроенно")
                except Exception as exc:
                    log(f"[close_wax_jobs] ⚠ Встроенное заполнение не удалось: {exc}")

                # Если не удалось — заполним вручную
                if not filled and issued_table:
                    accepted_table.Clear()
                    enum_norm = self.bridge.get_enum_by_description("ВидыНормативовНоменклатуры", "Номенклатура")
                    for r in issued_table:
                        if not getattr(r, "Номенклатура", None):
                            continue
                        if getattr(r, "Количество", 0) == 0:
                            continue

                        new_row = accepted_table.Add()
                        for attr in ("Номенклатура", "Размер", "Проба", "ЦветМеталла", "Характеристика"):
                            if hasattr(r, attr) and hasattr(new_row, attr):
                                setattr(new_row, attr, getattr(r, attr))

                        if hasattr(new_row, "ДатаПринятия"):
                            new_row.ДатаПринятия = datetime.now()

                        if hasattr(r, "Количество") and hasattr(new_row, "Количество"):
                            new_row.Количество = r.Количество

                        if hasattr(r, "Вес") and hasattr(new_row, "Вес"):
                            вес = getattr(r, "Вес", None)
                            if вес is not None and вес != 0:
                                new_row.Вес = вес

                        if enum_norm and hasattr(new_row, "ВидНорматива"):
                            new_row.ВидНорматива = enum_norm

                # Установим признак "Закрыт"
                try:
                    doc.Закрыт = True
                except Exception:
                    pass

                # Установим организацию, склад, ответственного, если не заданы
                if not getattr(doc, "Организация", None):
                    orgs = self.bridge.list_catalog_items("Организации")
                    if orgs:
                        doc.Организация = orgs[0]["Ref"]
                        log(f"[close_wax_jobs] ✅ Установлена организация: {orgs[0]['Description']}")

                if not getattr(doc, "Склад", None):
                    whs = self.bridge.list_catalog_items("Склады")
                    if whs:
                        doc.Склад = whs[0]["Ref"]
                        log(f"[close_wax_jobs] ✅ Установлен склад: {whs[0]['Description']}")

                if not getattr(doc, "Ответственный", None):
                    users = self.bridge.list_catalog_items("Пользователи")
                    if users:
                        doc.Ответственный = users[0]["Ref"]
                        log(f"[close_wax_jobs] ✅ Установлен ответственный: {users[0]['Description']}")

                # Запись и проведение
                doc.Write()
                log(f"[close_wax_jobs] ✅ Записан документ {doc.Номер}")
                doc.Провести()
                closed.append(str(doc.Номер))
                log(f"[close_wax_jobs] ✅ Проведён документ {doc.Номер}")

            except Exception as e:
                log(f"[close_wax_jobs] ❌ Ошибка при закрытии наряда: {e}")

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

            org = getattr(task, "Организация", None)
            wh = getattr(task, "Склад", None)
            if (org is None or wh is None) and order_obj:
                try:
                    org = org or getattr(order_obj, "Организация", None)
                    wh = wh or getattr(order_obj, "Склад", None)
                except Exception as exc:
                    log(f"[create_wax_job_from_task] ⚠ Ошибка получения данных заказа: {exc}")

            if org is not None:
                try:
                    doc.Организация = org if hasattr(org, "Ref") else org
                    log(f"[create_wax_job_from_task] ✅ Установлена организация: {safe_str(org)}")
                except Exception as e:
                    log(f"[create_wax_job_from_task] ⚠ Не удалось установить организацию: {e}")

            if wh is not None:
                try:
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

    def create_wax_jobs_from_task(self, task_ref, methods_map: dict[str, str], user_name: str) -> list:
        result = []
        task_obj = self.bridge.get_object_from_ref(task_ref)
        if not task_obj:
            raise Exception("Не удалось получить объект задания из ссылки")

        from datetime import datetime
        print(f"[DEBUG] loaded wax_bridge from: {__file__}")

        orders = self.bridge.find_documents_by_attribute("ЗаказВПроизводство", "Задание", task_ref)
        order_ref = orders[0] if orders else None
        if not order_ref:
            log("[create_wax_jobs_from_task] ⚠ Не найден заказ-основание")
        else:
            log("[create_wax_jobs_from_task] ✅ Получен заказ-основание")

        rows = getattr(task_obj, "Товары", [])
        if not rows:
            raise Exception("В задании нет строк")

        methods = {}
        for row in rows:
            method = str(getattr(row, "МетодИзготовления", None))
            if method not in methods:
                methods[method] = []
            methods[method].append(row)

        for method, method_rows in methods.items():
            doc = self.bridge.create_document_object("НарядВосковыеИзделия")
            doc.Дата = datetime.now()
            doc.Задание = task_ref
            if order_ref:
                doc.Заказ = order_ref

            # Установим метод изготовления
            if hasattr(doc, "МетодИзготовления"):
                method_enum = self.bridge.get_enum_by_ref("МетодыИзготовления", methods_map[method])
                doc.МетодИзготовления = method_enum

            # Установим организацию
            orgs = self.bridge.list_catalog_items("Организации")
            if orgs:
                doc.Организация = orgs[0]["Ref"]
                log(f"[create_wax_jobs_from_task] ✅ Установлена организация: {orgs[0]['Description']}")

            # Установим склад
            whs = self.bridge.list_catalog_items("Склады")
            if whs:
                doc.Склад = whs[0]["Ref"]
                log(f"[create_wax_jobs_from_task] ✅ Установлен склад: {whs[0]['Description']}")

            # Установим ответственного
            users = self.bridge.list_catalog_items("Пользователи")
            found_user = next((u for u in users if u["Description"] == user_name), None)
            if found_user:
                doc.Ответственный = found_user["Ref"]
                log(f"[create_wax_jobs_from_task] ✅ Ответственный: {found_user['Description']}")
            elif users:
                doc.Ответственный = users[0]["Ref"]
                log(f"[create_wax_jobs_from_task] ⚠ Установлен дефолтный ответственный: {users[0]['Description']}")

            tab = getattr(doc, "Товары", None)
            if not tab:
                raise Exception("Нет табличной части Товары в наряде")

            for r in method_rows:
                if not getattr(r, "Номенклатура", None):
                    continue

                new_row = tab.Add()
                for attr in ("Номенклатура", "Размер", "Проба", "ЦветМеталла", "Характеристика"):
                    if hasattr(r, attr) and hasattr(new_row, attr):
                        setattr(new_row, attr, getattr(r, attr))

                if hasattr(new_row, "Количество"):
                    new_row.Количество = r.Количество

                if hasattr(new_row, "Вес"):
                    new_row.Вес = r.Вес

            doc.Write()
            doc.Провести()
            result.append(doc.Ref)

            log(f"[create_wax_jobs_from_task] ✅ Создан наряд {method}: №{doc.Номер}")

        return result
