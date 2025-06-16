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
        """# DEPRECATED: используйте bridge._find_task_by_number"""
        return self.bridge._find_task_by_number(number)

    def post_task(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.post_task"""
        return self.bridge.post_task(number)

    def undo_post_task(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.undo_post_task"""
        return self.bridge.undo_post_task(number)

    def mark_task_for_deletion(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.mark_task_for_deletion"""
        return self.bridge.mark_task_for_deletion(number)

    def unmark_task_deletion(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.unmark_task_deletion"""
        return self.bridge.unmark_task_deletion(number)

    def delete_task(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.delete_task"""
        return self.bridge.delete_task(number)

    def list_tasks(self) -> list[dict]:
        """# DEPRECATED: используйте bridge.list_tasks"""
        return self.bridge.list_tasks()

    # -------------------------------------------------------------
    # Дополнительные операции с нарядами
    # -------------------------------------------------------------

    def _find_wax_job_by_number(self, number: str):
        """# DEPRECATED: используйте bridge._find_wax_job_by_number"""
        return self.bridge._find_wax_job_by_number(number)

    def post_wax_job(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.post_wax_job"""
        return self.bridge.post_wax_job(number)

    def undo_post_wax_job(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.undo_post_wax_job"""
        return self.bridge.undo_post_wax_job(number)

    def mark_wax_job_for_deletion(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.mark_wax_job_for_deletion"""
        return self.bridge.mark_wax_job_for_deletion(number)

    def unmark_wax_job_deletion(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.unmark_wax_job_deletion"""
        return self.bridge.unmark_wax_job_deletion(number)

    def delete_wax_job(self, number: str) -> bool:
        """# DEPRECATED: используйте bridge.delete_wax_job"""
        return self.bridge.delete_wax_job(number)

    def get_task_lines(self, doc_num: str) -> list[dict]:
        """# DEPRECATED: используйте bridge.get_task_lines"""
        return self.bridge.get_task_lines(doc_num)

    def list_wax_jobs(self) -> list[dict]:
        """# DEPRECATED: используйте bridge.list_wax_jobs"""
        return self.bridge.list_wax_jobs()

    def find_wax_jobs_by_task(self, task_ref) -> list:
        """# DEPRECATED: используйте bridge.find_wax_jobs_by_task"""
        return self.bridge.find_wax_jobs_by_task(task_ref)

    def close_wax_jobs(self, job_refs: list) -> list[str]:
        """# DEPRECATED: используйте bridge.close_wax_jobs"""
        return self.bridge.close_wax_jobs(job_refs)

    def get_wax_job_lines(self, doc_num: str) -> list[dict]:
        """# DEPRECATED: используйте bridge.get_wax_job_lines"""
        return self.bridge.get_wax_job_lines(doc_num)

    def get_wax_job_rows(self, num: str) -> list[dict]:
        """# DEPRECATED: используйте bridge.get_wax_job_rows"""
        return self.bridge.get_wax_job_rows(num)

    def create_wax_job_from_task(self, task_number: str) -> str:
        """# DEPRECATED: используйте bridge.create_wax_job_from_task"""
        return self.bridge.create_wax_job_from_task(task_number)

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
