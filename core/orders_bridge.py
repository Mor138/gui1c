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
