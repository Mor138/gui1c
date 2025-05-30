# JewelryMES/pages/form_builder_page.py
from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QComboBox, QPushButton, QLabel, QSizePolicy
)
from form_builder import open_form_builder
from core import models   # ← где у вас объявлены dataclass-модели

def _iter_models():
    """Берём все классы из core.models, помеченные атрибутом __formable__ = True
       (чтобы в выпадашку не попали «служебные» классы)."""
    for name, cls in models.__dict__.items():
        if getattr(cls, "__formable__", False):
            yield name, cls

class FormBuilderPage(QWidget):
    """Небольшая панель: выбрали модель → нажали кнопку → открылся конструктор."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._models = dict(_iter_models())

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Выберите объект 1С / dataclass"))

        self.cmb = QComboBox()
        self.cmb.addItems(self._models.keys())
        self.cmb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(self.cmb)

        btn = QPushButton("Открыть конструктор формы")
        btn.clicked.connect(self._open_builder)          # type: ignore[arg-type]
        layout.addWidget(btn)

        layout.addStretch()

    # ---------- slots ----------
    def _open_builder(self):
        model_cls = self._models[self.cmb.currentText()]
        open_form_builder(model_cls)
