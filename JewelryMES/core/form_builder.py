"""
Runtime form builder (конструктор форм) для проекта JewelryMES.

Позволяет:
- интроспективно считывать dataclass‑модели (или SQLAlchemy/пользовательские) и отображать список доступных полей;
- перетаскиванием (double‑click) добавлять поля в форму‑превью;
- изменять подписи и порядок (доступно через свойства QListWidgetItem – оставил как TODO);
- сохранять описание формы в лёгком JSON‑формате *.jfrm;
- в рантайме загружать сохранённую форму и получать готовый QWidget для вставки в окно приложения;
- подключать конструктор «на горячую» через функцию open_form_builder(model_cls).

Зависимости: PySide6 (pip install PySide6)
"""

from __future__ import annotations

import dataclasses
import inspect
import json
from typing import Any, Dict, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QToolBar,
    QWidget,
)


class FieldMeta:
    """Крошечный объект‑обёртка над атрибутом модели."""

    def __init__(self, name: str, ftype: str, title: str | None = None):
        self.name = name
        self.ftype = ftype
        self.title = title or name.capitalize()

    @classmethod
    def from_dataclass(cls, field: dataclasses.Field):  # type: ignore[arg-type]
        return cls(
            field.name,
            getattr(field.type, "__name__", str(field.type)),
            field.metadata.get("title") if hasattr(field, "metadata") else None,
        )

    def to_dict(self) -> Dict[str, Any]:
        return {"name": self.name, "ftype": self.ftype, "title": self.title}


class FormBlueprint(dict):
    """Serialisable описание формы (версия 1)."""

    version = 1

    def to_json(self, path: str):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self, f, ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, path: str) -> "FormBlueprint":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls(**data)


class FormBuilderWindow(QMainWindow):
    def __init__(self, model_cls):
        super().__init__()
        self.setWindowTitle(f"Конструктор формы — {model_cls.__name__}")
        self.model_cls = model_cls
        self.blueprint: FormBlueprint = FormBlueprint(fields=[])
        self._build_ui()

    # ----------------------------- UI helpers ----------------------------- #
    def _build_ui(self):
        splitter = QSplitter(Qt.Horizontal, self)
        self.setCentralWidget(splitter)

        # --- Панель доступных полей --- #
        self.list_fields = QListWidget()
        for fld in dataclasses.fields(self.model_cls):  # type: ignore[arg-type]
            meta = FieldMeta.from_dataclass(fld)
            item = QListWidgetItem(f"{meta.name} ({meta.ftype})")
            item.setData(Qt.UserRole, meta)
            self.list_fields.addItem(item)
        splitter.addWidget(self.list_fields)

        # --- Превью формочки --- #
        preview_container = QWidget()
        self.preview_layout = QFormLayout(preview_container)
        splitter.addWidget(preview_container)

        # --- Сигналы --- #
        self.list_fields.itemDoubleClicked.connect(self._add_field)  # type: ignore[arg-type]

        # --- Toolbar --- #
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self._save)
        validate_btn = QPushButton("Проверить")
        validate_btn.clicked.connect(self._validate)
        self.addToolBar(Qt.TopToolBarArea, self._build_toolbar([save_btn, validate_btn]))

    def _build_toolbar(self, widgets: List[QPushButton]) -> QToolBar:  # noqa: D401
        bar = QToolBar("Tools", self)
        for w in widgets:
            bar.addWidget(w)
        return bar

    # ----------------------------- callbacks ----------------------------- #
    def _add_field(self, item: QListWidgetItem):
        meta: FieldMeta = item.data(Qt.UserRole)
        label = QLabel(meta.title)
        editor: QWidget
        if meta.ftype in {"str", "uuid", "date", "datetime", "Decimal"}:
            editor = QLineEdit()
        else:
            editor = QLabel(f"<{meta.ftype}>")
            editor.setStyleSheet("color: grey; font-style: italic;")
        self.preview_layout.addRow(label, editor)
        self.blueprint.setdefault("fields", []).append(meta.to_dict())

    def _save(self):
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить форму", "", "Jewelry Form (*.jfrm)")
        if not path:
            return
        self.blueprint.to_json(path)
        QMessageBox.information(self, "Сохранено", f"Форма сохранена в {path}")

    def _validate(self):
        # simple duplicate check
        names = [f["name"] for f in self.blueprint.get("fields", [])]
        dups = {n for n in names if names.count(n) > 1}
        if dups:
            QMessageBox.warning(self, "Проблемы", f"Дублирующиеся поля: {', '.join(dups)}")
        else:
            QMessageBox.information(self, "OK", "Всё выглядит хорошо! 😊")


# ===================== Public API ===================== #

def open_form_builder(model_cls):
    """Запускает окно конструктора формы. Можно вызвать *прямо на горячую* из любого места кода."""
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    dlg = FormBuilderWindow(model_cls)
    dlg.resize(900, 600)
    dlg.show()
    if app.instance() is None:  # pragma: no cover — никогда не бывает
        app.exec()


def load_form(json_path: str) -> QWidget:
    """Создаёт готовый QWidget на основе сохранённого *.jfrm\n    Используйте как сетевой элемент в существующих окнах приложения."""

    bp = FormBlueprint.from_json(json_path)
    container = QWidget()
    layout = QFormLayout(container)
    for fld in bp["fields"]:
        lbl = QLabel(fld["title"])
        if fld["ftype"] in {"str", "uuid", "date", "datetime", "Decimal"}:
            ed = QLineEdit()
        else:
            ed = QLabel(f"<{fld['ftype']}>")
            ed.setStyleSheet("color: grey; font-style: italic;")
        layout.addRow(lbl, ed)
    return container


# ===================== Example usage ===================== #

if __name__ == "__main__":  # запускаем демо, если файл вызвать напрямую
    import datetime

    @dataclasses.dataclass
    class Order:
        id: int
        number: str
        customer: str
        date: datetime.date
        amount: float

    open_form_builder(Order)
