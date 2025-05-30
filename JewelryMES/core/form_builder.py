# JewelryMES/core/form_builder.py

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget,
    QTableWidgetItem, QHeaderView, QPushButton, QComboBox
)
import json
from pathlib import Path

SCHEMA_DIR = Path("form_schemas")
SCHEMA_DIR.mkdir(exist_ok=True)
_form_builder_window = None

FIELD_TYPES = ["Строка", "Число", "Дата", "Булево"]


def open_form_builder(tab_name: str):
    global _form_builder_window
    _form_builder_window = FormEditorWindow(tab_name)
    _form_builder_window.show()


class FormEditorWindow(QDialog):
    def __init__(self, tab_name):
        super().__init__()
        self.tab_name = tab_name
        self.setWindowTitle(f"Форма вкладки: {tab_name}")
        self.resize(700, 460)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"🧩 Настройка полей для вкладки: {tab_name}"))

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Поле", "Тип данных"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        buttons = QHBoxLayout()
        self.btn_add = QPushButton("➕ Добавить")
        self.btn_del = QPushButton("🗑️ Удалить")
        self.btn_save = QPushButton("💾 Сохранить")
        self.btn_close = QPushButton("Закрыть")

        self.btn_add.clicked.connect(self._add_row)
        self.btn_del.clicked.connect(self._del_row)
        self.btn_save.clicked.connect(self._save)
        self.btn_close.clicked.connect(self.close)

        buttons.addWidget(self.btn_add)
        buttons.addWidget(self.btn_del)
        buttons.addStretch()
        buttons.addWidget(self.btn_save)
        buttons.addWidget(self.btn_close)
        layout.addLayout(buttons)

        self._load()

    def _load(self):
        self.table.setRowCount(0)
        path = SCHEMA_DIR / f"{self.tab_name}.json"
        if not path.exists():
            return

        with open(path, "r", encoding="utf-8") as f:
            fields = json.load(f)

        for field in fields:
            self._add_row(field["name"], field["type"])

    def _add_row(self, name="", dtype="Строка"):
        row = self.table.rowCount()
        self.table.insertRow(row)

        self.table.setItem(row, 0, QTableWidgetItem(name))

        combo = QComboBox()
        combo.addItems(FIELD_TYPES)
        if dtype in FIELD_TYPES:
            combo.setCurrentText(dtype)
        self.table.setCellWidget(row, 1, combo)

    def _del_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def _save(self):
        fields = []
        for i in range(self.table.rowCount()):
            name_item = self.table.item(i, 0)
            combo = self.table.cellWidget(i, 1)
            if name_item and combo:
                fields.append({
                    "name": name_item.text(),
                    "type": combo.currentText()
                })
        path = SCHEMA_DIR / f"{self.tab_name}.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(fields, f, ensure_ascii=False, indent=2)
