
import os
import json
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QLabel, QLineEdit,
    QComboBox, QGroupBox, QDateEdit, QMessageBox
)
from PyQt5.QtCore import Qt


def render_form(tab_key):
    structured_file = os.path.join(os.path.dirname(__file__), "structured_form_fields.json")
    if not os.path.exists(structured_file):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel(f"Файл {structured_file} не найден."))
        return widget

    with open(structured_file, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
        except Exception as e:
            widget = QWidget()
            layout = QVBoxLayout(widget)
            layout.addWidget(QLabel(f"Ошибка чтения JSON: {e}"))
            return widget

    tab_data = data.get(tab_key)
    if not tab_data:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel(f"Нет данных для вкладки '{tab_key}'"))
        return widget

    forms = tab_data.get("forms", {})
    if not forms:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel(f"Нет форм для вкладки '{tab_key}'"))
        return widget

    wrapper = QWidget()
    layout = QVBoxLayout(wrapper)
    layout.setContentsMargins(20, 20, 20, 20)
    layout.setSpacing(16)

    for form_path, fields in forms.items():
        group = QGroupBox(form_path)
        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignRight)
        form_layout.setFormAlignment(Qt.AlignTop)
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(8)

        for field in fields:
            name = field.get("имя", "Без имени")
            field_type = field.get("тип", "Строка")

            label = QLabel(name)
            label.setMinimumWidth(200)

            if field_type == "Строка":
                editor = QLineEdit()
            elif field_type == "Дата":
                editor = QDateEdit()
                editor.setCalendarPopup(True)
            elif field_type == "Выпадающий список":
                editor = QComboBox()
                editor.addItems(["Выбор 1", "Выбор 2", "Выбор 3"])  # TODO: заменить на реальные
            elif field_type == "Множественный выбор":
                editor = QComboBox()
                editor.setEditable(True)
                editor.addItems(["Вариант A", "Вариант B", "Вариант C"])  # TODO: заменить на реальные
            else:
                editor = QLineEdit()

            form_layout.addRow(label, editor)

        group.setLayout(form_layout)
        layout.addWidget(group)

    return wrapper