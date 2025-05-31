import json
import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QLineEdit, QComboBox, QDateEdit

def render_form(tab_key):
    structured_file = os.path.join(os.path.dirname(__file__), 'structured_form_fields.json')
    if not os.path.exists(structured_file):
        return QLabel(f"Файл {structured_file} не найден.")

    with open(structured_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    tab_data = data.get(tab_key)
    if not tab_data:
        return QLabel(f"Нет данных для вкладки '{tab_key}'.")

    forms = tab_data.get("forms", {})
    if not forms:
        return QLabel(f"Нет форм для вкладки '{tab_key}'.")

    widget = QWidget()
    layout = QVBoxLayout()

    for form_path, fields in forms.items():
        layout.addWidget(QLabel(f"Форма: {form_path}"))
        for field in fields:
            field_name = field.get("имя", "Без имени")
            field_type = field.get("тип", "Строка")
            layout.addWidget(QLabel(field_name))
            if field_type == "Строка":
                layout.addWidget(QLineEdit())
            elif field_type == "Дата":
                layout.addWidget(QDateEdit())
            elif field_type == "Выпадающий список":
                combo = QComboBox()
                combo.addItems(["Значение 1", "Значение 2"])  # Замените на реальные значения
                layout.addWidget(combo)
            elif field_type == "Множественный выбор":
                combo = QComboBox()
                combo.addItems(["Опция 1", "Опция 2"])  # Замените на реальные значения
                combo.setEditable(True)
                layout.addWidget(combo)
            else:
                layout.addWidget(QLineEdit())

    widget.setLayout(layout)
    return widget
