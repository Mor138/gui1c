from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QComboBox, QTreeWidget, QTreeWidgetItem, QTableWidget,
    QTableWidgetItem, QMessageBox
)
import json
import os


class FormBuilderPage(QWidget):
    def __init__(self, structured_file="structured_form_fields.json"):
        super().__init__()
        self.setWindowTitle("Конструктор форм")
        self.structured_file = structured_file
        self.data = self.load_structured_file()
        self.current_tab = None
        self.current_fields = []

        self.tab_display_names = {
            "orders": "📄 Заказы",
            "wax": "🖨️ Воскование / 3D печать",
            "casting": "🔥 Отливка",
            "casting_in": "📥 Приём литья",
            "kit": "📦 Комплектация",
            "assembly": "🛠️ Монтировка",
            "sanding": "🪚 Шкурка",
            "tumbling": "🔄 Галтовка",
            "stone_set": "💎 Закрепка",
            "inspection": "📏 Палата",
            "polish": "✨ Полировка",
            "plating": "⚡ Гальваника",
            "release": "📑 Выпуск",
            "shipment": "📤 Отгрузка",
            "stats": "📊 Статистика",
            "stock": "🏬 Склады",
            "routes": "🗺️ Маршруты",
            "planning": "🗓️ Планирование",
            "payroll": "💰 Зарплата",
            "marking": "🏷️ Маркировка",
            "giis": "🌐 ГИИС ДМДК",
            "catalogs": "📚 Справочники",
            "form_builder": "🔧 Конструктор форм"
        }

        self.tab_order = list(self.tab_display_names.keys())

        self.init_ui()

    def load_structured_file(self):
        if os.path.exists(self.structured_file):
            with open(self.structured_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def init_ui(self):
        layout = QVBoxLayout()

        top_row = QHBoxLayout()
        self.tab_combo = QComboBox()
        self.tab_display_to_key = {}
        for k in self.tab_order:
            if k in self.data:
                display = self.tab_display_names.get(k, k)
                self.tab_combo.addItem(display)
                self.tab_display_to_key[display] = k
        self.tab_combo.currentTextChanged.connect(self.tab_selected)
        top_row.addWidget(QLabel("Выберите вкладку:"))
        top_row.addWidget(self.tab_combo)

        self.reload_button = QPushButton("🔁 Обновить файл")
        self.reload_button.clicked.connect(self.reload_data)
        top_row.addWidget(self.reload_button)

        layout.addLayout(top_row)

        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("🔍 Поиск по полям...")
        self.search_field.textChanged.connect(self.apply_filter)
        layout.addWidget(self.search_field)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Структура полей")
        layout.addWidget(self.tree)

        self.edit_button = QPushButton("🔧 Редактировать форму")
        self.edit_button.clicked.connect(self.open_editor)
        layout.addWidget(self.edit_button)

        self.setLayout(layout)
        self.tab_selected(self.tab_combo.currentText())

    def reload_data(self):
        self.data = self.load_structured_file()
        self.tab_combo.clear()
        self.tab_display_to_key.clear()
        for k in self.tab_order:
            if k in self.data:
                display = self.tab_display_names.get(k, k)
                self.tab_combo.addItem(display)
                self.tab_display_to_key[display] = k
        # принудительно перевыбрать текущую вкладку
        current = self.tab_combo.currentText()
        self.tab_combo.setCurrentIndex(-1)
        self.tab_combo.setCurrentText(current)

    def tab_selected(self, display_name):
        tab_key = self.tab_display_to_key.get(display_name, display_name)
        self.current_tab = tab_key
        self.refresh_tree()

    def refresh_tree(self):
        self.tree.clear()
        filter_text = self.search_field.text().lower()
        forms = self.data.get(self.current_tab, {}).get("forms", {})
        for form_path, fields in forms.items():
            match_fields = [f for f in fields if filter_text in f.get("имя", "").lower()]
            if filter_text and not match_fields:
                continue
            form_item = QTreeWidgetItem([form_path])
            self.tree.addTopLevelItem(form_item)
            for field in (match_fields if filter_text else fields):
                f_text = f'{field.get("имя", "")} : {field.get("тип", "")}'
                QTreeWidgetItem(form_item, [f_text])
            form_item.setExpanded(True)

    def apply_filter(self):
        self.refresh_tree()

    def open_editor(self):
        selected = self.tree.currentItem()
        if not selected:
            QMessageBox.warning(self, "Выбор формы", "Выберите элемент дерева.")
            return
        if selected.parent():
            selected = selected.parent()
        form_path = selected.text(0)
        self.edit_form(self.current_tab, form_path)

    def edit_form(self, tab_key, form_path):
        from PyQt5.QtWidgets import QDialog

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Конструктор формы: {tab_key}")

        layout = QVBoxLayout(dlg)
        self.fields_table = QTableWidget()
        self.fields_table.setColumnCount(3)
        self.fields_table.setHorizontalHeaderLabels(["Имя столбца", "Поле данных", "Тип данных"])

        self.current_fields = self.data.get(tab_key, {}).get("forms", {}).get(form_path, [])

        for row, field in enumerate(self.current_fields):
            self.fields_table.setRowCount(row + 1)

            item_name = QTableWidgetItem(field.get("имя", ""))
            item_field = QTableWidgetItem(field.get("имя", ""))
            combo_type = QComboBox()
            combo_type.addItems(["Строка", "Дата", "Выпадающий список", "Множественный выбор"])
            combo_type.setCurrentText(field.get("тип", "Строка"))

            self.fields_table.setItem(row, 0, item_name)
            self.fields_table.setItem(row, 1, item_field)
            self.fields_table.setCellWidget(row, 2, combo_type)

        layout.addWidget(self.fields_table)

        btns = QHBoxLayout()
        add_btn = QPushButton("➕ Добавить")
        save_btn = QPushButton("💾 Сохранить")
        close_btn = QPushButton("Закрыть")
        del_btn = QPushButton("🗑️ Удалить")
        add_btn.clicked.connect(self.add_field)
        save_btn.clicked.connect(lambda: self.save_form(tab_key, form_path))
        close_btn.clicked.connect(dlg.close)
        del_btn.clicked.connect(self.delete_selected_field)

        btns.addWidget(add_btn)
        btns.addWidget(del_btn)
        btns.addWidget(save_btn)
        btns.addWidget(close_btn)

        layout.addLayout(btns)
        dlg.setLayout(layout)
        dlg.exec_()

    def add_field(self):
        row = self.fields_table.rowCount()
        self.fields_table.insertRow(row)
        self.fields_table.setItem(row, 0, QTableWidgetItem(""))
        self.fields_table.setItem(row, 1, QTableWidgetItem(""))
        combo_type = QComboBox()
        combo_type.addItems(["Строка", "Дата", "Выпадающий список", "Множественный выбор"])
        self.fields_table.setCellWidget(row, 2, combo_type)

    def delete_selected_field(self):
        row = self.fields_table.currentRow()
        if row >= 0:
            self.fields_table.removeRow(row)

    def save_form(self, tab_key, form_path):
        fields = []
        for row in range(self.fields_table.rowCount()):
            name = self.fields_table.item(row, 0).text()
            field = self.fields_table.item(row, 1).text()
            type_widget = self.fields_table.cellWidget(row, 2)
            field_type = type_widget.currentText() if type_widget else "Строка"
            fields.append({
                "имя": name,
                "тип": field_type
            })

        self.data[tab_key]["forms"][form_path] = fields
        with open(self.structured_file, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        QMessageBox.information(self, "Сохранено", "Форма успешно сохранена.")
        self.reload_data()