# JewelryMES/pages/form_builder_page.py

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit,
    QTreeWidget, QTreeWidgetItem, QHBoxLayout, QComboBox
)
import json
from pathlib import Path
from core.form_builder import open_form_builder

STRUCTURED_FIELDS_PATH = Path("structured_form_fields.json")
TABS = ["Заказы", "Планирование", "Отгрузка", "Палата", "Гальваника"]  # можно расширить


class FormBuilderPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.fields = self._load_fields()
        self.init_ui()

    def _load_fields(self):
        if STRUCTURED_FIELDS_PATH.exists():
            with open(STRUCTURED_FIELDS_PATH, encoding="utf-8") as f:
                return json.load(f)
        return {}

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Выбор вкладки
        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("Вкладка:"))
        self.tab_selector = QComboBox()
        self.tab_selector.addItems(TABS)
        top_bar.addWidget(self.tab_selector)
        self.edit_btn = QPushButton("🛠 Редактировать отображение")
        self.edit_btn.clicked.connect(self.open_editor)
        top_bar.addStretch()
        top_bar.addWidget(self.edit_btn)

        layout.addLayout(top_bar)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Структура полей")
        self.tree.itemClicked.connect(self.show_path)
        layout.addWidget(self.tree, 4)

        right = QVBoxLayout()
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.btn_add = QPushButton("Добавить в форму")
        self.btn_add.clicked.connect(self.add_field)
        right.addWidget(QLabel("Информация о поле:"))
        right.addWidget(self.preview, 1)
        right.addWidget(self.btn_add)

        wrap = QHBoxLayout()
        wrap.addLayout(layout, 2)
        wrap.addLayout(right, 1)
        self.setLayout(wrap)

        self._populate_tree()

    def _populate_tree(self):
        self.tree.clear()

        def add_children(parent, struct, prefix=""):
            for key, value in struct.items():
                if isinstance(value, dict) and "field" in value:
                    item = QTreeWidgetItem([f"{key} : {value['field']}"])
                    item.setData(0, 1, prefix + key)
                    parent.addChild(item)
                elif isinstance(value, dict):
                    branch = QTreeWidgetItem([key])
                    parent.addChild(branch)
                    add_children(branch, value, prefix + key + ".")

        for source, subtree in self.fields.items():
            root = QTreeWidgetItem([source])
            self.tree.addTopLevelItem(root)
            add_children(root, subtree)
            root.setExpanded(True)

    def show_path(self, item):
        path = item.data(0, 1)
        if path:
            self.preview.setText(f"DataPath: {path}")

    def add_field(self):
        path = self.preview.toPlainText()
        if path:
            print(f"Добавлено поле: {path}")

    def open_editor(self):
        tab = self.tab_selector.currentText()
        open_form_builder(tab)
