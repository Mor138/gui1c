# JewelryMES/pages/form_builder_page.py
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit,
    QTreeWidget, QTreeWidgetItem, QHBoxLayout
)
import json
from pathlib import Path

STRUCTURED_FIELDS_PATH = Path("structured_form_fields.json")

class FormBuilderPage(QWidget):
    """Конструктор форм с древовидным просмотром полей."""
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
        layout = QHBoxLayout(self)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Структура полей")
        self.tree.itemClicked.connect(self.show_path)
        layout.addWidget(self.tree, 2)

        self.preview = QTextEdit()
        self.preview.setReadOnly(True)

        btn = QPushButton("Добавить в форму")
        btn.clicked.connect(self.add_field)

        right = QVBoxLayout()
        right.addWidget(QLabel("Информация о поле:"))
        right.addWidget(self.preview, 1)
        right.addWidget(btn)
        layout.addLayout(right, 1)

        self._populate_tree()

    def _populate_tree(self):
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
            print(f"Добавлено поле: {path}")  # здесь будет логика добавления в форму
