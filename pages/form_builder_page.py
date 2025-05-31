from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit,
    QTreeWidget, QTreeWidgetItem, QHBoxLayout, QComboBox
)
import json
from pathlib import Path
from core.form_builder import open_form_builder

STRUCTURED_FIELDS_PATH = Path("structured_form_fields.json")
SCHEMA_DIR = Path("form_schemas")

class FormBuilderPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.fields = self._load_fields()
        self.tabs = self._load_tabs()
        self.current_tab = None
        self.init_ui()

    def _load_fields(self):
        if STRUCTURED_FIELDS_PATH.exists():
            with open(STRUCTURED_FIELDS_PATH, encoding="utf-8") as f:
                return json.load(f)
        return {}

    def _load_tabs(self):
        return [f.stem for f in SCHEMA_DIR.glob("*.json")]

    def init_ui(self):
        layout = QVBoxLayout(self)

        hlayout = QHBoxLayout()
        hlayout.addWidget(QLabel("Выберите вкладку:"))
        self.cmb_tabs = QComboBox()
        self.cmb_tabs.addItems(self.tabs)
        self.cmb_tabs.currentTextChanged.connect(self._on_tab_changed)
        hlayout.addWidget(self.cmb_tabs)
        layout.addLayout(hlayout)

        tree_layout = QHBoxLayout()

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Структура полей")
        self.tree.itemClicked.connect(self._on_field_selected)
        tree_layout.addWidget(self.tree, 2)

        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        tree_layout.addWidget(self.preview, 1)

        layout.addLayout(tree_layout)

        btns = QHBoxLayout()
        self.btn_add = QPushButton("➕ Добавить в форму")
        self.btn_edit = QPushButton("⚙️ Редактировать форму")
        self.btn_add.clicked.connect(self._add_field)
        self.btn_edit.clicked.connect(self._open_editor)
        btns.addWidget(self.btn_add)
        btns.addStretch()
        btns.addWidget(self.btn_edit)
        layout.addLayout(btns)

        self._on_tab_changed(self.cmb_tabs.currentText())

    def _on_tab_changed(self, tab):
        self.current_tab = tab
        self.tree.clear()
        self.preview.clear()
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

    def _on_field_selected(self, item):
        path = item.data(0, 1)
        if path:
            self.preview.setText(f"{path}")

    def _add_field(self):
        if not self.current_tab:
            return
        path = self.preview.toPlainText().strip()
        if not path:
            return
        schema_path = SCHEMA_DIR / f"{self.current_tab}.json"
        fields = []
        if schema_path.exists():
            with open(schema_path, "r", encoding="utf-8") as f:
                fields = json.load(f)

        field_name = path.split(".")[-1]
        fields.append({"name": field_name, "field": path, "type": "Строка"})

        with open(schema_path, "w", encoding="utf-8") as f:
            json.dump(fields, f, ensure_ascii=False, indent=2)

        self.preview.setText(f"Добавлено: {field_name} → {path}")

    def _open_editor(self):
        if self.current_tab:
            open_form_builder(self.current_tab)
