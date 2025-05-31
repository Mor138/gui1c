from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTabWidget
from widgets.tables import CatalogTable
from logic.normalize_catalogs import NORMALIZED

class CatalogsPage(QWidget):
    def __init__(self, bridge):
        super().__init__()
        self.bridge = bridge
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        tab_widget = QTabWidget()

        tabs = [
            ("🧾 Изделия", self._make_items_tab()),
            ("💍 Вставки", self._make_inserts_tab()),
            ("💎 Камни", self._make_stones_tab()),
        ]

        for label, widget in tabs:
            tab_widget.addTab(widget, label)

        layout.addWidget(tab_widget)

    def _make_items_tab(self):
        return CatalogTable(NORMALIZED["Изделия"], columns=["Название", "Метод"])

    def _make_inserts_tab(self):
        return CatalogTable(NORMALIZED["Вставки"], columns=["Название", "Вставка"])

    def _make_stones_tab(self):
        return CatalogTable(NORMALIZED["Камни"], columns=["Название"])
