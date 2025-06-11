
#!/usr/bin/env python
# -*- coding: utf-8 -*-
##############################################################################
#  Jewelry MES — оболочка + страницы интерфейса                     • PyQt5 •
##############################################################################

import sys
from pathlib import Path

# Добавляем пути для импорта
base_dir = Path(__file__).parent.resolve()
sys.path.append(str(base_dir))
sys.path.append(str(base_dir / "pages"))
sys.path.append(str(base_dir / "core"))
sys.path.append(str(base_dir / "logic"))

from config import (
    BRIDGE as bridge,
    APP_NAME as APP,
    APP_VERSION as VER,
    MENU_ITEMS,
    HEADER_HEIGHT as HEADER_H,
    SIDEBAR_WIDTH,
    HEADER_CSS,
    SIDEBAR_CSS,
)
from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QFont, QCursor
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QListWidget, QStackedWidget,
    QVBoxLayout, QHBoxLayout, QToolButton
)

# Страницы
from pages.orders_page import OrdersPage
from pages.wax_page import WaxPage
from pages.catalogs_page import CatalogsPage

# Заглушка
class StubPage(QWidget):
    def __init__(self, title: str):
        super().__init__()
        v = QVBoxLayout(self)
        v.setContentsMargins(40, 30, 40, 30)
        lbl = QLabel(title)
        lbl.setFont(QFont("Arial", 22, QFont.Bold))
        v.addWidget(lbl, alignment=Qt.AlignTop)

# Главное окно
class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP} — {VER}")
        self.resize(1400, 800)

        central = QWidget(); self.setCentralWidget(central)
        outer = QVBoxLayout(central); outer.setContentsMargins(0, 0, 0, 0)

        header = QWidget(); header.setFixedHeight(HEADER_H)
        header.setStyleSheet(HEADER_CSS)
        h_lay = QHBoxLayout(header); h_lay.setContentsMargins(6, 0, 10, 0)

        self.btn_toggle = QToolButton()
        self.btn_toggle.setText("◀")
        self.btn_toggle.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_toggle.setToolTip("Свернуть меню")
        self.btn_toggle.clicked.connect(self.toggle_sidebar)

        brand = QLabel("Рост Золото"); brand.setObjectName("brand")

        h_lay.addWidget(self.btn_toggle, alignment=Qt.AlignLeft)
        h_lay.addWidget(brand, alignment=Qt.AlignLeft)
        h_lay.addStretch(1)
        outer.addWidget(header)

        body = QWidget()
        body_lay = QHBoxLayout(body); body_lay.setContentsMargins(0, 0, 0, 0)

        self.menu = QListWidget(); self.menu.setStyleSheet(SIDEBAR_CSS)
        self.menu.setFixedWidth(SIDEBAR_WIDTH)

        self.pages = QStackedWidget()
        body_lay.addWidget(self.menu)
        body_lay.addWidget(self.pages, 1)
        outer.addWidget(body, 1)

        self.page_idx = {}
        self.page_refs = {}
        wax_index = next(i for i, (_, k) in enumerate(MENU_ITEMS) if k == "wax")
        for idx, (title, key) in enumerate(MENU_ITEMS):
            self.menu.addItem(title)
            if key == "orders":
                def _open_wax_with_order(order_num=None):
                    self.page_refs["wax"].refresh()
                    if order_num:
                        self.page_refs["wax"].task_form.load_order_by_number(order_num)
                    self.menu.setCurrentRow(self.page_idx.get("wax", wax_index))

                page = OrdersPage(on_send_to_wax=_open_wax_with_order)
            elif key == "wax":
                page = WaxPage()
            elif key == "catalogs":
                page = CatalogsPage(bridge)  # ← вот это обязательно!
            else:
                page = StubPage(title.strip())
            self.pages.addWidget(page)
            self.page_idx[key] = idx
            self.page_refs[key] = page

        self.menu.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.menu.setCurrentRow(0)
        self.sidebar_open = True

    def toggle_sidebar(self):
        self.sidebar_open = not self.sidebar_open
        self.menu.setVisible(self.sidebar_open)
        self.btn_toggle.setText("◀" if self.sidebar_open else "▶")
        self.btn_toggle.setToolTip("Свернуть меню" if self.sidebar_open else "Развернуть меню")


# Точка входа
if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    app = QApplication(sys.argv); app.setStyle("Fusion")
    win = Main(); win.show(); sys.exit(app.exec_())
