
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

from core.com_bridge import COM1CBridge
bridge = COM1CBridge("C:\\Users\\Mor\\Desktop\\1C\\proiz")
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


APP = "Jewelry MES (shell-only)"
VER = "v0.3b"

MENU_ITEMS = [
    ("📄  Заказы",            "orders"),       # → pages/orders_page.py     + logic/production_docs.py
    ("🖨️  Воскование / 3D печать", "wax"),      # → pages/wax_page.py        + logic/production_docs.py
    ("🔥  Отливка",           "casting"),      # → pages/casting_page.py    [в разработке]
    ("📥  Приём литья",       "casting_in"),   # → pages/casting_in_page.py [в разработке]
    ("📦  Комплектация",      "kit"),          # → pages/kit_page.py        [в разработке]
    ("🛠️  Монтировка",        "assembly"),     # → pages/assembly_page.py   [в разработке]
    ("🪚  Шкурка",            "sanding"),      # → pages/sanding_page.py    [в разработке]
    ("🔄  Галтовка",          "tumbling"),     # → pages/tumbling_page.py   + logic/loss_calc.py
    ("💎  Закрепка",          "stone_set"),    # → pages/stone_set_page.py  + logic/normalize_catalogs.py
    ("📏  Палата",            "inspection"),   # → pages/inspection_page.py + logic/validation.py (возможн.)
    ("✨  Полировка",         "polish"),       # → pages/polish_page.py     [в разработке]
    ("⚡  Гальваника",        "plating"),      # → pages/plating_page.py    [в разработке]
    ("📑  Выпуск",            "release"),      # → pages/release_page.py    + logic/production_docs.py
    ("📤  Отгрузка",          "shipment"),     # → pages/shipment_page.py   [в разработке]
    ("📊  Статистика",        "stats"),        # → pages/stats_page.py      + widgets/charts.py + logic/loss_calc.py
    ("🏬  Склады",            "stock"),        # → pages/stock_page.py      + core/com_bridge.py
    ("🗺️  Маршруты",          "routes"),       # → pages/routes_page.py     [в разработке]
    ("🗓️  Планирование",      "planning"),     # → pages/planning_page.py   [в разработке]
    ("💰  Зарплата",          "payroll"),      # → pages/payroll_page.py    [в разработке]
    ("🏷️  Маркировка",        "marking"),      # → pages/marking_page.py    [в разработке]
    ("🌐  ГИИС ДМДК",         "giis"),         # → pages/giis_page.py       [в разработке]
    ("📚  Справочники",       "catalogs")      # → pages/catalogs_page.py   + logic/normalize_catalogs.py + core/catalogs.py
]

HEADER_H       = 38
SIDEBAR_WIDTH  = 260

HEADER_CSS = """
QWidget{background:#111827;}
QLabel#brand{color:#e5e7eb;font-size:15px;font-weight:600;}
QToolButton{background:#111827;color:#9ca3af;border:none;font-size:16px;}
QToolButton:hover{color:white;}
"""

SIDEBAR_CSS = """
QListWidget{background:#1f2937;border:none;color:#e5e7eb;
            padding-top:6px;font-size:15px;}
QListWidget::item{height:46px;margin:4px 8px;padding-left:14px;
                  border-radius:12px;}
QListWidget::item:selected{background:#3b82f6;color:white;}
QListWidget QScrollBar:vertical{width:0px;background:transparent;}
QListWidget:hover QScrollBar:vertical{width:8px;}
"""

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
                page = OrdersPage(on_send_to_wax=lambda: (self.page_refs["wax"].refresh(), self.menu.setCurrentRow(self.page_idx.get("wax", wax_index))))
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
    
    from core.com_bridge import COM1CBridge

if __name__ == "__main__":
    bridge = COM1CBridge("C:\\path\\to\\your\\base")  # замените на путь к вашей базе
    tasks = bridge.list_production_orders()
    print("Задания:", tasks)

    wax_jobs = bridge.list_wax_work_orders()
    print("Наряды:", wax_jobs)