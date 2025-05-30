#!/usr/bin/env python
# -*- coding: utf-8 -*-
##############################################################################
#  Jewelry MES — оболочка + страница «Заказы»                 • PyQt5 • v0.3b
##############################################################################

import sys
from pathlib import Path

# Добавляем пути для импорта
base_dir = Path(__file__).parent.resolve()
sys.path.append(str(base_dir))
sys.path.append(str(base_dir / "pages"))
sys.path.append(str(base_dir / "core"))
sys.path.append(str(base_dir / "logic"))

from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QFont, QCursor
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QListWidget, QStackedWidget,
    QVBoxLayout, QHBoxLayout, QToolButton
)

# Страницы
from pages.orders_page import OrdersPage
from pages.wax_page import WaxPage

# ──────────────────────────────── UI Константы ────────────────────────────────
APP = "Jewelry MES (shell-only)"
VER = "v0.3b"

MENU_ITEMS = [
    ("📄  Заказы",            "orders"),
    ("🖨️  Воскование / 3D печать","wax"),
    ("🔥  Отливка",           "casting"),
    ("📥  Приём литья",       "casting_in"),
    ("📦  Комплектация",      "kit"),
    ("🛠️  Монтировка",        "assembly"),
    ("🪚  Шкурка",            "sanding"),
    ("🔄  Галтовка",          "tumbling"),
    ("💎  Закрепка",          "stone_set"),
    ("📏  Палата",            "inspection"),
    ("✨  Полировка",         "polish"),
    ("⚡  Гальваника",        "plating"),
    ("📑  Выпуск",            "release"),
    ("📤  Отгрузка",          "shipment"),
    ("📊  Статистика",        "stats"),
    ("🏬  Склады",            "stock"),
    ("🗺️  Маршруты",          "routes"),
    ("🗓️  Планирование",      "planning"),
    ("💰  Зарплата",          "payroll"),
    ("🏷️  Маркировка",        "marking"),
    ("🌐  ГИИС ДМДК",         "giis"),
    ("📚  Справочники",       "catalogs"),
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

# ───────────────────────  Заглушка страницы  ──────────────────────────────
class StubPage(QWidget):
    def __init__(self, title: str):
        super().__init__()
        v = QVBoxLayout(self)
        v.setContentsMargins(40, 30, 40, 30)
        lbl = QLabel(title)
        lbl.setFont(QFont("Arial", 22, QFont.Bold))
        v.addWidget(lbl, alignment=Qt.AlignTop)

# ───────────────────────  Главное окно  ──────────────────────────────
class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP} — {VER}")
        self.resize(1400, 800)

        # Оболочка
        central = QWidget(); self.setCentralWidget(central)
        outer = QVBoxLayout(central); outer.setContentsMargins(0, 0, 0, 0)

        # Верхняя панель
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

        # Тело интерфейса
        body = QWidget()
        body_lay = QHBoxLayout(body); body_lay.setContentsMargins(0, 0, 0, 0)

        self.menu = QListWidget(); self.menu.setStyleSheet(SIDEBAR_CSS)
        self.menu.setFixedWidth(SIDEBAR_WIDTH)

        self.pages = QStackedWidget()
        body_lay.addWidget(self.menu)
        body_lay.addWidget(self.pages, 1)
        outer.addWidget(body, 1)

        # Добавляем страницы
        for title, key in MENU_ITEMS:
            self.menu.addItem(title)
            if key == "orders":    page = OrdersPage()
            elif key == "wax":     page = WaxPage()
            else:                  page = StubPage(title.strip())
            self.pages.addWidget(page)

        self.menu.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.menu.setCurrentRow(0)
        self.sidebar_open = True

    def toggle_sidebar(self):
        self.sidebar_open = not self.sidebar_open
        self.menu.setVisible(self.sidebar_open)
        self.btn_toggle.setText("◀" if self.sidebar_open else "▶")
        self.btn_toggle.setToolTip("Свернуть меню" if self.sidebar_open else "Развернуть меню")

# ───────────────────────────────  main()  ────────────────────────────────
if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    app = QApplication(sys.argv); app.setStyle("Fusion")
    win = Main(); win.show(); sys.exit(app.exec_())
