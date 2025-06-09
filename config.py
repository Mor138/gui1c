from pathlib import Path
from core.com_bridge import COM1CBridge

# Base directory of the project
BASE_DIR = Path(__file__).resolve().parent

# Path to 1C database
ONEC_PATH = "C:/Users/Mor/Desktop/1C/proiz"

# Global COM bridge instance
BRIDGE = COM1CBridge(ONEC_PATH)

# Application
APP_NAME = "Jewelry MES (shell-only)"
APP_VERSION = "v0.3b"

# Menu structure used in main window
MENU_ITEMS = [
    ("📄  Заказы", "orders"),
    ("🖨️  Воскование / 3D печать", "wax"),
    ("🔥  Отливка", "casting"),
    ("📥  Приём литья", "casting_in"),
    ("📦  Комплектация", "kit"),
    ("🛠️  Монтировка", "assembly"),
    ("🪚  Шкурка", "sanding"),
    ("🔄  Галтовка", "tumbling"),
    ("💎  Закрепка", "stone_set"),
    ("📏  Палата", "inspection"),
    ("✨  Полировка", "polish"),
    ("⚡  Гальваника", "plating"),
    ("📑  Выпуск", "release"),
    ("📤  Отгрузка", "shipment"),
    ("📊  Статистика", "stats"),
    ("🏬  Склады", "stock"),
    ("🗺️  Маршруты", "routes"),
    ("🗓️  Планирование", "planning"),
    ("💰  Зарплата", "payroll"),
    ("🏷️  Маркировка", "marking"),
    ("🌐  ГИИС ДМДК", "giis"),
    ("📚  Справочники", "catalogs"),
]

# UI constants
HEADER_HEIGHT = 38
SIDEBAR_WIDTH = 260

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

# Style for tree widgets used on the wax page
CSS_TREE = """
QTreeWidget{
  background:#ffffff;
  border:1px solid #d1d5db;
  color:#111827;
  font-size:14px;
}
QTreeWidget::item{
  padding:4px 8px;
  border-bottom:1px solid #e5e7eb;
}

/* выделение строки */
QTreeView::item:selected{
  background:#3b82f6;
  color:#ffffff;
}

/* hover */
QTreeView::item:hover:!selected{
  background:rgba(59,130,246,0.30);
}

/*  — если хотите зебру, раскомментируйте ↓ —
QTreeView::item:nth-child(even):!selected{ background:#f9fafb; }
QTreeView::item:nth-child(odd):!selected { background:#ffffff; }
*/
"""

# Columns for the orders table
ORDERS_COLS = ["Артикул", "Наим.", "Вариант", "Размер", "Кол-во", "Вес, г", "Примечание"]
