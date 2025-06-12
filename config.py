import os
from pathlib import Path
from core.logger import logger  # инициализация логирования
from core.com_bridge import COM1CBridge
from core import config_parser

# Base directory of the project
BASE_DIR = Path(__file__).resolve().parent

# Path to 1C database
ONEC_PATH = os.getenv("ONEC_PATH", "C:/Users/Mor/Desktop/1C/proiz")

# Глобальный экземпляр COM‑моста
BRIDGE = None


def init_bridge(user: str, password: str, base_path: str | None = None):
    """Инициализирует подключение к 1С с указанными учётными данными."""
    global BRIDGE
    if base_path is None:
        base_path = ONEC_PATH
    BRIDGE = COM1CBridge(base_path, usr=user, pwd=password)
    return BRIDGE

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

# Style for login dialog
LOGIN_DIALOG_CSS = """
QDialog{background:#1f2937;}
QLabel#title{color:#e5e7eb;font-size:18px;font-weight:600;}
QLabel{color:#e5e7eb;font-size:14px;}
QLineEdit,QComboBox{background:#374151;color:#e5e7eb;border:1px solid #4b5563;
                    border-radius:6px;padding:4px;}
QPushButton{background:#3b82f6;color:#ffffff;border:none;padding:6px 12px;
            border-radius:4px;}
QPushButton:hover{background:#2563eb;}
"""

# Список логинов сотрудников для выпадающего меню
def load_employee_logins() -> list[str]:

    """Возвращает список логинов сотрудников через COM с резервным чтением."""
    fallback = ["Администратор"]

    try:
        bridge = BRIDGE or COM1CBridge(ONEC_PATH, usr="Администратор", pwd="")
        items = bridge.list_catalog_items("Пользователи", 1000)
        logins = [it.get("Description", "") for it in items if it.get("Description")]
        if logins:
            return logins
    except Exception as exc:
        logger.error("Не удалось получить логины через COM: %s", exc)

    try:
        users = config_parser.get_catalog_items("Пользователи")
        return users or fallback
    except Exception as exc:
        logger.error("Не удалось загрузить логины из XML: %s", exc)
        return fallback


    """Загружает логины сотрудников из дампа конфигурации."""
    try:
        users = config_parser.get_catalog_items("Пользователи")
        return users or ["Администратор"]
    except Exception as exc:  # безопасность на случай проблем парсинга
        logger.error("Не удалось загрузить логины сотрудников: %s", exc)
        return ["Администратор"]


EMPLOYEE_LOGINS = load_employee_logins()


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

# Число знаков после запятой при отображении веса
WEIGHT_DECIMALS = 3
