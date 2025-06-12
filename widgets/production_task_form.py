from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QFormLayout, QHBoxLayout, QComboBox,
    QDateEdit, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QInputDialog
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QDate, pyqtSignal
from core.com_bridge import safe_str, log
import getpass


class ProductionTaskEditForm(QWidget):
    """Форма редактирования документа 'Задание на производство'."""

    task_saved = pyqtSignal(object)

    COLS = [
        "№", "Дата начала", "Дата окончания", "Артикул", "Раб. центр",
        "Номенклатура", "Вариант", "Размер", "Кол-во", "Вес",
        "Проба", "Цвет металла", "Вставки", "Заказ"
    ]

    def __init__(self, bridge):
        super().__init__()
        self.bridge = bridge
        self._order_ref = None
        self._build_ui()
        self.c_center.setCurrentText(getpass.getuser())

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        hdr = QLabel("Задание на производство")
        hdr.setFont(QFont("Arial", 16, QFont.Bold))
        layout.addWidget(hdr)

        form = QFormLayout()
        self.lbl_number = QLabel(self.bridge.get_next_task_number())
        self.d_date = QDateEdit(QDate.currentDate()); self.d_date.setCalendarPopup(True)
        self.d_end = QDateEdit(QDate.currentDate()); self.d_end.setCalendarPopup(True)

        self.c_section = QComboBox()
        sections = [x["Description"] for x in self.bridge.list_catalog_items("ПроизводственныеУчастки")]
        self.c_section.addItems(sections)

        self.c_op = QComboBox()
        ops = [x["Description"] for x in self.bridge.list_catalog_items("ТехОперации")]
        self.c_op.addItems(ops)

        self.c_center = QComboBox()
        centers = [x["Description"] for x in self.bridge.list_catalog_items("ФизическиеЛица", limit=200)]
        self.c_center.addItems(centers)

        self.order_line = QLineEdit(); self.order_line.setReadOnly(True)
        self.btn_order = QPushButton("Выбрать")
        self.btn_order.clicked.connect(self.request_order_selection)
        h_order = QHBoxLayout(); h_order.addWidget(self.order_line); h_order.addWidget(self.btn_order)

        form.addRow("Номер", self.lbl_number)
        form.addRow("Дата", self.d_date)
        form.addRow("По", self.d_end)
        form.addRow("Произ. участок", self.c_section)
        form.addRow("Тех. операция", self.c_op)
        form.addRow("Рабочий центр", self.c_center)
        form.addRow("Заказ", h_order)
        layout.addLayout(form)

        self.tbl = QTableWidget(0, len(self.COLS))
        self.tbl.setHorizontalHeaderLabels(self.COLS)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl.verticalHeader().setVisible(False)
        layout.addWidget(self.tbl, 1)

        btns = QHBoxLayout()
        self.btn_save = QPushButton("💾 Записать")
        self.btn_post = QPushButton("✅ Провести")
        self.btn_add = QPushButton("Строка +")
        self.btn_del = QPushButton("Строка -")
        self.btn_copy = QPushButton("Копировать строку")
        self.btn_new = QPushButton("Новое задание")

        self.btn_save.clicked.connect(self.save_task)
        self.btn_post.clicked.connect(self.post_task)
        self.btn_add.clicked.connect(self.add_row)
        self.btn_del.clicked.connect(self.remove_row)
        self.btn_copy.clicked.connect(self.copy_row)
        self.btn_new.clicked.connect(self.new_task)

        for b in [self.btn_save, self.btn_post, self.btn_add, self.btn_del, self.btn_copy, self.btn_new]:
            btns.addWidget(b)
        layout.addLayout(btns)

        form2 = QFormLayout()
        self.ed_comment = QLineEdit()
        self.lbl_status = QLabel("Черновик")
        form2.addRow("Комментарий", self.ed_comment)
        form2.addRow("Статус", self.lbl_status)
        layout.addLayout(form2)

    # ------------------------------------------------------------------
    def new_task(self):
        self.lbl_number.setText(self.bridge.get_next_task_number())
        self.d_date.setDate(QDate.currentDate())
        self.d_end.setDate(QDate.currentDate())
        self.order_line.clear()
        self.tbl.setRowCount(0)
        self.ed_comment.clear()
        self.lbl_status.setText("Черновик")
        self._order_ref = None
        self.c_center.setCurrentText(getpass.getuser())

    def add_row(self, copy_from=None):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)
        for c in range(len(self.COLS)):
            self.tbl.setItem(r, c, QTableWidgetItem(""))
        if copy_from is not None and 0 <= copy_from < r:
            for c in range(len(self.COLS)):
                src = self.tbl.item(copy_from, c)
                if src:
                    self.tbl.item(r, c).setText(src.text())

    def copy_row(self):
        row = self.tbl.currentRow()
        if row >= 0:
            self.add_row(copy_from=row)

    def remove_row(self):
        row = self.tbl.currentRow()
        if row >= 0:
            self.tbl.removeRow(row)

    def request_order_selection(self):
        """Запрашивает выбор заказа на внешней вкладке."""
        parent = self.parent()
        while parent and not hasattr(parent, "goto_order_selection"):
            parent = parent.parent() if hasattr(parent, "parent") else None

        if parent and hasattr(parent, "goto_order_selection"):
            parent.goto_order_selection(callback=self._on_order_selected)

    def _on_order_selected(self, order):
        if isinstance(order, dict):
            self.load_order_by_number(order.get("num", ""), order.get("date"))
        else:
            self.load_order_by_number(str(order))

    def load_from_order(self):
        orders = self.bridge.list_orders()
        nums = [o.get("num") for o in orders]
        num, ok = QInputDialog.getItem(self, "Выбор заказа", "Заказ", nums, 0, False)
        if not ok or not num:
            return
        self.load_order_by_number(num)
        self.lbl_status.setText("Заполнено")

    def load_order_by_number(self, order_number: str, date: str | None = None):
        """Загружается из внешнего выбора (например, из вкладки Заказы)."""
        self.order_line.setText(order_number)
        self._order_ref = self.bridge.get_doc_ref("ЗаказВПроизводство", order_number, date)
        self.c_center.setCurrentText(getpass.getuser())
        self._load_lines(order_number, date)

        self.lbl_status.setText("Заполнено")

    def load_order_dict(self, order: dict):
        """Загружает заказ из словаря, полученного с вкладки Заказы."""
        if not order:
            return
        self.order_line.setText(order.get("num", ""))
        self._order_ref = order.get("Ref")
        self.c_center.setCurrentText(getpass.getuser())
        self.tbl.setRowCount(0)
        for line in order.get("rows", []):
            self.add_row()
            r = self.tbl.rowCount() - 1
            self.tbl.item(r, 0).setText(str(r + 1))
            self.tbl.item(r, 1).setText(QDate.currentDate().toString("dd.MM.yyyy"))
            self.tbl.item(r, 2).setText(QDate.currentDate().addDays(1).toString("dd.MM.yyyy"))
            art = line.get("article") or line.get("nomenclature", "")
            self.tbl.item(r, 3).setText(art)
            self.tbl.item(r, 4).setText(self.c_center.currentText())
            self.tbl.item(r, 5).setText(line.get("nomenclature", ""))
            self.tbl.item(r, 6).setText(line.get("variant", ""))
            self.tbl.item(r, 7).setText(str(line.get("size", "")))
            self.tbl.item(r, 8).setText(str(line.get("qty", "")))
            self.tbl.item(r, 9).setText(str(line.get("w", "")))
            self.tbl.item(r, 10).setText("")
            self.tbl.item(r, 11).setText("")
            self.tbl.item(r, 12).setText("")
            self.tbl.item(r, 13).setText(order.get("num", ""))
        self.lbl_status.setText("Заполнено")

    def load_task_object(self, task_obj):
        """Заполняет форму данными выбранного задания."""
        if not task_obj:
            return

        self.lbl_number.setText(str(getattr(task_obj, "Номер", "")))
        try:
            d_start = QDate.fromString(str(task_obj.Дата)[:10], "yyyy-MM-dd")
            self.d_date.setDate(d_start)
        except Exception:
            pass
        try:
            d_end = QDate.fromString(str(task_obj.ДатаОкончания)[:10], "yyyy-MM-dd")
            self.d_end.setDate(d_end)
        except Exception:
            pass

        # --- преобразуем COM-поля правильно ---
        def to_str(val):
            return safe_str(val.Description) if hasattr(val, "Description") else safe_str(val)

        self._set_combo_value(self.c_section, to_str(getattr(task_obj, "ПроизводственныйУчасток", "")))
        self._set_combo_value(self.c_op, to_str(getattr(task_obj, "ТехОперация", "")))
        self._set_combo_value(self.c_center, to_str(getattr(task_obj, "РабочийЦентр", getpass.getuser())))

        base = getattr(task_obj, "ДокументОснование", None)
        if base:
            try:
                base_obj = base.GetObject() if hasattr(base, "GetObject") else base

                order_number = str(getattr(base_obj, "Номер", ""))
                self.order_line.setText(order_number)

                self.order_line.setText(str(getattr(base_obj, "Номер", "")))

                self._order_ref = base
            except Exception as e:  # noqa: PIE786
                log(f"❌ Ошибка чтения осн. документа: {e}")

        self.tbl.setRowCount(0)
        lines = self.bridge.get_task_lines(str(getattr(task_obj, "Номер", "")))
        for line in lines:
            self.add_row()
            r = self.tbl.rowCount() - 1
            self.tbl.item(r, 0).setText(str(r + 1))
            self.tbl.item(r, 1).setText(self.d_date.date().toString("dd.MM.yyyy"))
            self.tbl.item(r, 2).setText(self.d_end.date().toString("dd.MM.yyyy"))
            self.tbl.item(r, 3).setText(line.get("article", ""))
            self.tbl.item(r, 4).setText(self.c_center.currentText())
            self.tbl.item(r, 5).setText(line.get("nomen", ""))
            self.tbl.item(r, 6).setText(line.get("method", ""))
            self.tbl.item(r, 7).setText(str(line.get("size", "")))
            self.tbl.item(r, 8).setText(str(line.get("qty", "")))
            self.tbl.item(r, 9).setText(str(line.get("weight", "")))
            self.tbl.item(r, 10).setText(line.get("sample", ""))
            self.tbl.item(r, 11).setText(line.get("color", ""))
            self.tbl.item(r, 12).setText("")
            self.tbl.item(r, 13).setText(self.order_line.text())

        self.lbl_status.setText("Заполнено")


    def _load_lines(self, num, date: str | None = None):
        lines = self.bridge.get_order_lines(num, date)
        self.tbl.setRowCount(0)
        for line in lines:
            self.add_row()
            r = self.tbl.rowCount() - 1
            self.tbl.item(r, 0).setText(str(r + 1))
            self.tbl.item(r, 1).setText(QDate.currentDate().toString("dd.MM.yyyy"))
            self.tbl.item(r, 2).setText(QDate.currentDate().addDays(1).toString("dd.MM.yyyy"))
            self.tbl.item(r, 3).setText(line.get("article", ""))
            self.tbl.item(r, 4).setText(self.c_center.currentText())
            self.tbl.item(r, 5).setText(line.get("name", ""))
            self.tbl.item(r, 6).setText(line.get("method", ""))
            self.tbl.item(r, 7).setText(str(line.get("size", "")))
            self.tbl.item(r, 8).setText(str(line.get("qty", "")))
            self.tbl.item(r, 9).setText(str(line.get("weight", "")))
            self.tbl.item(r, 10).setText(line.get("assay", ""))
            self.tbl.item(r, 11).setText(line.get("color", ""))
            self.tbl.item(r, 12).setText(line.get("insert", ""))
            self.tbl.item(r, 13).setText(num)

    def _collect_rows(self):
        rows = []
        for r in range(self.tbl.rowCount()):
            rows.append({
                "article": self._cell_text(r, 3),
                "name": self._cell_text(r, 5),
                "method": self._cell_text(r, 6),
                "size": self._cell_text(r, 7),
                "qty": float(self._cell_text(r, 8) or 0),
                "weight": float(self._cell_text(r, 9) or 0),
                "assay": self._cell_text(r, 10),
                "color": self._cell_text(r, 11),
                "insert": self._cell_text(r, 12),
                "employee": self.c_center.currentText(),
            })
        return rows

    def _cell_text(self, row, col):
        item = self.tbl.item(row, col)
        return item.text().strip() if item else ""

    def _set_combo_value(self, combo: QComboBox, text: str) -> None:
        """Устанавливает значение в выпадающем списке по тексту."""
        idx = combo.findText(text)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        else:
            if text:
                combo.addItem(text)
                combo.setCurrentIndex(combo.count() - 1)

    def save_task(self):
        if not self._order_ref:
            QMessageBox.warning(self, "Ошибка", "Заказ не выбран")
            return
        rows = self._collect_rows()
        res = self.bridge.create_production_task(self._order_ref, rows)
        if res:
            self.lbl_number.setText(res.get("Номер", ""))
            self.lbl_status.setText("Записан")
            QMessageBox.information(self, "Готово", f"Создано задание №{res.get('Номер')}")
            self.task_saved.emit(res.get("Ref"))
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось создать задание")

    def post_task(self):
        number = self.lbl_number.text().strip()
        if not number:
            QMessageBox.warning(self, "Ошибка", "Нет номера задания")
            return
        if self.bridge.post_task(number):
            self.lbl_status.setText("Проведено")
            QMessageBox.information(self, "Готово", f"Задание №{number} проведено")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось провести задание")

