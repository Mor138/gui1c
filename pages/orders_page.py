import pywintypes
from PyQt5.QtCore import QDate, QTimer
from core.com_bridge import safe_str, PRODUCTION_STATUS_MAP
from datetime import datetime
from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QFormLayout, QComboBox, QDateEdit,
    QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QPushButton,
    QMessageBox, QHeaderView, QTabWidget, QHBoxLayout, QAbstractItemView,
    QLineEdit  # ← вот это добавьте
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QDate
from core.com_bridge import COM1CBridge

bridge = COM1CBridge("C:\\Users\\Mor\\Desktop\\1C\\proiz")

class OrdersPage(QWidget):
    COLS = ["Артикул", "Наим.", "Вариант", "Размер", "Кол-во", "Вес, г", "Примечание"]

    def __init__(self):
        super().__init__()
        self.articles = bridge.get_articles()
        self.organizations = bridge.list_catalog_items("Организации")
        self.counterparties = bridge.list_catalog_items("Контрагенты")
        self.contracts = bridge.list_catalog_items("ДоговорыКонтрагентов")
        self.warehouses = bridge.list_catalog_items("Склады")
        self.production_statuses = bridge.PRODUCTION_STATUSES
        self._ui()
        self._load_orders()
        self._edit_mode = False
        self._current_number = ""
        
    def _update_order(self):
        if not self._edit_mode or not self._current_number:
            QMessageBox.warning(self, "Ошибка", "Нет редактируемого заказа.")
            return

        fields = {
            "Организация": self.c_org.currentText(),
            "Контрагент": self.c_contr.currentText(),
            "ДоговорКонтрагента": self.c_ctr.currentText(),
            "Склад": self.c_wh.currentText(),
            "Ответственный": "Администратор",
            "Комментарий": self.comment_input.text().strip(),
            "Дата": pywintypes.Time(self.d_date.date().toPyDate()),
            "ВидСтатусПродукции": str(self.status_combo.currentText()).strip()
        }

        items = []
        for row in range(self.tbl.rowCount()):
            art = self.tbl.cellWidget(row, 0).currentText()
            card = self.articles.get(art, {})
            items.append({
                "Номенклатура": card.get("name", ""),
                "АртикулГП": card.get("name", ""),
                "ВариантИзготовления": self.tbl.cellWidget(row, 2).currentText(),
                "Размер": self.tbl.cellWidget(row, 3).value(),
                "Количество": self.tbl.cellWidget(row, 4).value(),
                "Вес": self.tbl.cellWidget(row, 5).value(),
                "Примечание": self.tbl.cellWidget(row, 6).text() if self.tbl.cellWidget(row, 6) else "",
                "ЕдиницаИзмерения": "шт"
            })

        success = bridge.update_order(self._current_number, fields, items)
        if success:
            QMessageBox.information(self, "Готово", f"Заказ №{self._current_number} обновлён.")
            self._load_orders()
            self.tabs.setCurrentIndex(1)
        else:
            QMessageBox.critical(self, "Ошибка", f"Не удалось обновить заказ №{self._current_number}")    

    def _ui(self):
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        outer = QVBoxLayout(self)
        self.tabs = QTabWidget()
        outer.addWidget(self.tabs)

        self.frm_new = QWidget()
        v = QVBoxLayout(self.frm_new)
        hdr = QLabel("Заказ в производство")
        hdr.setFont(QFont("Arial", 22, QFont.Bold))
        v.addWidget(hdr)

        form = QFormLayout()
        self.comment_input = QLineEdit()
        form.addRow("Комментарий к заказу:", self.comment_input)
        self.ed_num = QLabel(bridge.get_next_order_number())
        self.d_date = QDateEdit(datetime.now()); self.d_date.setCalendarPopup(True)
        self.c_org = QComboBox(); self.c_org.addItems([x["Description"] for x in self.organizations])
        self.c_contr = QComboBox(); self.c_contr.addItems([x["Description"] for x in self.counterparties])
        self.c_ctr = QComboBox(); self.c_ctr.addItems([x["Description"] for x in self.contracts])
        self.c_wh = QComboBox(); self.c_wh.addItems([x["Description"] for x in self.warehouses])
        self.status_combo = QComboBox(); self.status_combo.addItems(self.production_statuses)

        for label, widget in [
            ("Номер", self.ed_num), ("Дата", self.d_date),
            ("Организация", self.c_org), ("Контрагент", self.c_contr),
            ("Договор", self.c_ctr), ("Склад", self.c_wh),
            ("Вид продукции", self.status_combo)
        ]:
            form.addRow(label, widget)
        v.addLayout(form)

        self.tbl = QTableWidget(0, len(self.COLS))
        self.tbl.setHorizontalHeaderLabels(self.COLS)
        QShortcut(QKeySequence("F9"), self).activated.connect(self._copy_row)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl.verticalHeader().setVisible(False)
        v.addWidget(self.tbl)
        self._add_row()

        btns = QHBoxLayout()
        self.btn_update = QPushButton("🔁 Обновить заказ")
        self.btn_update.clicked.connect(self._update_order)
        self.btn_update.setVisible(False)
        btns.addWidget(self.btn_update)
        for label, func in [
            ("🖨 Печать", self._print_selected_order),
            ("🖨 С фото", self._print_selected_order_with_photo),
            ("+ строка", self._add_row),
            ("− строка", self._remove_row),
            ("Новый заказ", self._new_order),
            ("📋 Копировать строку", self._copy_row),
            ("💾 Записать", self._post_close)
        ]:
            btn = QPushButton(label)
            btn.clicked.connect(func)
            btns.addWidget(btn)
        v.addLayout(btns)
        self.tabs.addTab(self.frm_new, "Новый заказ")

        self.tbl_orders = QTableWidget(0, 10)
        self.tbl_orders.setHorizontalHeaderLabels([
            "✓", "Номер", "Дата", "Вид/статус продукции", "Количество", "Вес",
            "Организация", "Контрагент", "Договор", "Комментарий"
        ])
        self.tbl_orders.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_orders.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_orders.cellDoubleClicked.connect(self._show_order)

        tab_orders = QWidget()
        layout = QVBoxLayout(tab_orders)
        layout.addWidget(QLabel("📋 Заказы в производстве"))

        buttons = QHBoxLayout()
        for label, func in [
            ("🔄 Обновить", self._load_orders),
            ("✅ Провести отмеченные", self._mass_post),
            ("🧷 Пометить", lambda: self._mark_deleted(True)),
            ("📍 Снять пометку", lambda: self._mark_deleted(False)),
            ("🗑 Удалить", self._delete_selected_order)
        ]:
            btn = QPushButton(label); btn.clicked.connect(func); buttons.addWidget(btn)
        layout.addLayout(buttons)
        layout.addWidget(self.tbl_orders)
        self.tabs.addTab(tab_orders, "Заказы")
        
    def _print_selected_order(self):
        selected = self.tbl_orders.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите заказ для печати")
            return
        number = self.tbl_orders.item(selected, 1).text().strip().replace("⚪", "")
        success = bridge.print_order_preview_pdf(number)
        if not success:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сформировать предпросмотр для заказа №{number}")

    def _print_selected_order_with_photo(self):
        selected = self.tbl_orders.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите заказ для печати")
            return
        number = self.tbl_orders.item(selected, 1).text().strip().replace("⚪", "")
        success = bridge.print_order_with_photo(number)
        if not success:
            QMessageBox.critical(self, "Ошибка", f"Не удалось напечатать заказ №{number}")

    def _add_row(self, copy_from: int = None):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        art = QComboBox(); art.setEditable(True); art.addItems(self.articles.keys())
        name = QTableWidgetItem("")
        variant = QComboBox()
        size = QDoubleSpinBox()
        size.setDecimals(1)
        size.setRange(0.5, 50.0)
        size.setSingleStep(0.5)     # шаг изменения
        size.setValue(16.0)
        qty = QSpinBox(); qty.setRange(1, 999); qty.setValue(1)
        wgt = QDoubleSpinBox(); wgt.setDecimals(3); wgt.setMaximum(9999)
        comment = QLineEdit()
        self.tbl.setCellWidget(r, 6, comment)

        widgets = [art, name, variant, size, qty, wgt, comment]
        for c, w in enumerate(widgets):
            if isinstance(w, (QComboBox, QSpinBox, QDoubleSpinBox, QLineEdit)):
                self.tbl.setCellWidget(r, c, w)
            else:
                self.tbl.setItem(r, c, w)

        def fill():
            selected = art.currentText().strip()
            card = self.articles.get(selected, {})
            name.setText(card.get("name", ""))
            if card.get("size"):
                size.setValue(float(card["size"]))
            wgt.setValue(round(card.get("w", 0) * qty.value(), 3))

            # 🟡 добавляем недостающую строку
            article = art.currentText()
            prev = variant.currentText()
            variant.clear()
            variants = bridge.get_variants_by_article(article)
            variant.addItems(variants)
            if prev in variants:
                variant.setCurrentText(prev)

        art.currentTextChanged.connect(fill)
        qty.valueChanged.connect(fill)

        if copy_from is not None:
            for i, (src, dst) in enumerate(zip(range(self.tbl.columnCount()), widgets)):
                if isinstance(dst, QComboBox):
                    current = self.tbl.cellWidget(copy_from, i).currentText()
                    dst.setCurrentText(current)
                elif isinstance(dst, QDoubleSpinBox) or isinstance(dst, QSpinBox):
                    dst.setValue(self.tbl.cellWidget(copy_from, i).value())
                elif isinstance(dst, QTableWidgetItem):
                    dst.setText(self.tbl.item(copy_from, i).text())
        else:
            fill()


    def _copy_last_row(self):
        r = self.tbl.rowCount()
        if r > 0:
            self._add_row(copy_from=r - 1)

    def _remove_row(self):
        r = self.tbl.rowCount()
        if r > 0:
            self.tbl.removeRow(r - 1)

    def _new_order(self):
        self.ed_num.setText(bridge.get_next_order_number())
        self.tbl.setRowCount(0)
        self._add_row()
        self.tabs.setCurrentIndex(0)
        self._edit_mode = False
        self._current_number = ""
        self.btn_update.setVisible(False)  # ← скрыть кнопку "Обновить"

    def _post(self):
        fields = {
            "Организация": self.c_org.currentText(),
            "Контрагент": self.c_contr.currentText(),
            "ДоговорКонтрагента": self.c_ctr.currentText(),
            "Склад": self.c_wh.currentText(),
            "Ответственный": "Администратор",
            "Комментарий": self.comment_input.text().strip(),
            "Дата": pywintypes.Time(self.d_date.date().toPyDate()),
            "ВидСтатусПродукции": str(self.status_combo.currentText()).strip()
        }

        items = []  # ← вот это и нужно было добавить
        for row in range(self.tbl.rowCount()):
            art = self.tbl.cellWidget(row, 0).currentText()
            card = self.articles.get(art, {})
            items.append({
                "Номенклатура": card.get("name", ""),
                "АртикулГП": card.get("name", ""),
                "ВариантИзготовления": self.tbl.cellWidget(row, 2).currentText(),
                "Размер": self.tbl.cellWidget(row, 3).value(),
                "Количество": self.tbl.cellWidget(row, 4).value(),
                "Вес": self.tbl.cellWidget(row, 5).value(),
                "Примечание": self.tbl.cellWidget(row, 6).text() if self.tbl.cellWidget(row, 6) else "",
                "ЕдиницаИзмерения": "шт"
            })

        number = bridge.create_order(fields, items)
        self.ed_num.setText(number)
        self._load_orders()
        QMessageBox.information(self, "Готово", f"Сохранено: заказ №{number}")

    def _post_close(self):
        self._post()
        self.tabs.setCurrentIndex(1)

    def _load_orders(self):
        self.tbl_orders.setRowCount(0)
        self._orders = bridge.list_orders()
        for o in self._orders:
            r = self.tbl_orders.rowCount()
            self.tbl_orders.insertRow(r)
            chk = QTableWidgetItem(); chk.setCheckState(Qt.Unchecked)
            self.tbl_orders.setItem(r, 0, chk)
            status = "🟢" if o.get("posted") else ("❌" if o.get("deleted") else "⚪")
            vals = [
                f"{status} {o['num']}", o["date"], o.get("prod_status", ""),
                o.get("qty", 0), f"{o.get('weight', 0):.3f}",
                o.get("org", ""), o.get("contragent", ""), o.get("contract", ""), o.get("comment", "")
            ]
            for i, v in enumerate(vals):
                self.tbl_orders.setItem(r, i + 1, QTableWidgetItem(str(v)))

    def _mass_post(self):
        for i, o in enumerate(self._orders):
            if self.tbl_orders.item(i, 0).checkState() == Qt.Checked:
                bridge.post_order(o["num"])
        self._load_orders()

    def _delete_selected_order(self):
        selected = [i for i in range(self.tbl_orders.rowCount())
                    if self.tbl_orders.item(i, 0).checkState() == Qt.Checked]
        for i in selected:
            bridge.delete_order_by_number(self._orders[i]["num"])
        self._load_orders()

    def _mark_deleted(self, mark=True):
        for i in range(self.tbl_orders.rowCount()):
            if self.tbl_orders.item(i, 0).checkState() == Qt.Checked:
                number = self._orders[i]["num"]
                if mark:
                    bridge.mark_order_for_deletion(number)
                else:
                    bridge.unmark_order_deletion(number)
        self._load_orders()

    def _show_order(self, row, col):
        o = self._orders[row]
        
        self.comment_input.setText(o.get("comment", ""))

        # Переключаемся на вкладку редактирования
        self.tabs.setCurrentIndex(0)

        # Устанавливаем номер и дату
        self.ed_num.setText(o["num"])
        date = QDate.fromString(o["date"].split()[0], "yyyy-MM-dd")
        self.d_date.setDate(date)

        # Устанавливаем поля верхней формы
        self.c_org.setCurrentText(o["org"])
        self.c_contr.setCurrentText(o["contragent"])
        self.c_ctr.setCurrentText(o["contract"])

        # Вид продукции (преобразование из внутреннего в отображаемое имя)
        for internal in self.production_statuses:
            label = {v: k for k, v in PRODUCTION_STATUS_MAP.items()}.get(internal, internal)
            if label == o["prod_status"]:
                self.status_combo.setCurrentText(internal)
                break

        # Табличная часть
        self.tbl.setRowCount(0)
        for r in o.get("rows", []):
            self._add_row()
            row_index = self.tbl.rowCount() - 1

            # Найти артикул по наименованию
            name = r["nomenclature"]
            article = next((art for art, card in self.articles.items() if card["name"] == name), None)
            if article:
                self.tbl.cellWidget(row_index, 0).setCurrentText(article)

                # Установить вариант изготовления чуть позже, когда заполнится список
                def set_variant(row=row_index, val=r.get("variant", "—")):
                    self.tbl.cellWidget(row, 2).setCurrentText(val)

                QTimer.singleShot(0, set_variant)

            # Размер
            size_val = r.get("size", 0)
            try:
                size_float = float(str(safe_str(size_val)).replace(",", "."))
            except Exception:
                size_float = 0.0
            self.tbl.cellWidget(row_index, 3).setValue(size_float)

            # Кол-во, вес, примечание
            self.tbl.cellWidget(row_index, 4).setValue(int(r.get("qty", 1)))
            self.tbl.cellWidget(row_index, 5).setValue(float(r.get("w", 0)))
            self.tbl.cellWidget(row_index, 6).setText(r.get("note", ""))
            self._edit_mode = True
            self._current_number = o["num"]
            self.btn_update.setVisible(True)  # ← покажем кнопку "Обновить"
        
    def _copy_row(self):
        row = self.tbl.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Нет строки", "Выберите строку для копирования.")
            return
        self._add_row(copy_from=row)
