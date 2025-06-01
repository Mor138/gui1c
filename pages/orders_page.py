import pywintypes
from datetime import datetime
from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QFormLayout, QComboBox, QDateEdit,
    QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QPushButton,
    QMessageBox, QHeaderView, QTabWidget, QHBoxLayout
)
from PyQt5.QtGui import QFont
from core.com_bridge import COM1CBridge

bridge = COM1CBridge("C:\\Users\\Mor\\Desktop\\1C\\proiz")

class OrdersPage(QWidget):
    COLS = ["Артикул", "Наим.", "Вариант", "Размер", "Кол-во", "Вес, г"]

    def __init__(self):
        super().__init__()
        self.articles = bridge.get_articles()
        self.organizations = bridge.list_catalog_items("Организации")
        self.counterparties = bridge.list_catalog_items("Контрагенты")
        self.contracts = bridge.list_catalog_items("ДоговорыКонтрагентов")
        self.warehouses = bridge.list_catalog_items("Склады")
        self.production_statuses = bridge.get_production_status_variants()
        self._ui()
        self._load_orders()

    def _ui(self):
        outer = QVBoxLayout(self)
        self.tabs = QTabWidget()
        outer.addWidget(self.tabs)

        self.frm_new = QWidget()
        v = QVBoxLayout(self.frm_new)

        hdr = QLabel("Заказ в производство")
        hdr.setFont(QFont("Arial", 22, QFont.Bold))
        v.addWidget(hdr)

        form = QFormLayout()
        self.ed_num = QLabel(bridge.get_next_order_number())
        self.d_date = QDateEdit(datetime.now()); self.d_date.setCalendarPopup(True)
        self.c_org = QComboBox(); self.c_org.addItems([x["Description"] for x in self.organizations])
        self.c_contr = QComboBox(); self.c_contr.addItems([x["Description"] for x in self.counterparties])
        self.c_ctr = QComboBox(); self.c_ctr.addItems([x["Description"] for x in self.contracts])
        self.c_wh = QComboBox(); self.c_wh.addItems([x["Description"] for x in self.warehouses])
        self.status_combo = QComboBox(); self.status_combo.addItems(self.production_statuses)

        for lab, w in [
            ("Номер", self.ed_num), ("Дата", self.d_date),
            ("Организация", self.c_org), ("Контрагент", self.c_contr),
            ("Договор", self.c_ctr), ("Склад", self.c_wh),
            ("Вид продукции", self.status_combo)
        ]:
            form.addRow(lab, w)
        v.addLayout(form)

        self.tbl = QTableWidget(0, len(self.COLS))
        self.tbl.setHorizontalHeaderLabels(self.COLS)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl.verticalHeader().setVisible(False)
        v.addWidget(self.tbl)
        self._add_row()

        btns = QHBoxLayout()
        for txt, slot in [
            ("+ строка", self._add_row),
            ("− строка", self._remove_row),
            ("Новый заказ", self._new_order),
            ("Провести", self._post),
            ("Провести и закрыть", self._post_close)
        ]:
            b = QPushButton(txt); b.clicked.connect(slot)
            btns.addWidget(b)
        v.addLayout(btns)

        self.tabs.addTab(self.frm_new, "Новый заказ")

        self.tbl_orders = QTableWidget(0, 3)
        self.tbl_orders.setHorizontalHeaderLabels(["№", "Дата", "Контрагент"])
        self.tbl_orders.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_orders.cellDoubleClicked.connect(self._show_order)

        tab_orders = QWidget()
        layout_orders = QVBoxLayout(tab_orders)
        btn_reload = QPushButton("🔄 Обновить список заказов")
        btn_reload.clicked.connect(self._load_orders)
        layout_orders.addWidget(btn_reload)
        layout_orders.addWidget(self.tbl_orders)
        self.tabs.addTab(tab_orders, "Заказы")

    def _add_row(self):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        art = QComboBox(); art.setEditable(True); art.addItems(self.articles.keys())
        name = QTableWidgetItem("")
        variant = QComboBox()
        size = QDoubleSpinBox(); size.setDecimals(1); size.setRange(0.5, 50.0)
        qty = QSpinBox(); qty.setRange(1, 999); qty.setValue(1)
        wgt = QDoubleSpinBox(); wgt.setDecimals(3); wgt.setMaximum(9999)

        widgets = [art, name, variant, size, qty, wgt]
        for c, w in enumerate(widgets):
            if isinstance(w, (QComboBox, QSpinBox, QDoubleSpinBox)):
                self.tbl.setCellWidget(r, c, w)
            else:
                self.tbl.setItem(r, c, w)

        def fill():
            selected_art = art.currentText().strip()
            card = self.articles.get(selected_art, {})
            name.setText(card.get("name", ""))
            if card.get("size"):
                size.setValue(float(card["size"]))
            wgt.setValue(round(card.get("w", 0) * qty.value(), 3))

            # Получаем варианты динамически по артикулу
            variant.clear()
            if selected_art:
                filtered = bridge.get_variants_by_article(selected_art)
                variant.addItems(filtered or ["—"])
            else:
                variant.addItem("—")

        art.currentTextChanged.connect(fill)
        qty.valueChanged.connect(fill)
        fill()

    def _remove_row(self):
        if self.tbl.rowCount() > 0:
            self.tbl.removeRow(self.tbl.rowCount() - 1)

    def _new_order(self):
        self.ed_num.setText(bridge.get_next_order_number())
        self.tbl.setRowCount(0)
        self._add_row()
        self.tabs.setCurrentIndex(0)

    def _post(self):
        fields = {
            "Организация": self.c_org.currentText(),
            "Контрагент": self.c_contr.currentText(),
            "ДоговорКонтрагента": self.c_ctr.currentText(),
            "Склад": self.c_wh.currentText(),
            "Ответственный": "Администратор",
            "Комментарий": f"Создан через GUI {datetime.now().strftime('%d.%m.%Y %H:%M')}",
            "Дата": pywintypes.Time(self.d_date.date().toPyDate()),
            "ВидСтатусПродукции": self.status_combo.currentText()
        }

        items = []
        for row in range(self.tbl.rowCount()):
            art = self.tbl.cellWidget(row, 0).currentText()
            card = self.articles.get(art, {})
            variant = self.tbl.cellWidget(row, 2).currentText()
            size = self.tbl.cellWidget(row, 3).value()
            qty = self.tbl.cellWidget(row, 4).value()
            wgt = self.tbl.cellWidget(row, 5).value()

            items.append({
                "Номенклатура": card.get("name", ""),
                "АртикулГП": card.get("name", ""),
                "ВариантИзготовления": variant,
                "Размер": size,
                "Количество": qty,
                "Вес": wgt,
                "ЕдиницаИзмерения": "шт"
            })

        number = bridge.create_order(fields, items)
        self.ed_num.setText(number)
        self._load_orders()
        QMessageBox.information(self, "Готово", f"Проведено: заказ №{number}")

    def _post_close(self):
        self._post()
        self.tabs.setCurrentIndex(1)

    def _load_orders(self):
        self.tbl_orders.setRowCount(0)
        self._orders = bridge.list_orders()
        for o in self._orders:
            r = self.tbl_orders.rowCount()
            self.tbl_orders.insertRow(r)
            for c, val in enumerate([o["num"], o["date"], o["contragent"]]):
                self.tbl_orders.setItem(r, c, QTableWidgetItem(str(val)))

    def _show_order(self, row, col):
        if row >= len(self._orders): return
        o = self._orders[row]
        dlg = QMessageBox()
        dlg.setWindowTitle(f"Заказ №{o['num']} от {o['date']}")
        rows = o.get("rows", [])
        txt = "\n".join([
            f"{r['nomenclature']} [{r['variant']}] {r['qty']}шт {r['w']}г"
            for r in rows
        ]) or "(нет строк)"
        dlg.setText(txt)
        dlg.exec_()
