import uuid
from datetime import datetime
from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QFormLayout, QComboBox, QDateEdit,
    QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QPushButton,
    QMessageBox, QHeaderView, QCheckBox, QTabWidget, QHBoxLayout, QLineEdit
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from core.com_bridge import COM1CBridge

bridge = COM1CBridge("C:\\Users\\Mor\\Desktop\\1C\\proiz")

class OrdersPage(QWidget):
    COLS = ["Артикул", "Наим.", "Металл", "Проба", "Цвет", "Вставки",
            "Размер", "Кол-во", "Вес, г", "Комментарий"]

    def __init__(self):
        super().__init__()
        self.articles = bridge.get_articles()
        self.variants = bridge.get_manufacturing_options()
        self.organizations = bridge.list_catalog_items("Организации")
        self.counterparties = bridge.list_catalog_items("Контрагенты")
        self.contracts = bridge.list_catalog_items("ДоговорыКонтрагентов")
        self.warehouses = bridge.list_catalog_items("Склады")

        self.metals = sorted({v["metal"] for v in self.variants.values()})
        self.hallmarks = sorted({v["hallmark"] for v in self.variants.values()})
        self.colors = sorted({v["color"] for v in self.variants.values()})
        self.inserts = sorted({a["insert"] for a in self.articles.values() if a.get("insert")})

        self._ui()
        self._load_orders()

    def _ui(self):
        outer = QVBoxLayout(self)
        self.tabs = QTabWidget()
        outer.addWidget(self.tabs)

        self.frm_new = QWidget()
        v = QVBoxLayout(self.frm_new)
        v.setContentsMargins(40, 30, 40, 30)

        hdr = QLabel("Заказ в производство")
        hdr.setFont(QFont("Arial", 22, QFont.Bold))
        v.addWidget(hdr)

        form = QFormLayout()
        v.addLayout(form)

        self.ed_num = QLabel(datetime.now().strftime("%Y%m%d-") + uuid.uuid4().hex[:4])
        self.d_date = QDateEdit(datetime.now()); self.d_date.setCalendarPopup(True)

        self.c_org = QComboBox(); self.c_org.addItems([x.get("Description", "—") for x in self.organizations])
        self.c_contr = QComboBox(); self.c_contr.addItems([x.get("Description", "—") for x in self.counterparties])
        self.c_ctr = QComboBox(); self.c_ctr.addItems([x.get("Description", "Купля-продажа") for x in self.contracts])
        self.c_wh = QComboBox(); self.c_wh.addItems([x.get("Description", "Основной") for x in self.warehouses])

        self.chk_ok = QCheckBox("Согласовано")
        self.chk_res = QCheckBox("Резервировать товары")

        for lab, w in [
            ("Номер", self.ed_num), ("Дата", self.d_date),
            ("Организация", self.c_org), ("Контрагент", self.c_contr),
            ("Договор", self.c_ctr), ("Склад", self.c_wh)
        ]:
            form.addRow(lab, w)
        form.addRow(self.chk_ok)
        form.addRow(self.chk_res)

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
        metal = QComboBox(); metal.addItems(self.metals or ["Золото", "Серебро"])
        probe = QComboBox(); probe.addItems(self.hallmarks or ["585", "750"])
        color = QComboBox(); color.addItems(self.colors or ["Красный", "Белый", "Желтый"])
        ins = QComboBox(); ins.addItems(self.inserts or ["Фианит", "Нет"])
        size = QDoubleSpinBox(); size.setDecimals(1); size.setRange(0.5, 50.0)
        qty = QSpinBox(); qty.setRange(1, 999); qty.setValue(1)
        wgt = QDoubleSpinBox(); wgt.setDecimals(3); wgt.setMaximum(9999)
        cmnt = QLineEdit()

        widgets = [art, name, metal, probe, color, ins, size, qty, wgt, cmnt]
        for c, w in enumerate(widgets):
            if isinstance(w, (QComboBox, QSpinBox, QDoubleSpinBox, QLineEdit)):
                self.tbl.setCellWidget(r, c, w)
            else:
                self.tbl.setItem(r, c, w)

        def fill():
            card = self.articles.get(art.currentText(), {})
            variant = self.variants.get(art.currentText(), {})
            name.setText(card.get("name", ""))
            metal.setCurrentText(variant.get("metal", ""))
            probe.setCurrentText(variant.get("hallmark", ""))
            color.setCurrentText(variant.get("color", ""))
            wgt.setValue(round(card.get("w", 0) * qty.value(), 3))
            if card.get("size"): size.setValue(float(card["size"]))
            ins.setCurrentText(card.get("insert", ""))

        art.currentTextChanged.connect(fill)
        qty.valueChanged.connect(fill)
        fill()

    def _remove_row(self):
        if self.tbl.rowCount() > 0:
            self.tbl.removeRow(self.tbl.rowCount() - 1)

    def _new_order(self):
        self.ed_num.setText(datetime.now().strftime("%Y%m%d-") + uuid.uuid4().hex[:4])
        self.tbl.setRowCount(0)
        self._add_row()
        self.tabs.setCurrentIndex(0)

    def _post(self):
        QMessageBox.information(self, "Готово", "Проведено!")

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
            f"{r['article']}  [{r['metal']} {r['hallmark']} {r['color']}]  x{r['qty']}  {r['w']} г — {r['comment']}"
            for r in rows
        ]) or "(нет строк)"
        dlg.setText(txt)
        dlg.exec_()
