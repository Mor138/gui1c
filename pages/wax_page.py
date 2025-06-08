# wax_page.py • v0.8
# ─────────────────────────────────────────────────────────────────────────
from collections import defaultdict
from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QFont
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QPushButton, QMessageBox, QTabWidget
)
from logic.production_docs import WAX_JOBS_POOL, ORDERS_POOL, METHOD_LABEL
from core.com_bridge import COM1CBridge

bridge = COM1CBridge("C:\\Users\\Mor\\Desktop\\1C\\proiz")

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

class WaxPage(QWidget):
    def __init__(self):
        super().__init__()
        self._ui()
        self.refresh()
        
    def refresh(self):
        self._fill_jobs_tree()
        self._fill_parties_tree()
        self._fill_tasks_tree()
        self._fill_wax_jobs_tree()    
        self.tree_acts.itemDoubleClicked.connect(self._on_wax_job_double_click)
        self.tree_tasks.itemDoubleClicked.connect(self._on_task_double_click)

    # ------------------------------------------------------------------
    def _ui(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(40, 30, 40, 30)

        hdr = QLabel("Воскование / 3-D печать")
        hdr.setFont(QFont("Arial", 22, QFont.Bold))
        v.addWidget(hdr)

        self.tabs = QTabWidget()
        v.addWidget(self.tabs, 1)

        # ----- Tab 1: Jobs and Batches -----
        tab1 = QWidget(); t1 = QVBoxLayout(tab1)

        btn_row = QHBoxLayout()

        btn_create_task = QPushButton("📋 Создать задание")
        btn_create_task.clicked.connect(self._create_task)

        btn_create_wax_jobs = QPushButton("📄 Создать наряды")
        btn_create_wax_jobs.clicked.connect(self._create_wax_jobs)

        for b in [btn_create_task, btn_create_wax_jobs]:
            btn_row.addWidget(b, alignment=Qt.AlignLeft)

        t1.addLayout(btn_row)

        # — дерево нарядов —
        lab1 = QLabel("Наряды (по методам)")
        lab1.setFont(QFont("Arial", 16, QFont.Bold))
        t1.addWidget(lab1)

        self.tree_jobs = QTreeWidget()
        self.tree_jobs.setHeaderLabels([
            "Артикулы", "Метод", "Qty", "Вес", "Статус", "1С"
        ])
        self.tree_jobs.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_jobs.setStyleSheet(CSS_TREE)
        t1.addWidget(self.tree_jobs, 1)

        # — дерево партий —
        lab2 = QLabel("Партии (металл / проба / цвет)")
        lab2.setFont(QFont("Arial", 16, QFont.Bold))
        t1.addWidget(lab2)

        self.tree_part = QTreeWidget()
        self.tree_part.setHeaderLabels(["Наименование", "Qty", "Вес"])
        self.tree_part.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_part.setStyleSheet(CSS_TREE)
        t1.addWidget(self.tree_part, 1)

        self.tabs.addTab(tab1, "Наряды")
        
        # ----- Tab 2: Задания -----
        tab2 = QWidget(); t2 = QVBoxLayout(tab2)
        lbl2 = QLabel("Задания на производство")
        lbl2.setFont(QFont("Arial", 16, QFont.Bold))
        t2.addWidget(lbl2)

        self.tree_tasks = QTreeWidget()
        self.tree_tasks.setHeaderLabels(["Номер", "Дата", "Участок", "Операция", "Ответственный"])
        self.tree_tasks.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_tasks.setStyleSheet(CSS_TREE)
        t2.addWidget(self.tree_tasks, 1)

        self.tabs.addTab(tab2, "Задания")

        # ----- Tab 3: Наряды -----
        tab3 = QWidget(); t3 = QVBoxLayout(tab3)
        lbl3 = QLabel("Наряды на восковку")
        lbl3.setFont(QFont("Arial", 16, QFont.Bold))
        t3.addWidget(lbl3)

        self.tree_acts = QTreeWidget()
        self.tree_acts.setHeaderLabels(["Номер", "Дата", "Орг.", "Склад", "Участок", "Сотрудник", "Операция", "Статус", "Основание"])
        self.tree_acts.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_acts.setStyleSheet(CSS_TREE)
        t3.addWidget(self.tree_acts, 1)

        self.tabs.addTab(tab3, "Наряды из 1С")
        
    def _show_wax_job_detail(self, item):
        from PyQt5.QtWidgets import QDialog, QTableWidget, QTableWidgetItem, QVBoxLayout
        num = item.text(0)
        rows = bridge.get_wax_job_rows(num)

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Строки наряда {num}")
        dlg.setMinimumWidth(1000)
        layout = QVBoxLayout(dlg)

        tbl = QTableWidget()
        tbl.setRowCount(len(rows))
        tbl.setColumnCount(9)
        tbl.setHorizontalHeaderLabels([
            "Номенклатура", "Размер", "Проба", "Цвет", "Кол-во", "Вес", "Партия", "Ёлка", "Состав набора"
        ])
        for i, r in enumerate(rows):
            for j, k in enumerate([
                "Номенклатура", "Размер", "Проба", "Цвет", "Количество", "Вес", "Партия", "Номер ёлки", "Состав набора"
            ]):
                tbl.setItem(i, j, QTableWidgetItem(str(r.get(k, ""))))

        tbl.resizeColumnsToContents()
        layout.addWidget(tbl)
        dlg.setLayout(layout)
        dlg.exec_()    
        
    def _on_task_double_click(self, item, column):
        num = item.text(0).strip()
        if not num:
            return

        lines = bridge.get_task_lines(num)
        if not lines:
            QMessageBox.information(self, "Нет данных", f"В задании {num} нет строк")
            return

        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTreeWidget, QTreeWidgetItem

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Продукция по заданию {num}")
        layout = QVBoxLayout(dlg)

        tree = QTreeWidget()
        tree.setHeaderLabels(["Номенклатура", "Размер", "Проба", "Цвет", "Кол-во", "Вес"])
        for row in lines:
            QTreeWidgetItem(tree, [
                row["nomen"], str(row["size"]), str(row["sample"]),
                str(row["color"]), str(row["qty"]), str(row["weight"])
            ])
        layout.addWidget(tree)
        dlg.resize(700, 400)
        dlg.exec_()    
        
    def _on_wax_job_double_click(self, item, column):
        num = item.text(0).strip()
        if not num:
            return

        lines = bridge.get_wax_job_lines(num)
        if not lines:
            QMessageBox.information(self, "Нет данных", f"В наряде {num} нет строк")
            return

        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTreeWidget, QTreeWidgetItem

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Строки наряда {num}")
        layout = QVBoxLayout(dlg)

        tree = QTreeWidget()
        tree.setHeaderLabels(["Номенклатура", "Размер", "Проба", "Цвет", "Кол-во", "Вес"])
        for row in lines:
            QTreeWidgetItem(tree, [
                row["nomen"], str(row["size"]), str(row["sample"]),
                str(row["color"]), str(row["qty"]), str(row["weight"])
            ])
        layout.addWidget(tree)
        dlg.resize(700, 400)
        dlg.exec_()    
        
    def _fill_tasks_tree(self):
        self.tree_tasks.clear()
        for t in bridge.list_tasks():
            QTreeWidgetItem(self.tree_tasks, [
                t.get("num", ""),
                t.get("date", ""),
                t.get("employee", ""),
                t.get("tech_op", ""),
                ""
            ])

    def _fill_wax_jobs_tree(self):
        self.tree_acts.clear()
        for t in bridge.list_wax_jobs():
            QTreeWidgetItem(self.tree_acts, [
                t.get("num", ""),
                t.get("date", ""),
                t.get("organization", ""),
                t.get("warehouse", ""),
                t.get("section", ""),
                t.get("employee", ""),
                t.get("tech_op", ""),
                t.get("status", ""),
                t.get("based_on", "")
            ])    
        

    def _create_task(self):
        if not ORDERS_POOL:
            QMessageBox.warning(self, "Нет данных", "Нет заказов для создания")
            return

        for o in ORDERS_POOL:
            order = o.get("order", {})
            order_ref = order.get("Ref")
            if not order_ref:
                QMessageBox.warning(self, "Ошибка", "У заказа нет ссылки Ref для создания задания")
                continue

            method_to_items = defaultdict(list)

            try:
                num = bridge.create_task_from_order(order)
                QMessageBox.information(self, "Готово", f"Задание №{num} создано")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", str(e))

            for row in o["order"]["rows"]:
                method = _wax_method(row["article"])
                method_to_items[method].append({
                    "name": row["article"],
                    "qty": row["qty"],
                    "size": row.get("size", ""),
                    "assay": row.get("hallmark", ""),
                    "color": row.get("color", ""),
                    "insert": row.get("insert", "")
                })

            for method, items in method_to_items.items():
                try:
                    result = bridge.create_production_task(order_ref, method, items)
                    QMessageBox.information(self, "Готово", f"Задание №{result['Номер']} создано")
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка создания задания", str(e))
        
        self.refresh()

    # ------------------------------------------------------------------
    def _create_wax_jobs(self):
        if not ORDERS_POOL:
            QMessageBox.warning(self, "Нет данных", "Нет партий для создания нарядов")
            return

        created = 0

        for pack in ORDERS_POOL:
            for batch in pack["docs"].get("batches", []):
                items = []
                for row in pack["order"]["rows"]:
                    if (row["metal"], row["hallmark"], row["color"]) == (
                        batch["metal"], batch["hallmark"], batch["color"]
                    ):
                        items.append({
                            "name": row["article"],
                            "qty": row["qty"],
                            "size": row.get("size", ""),
                            "assay": row.get("hallmark", ""),
                            "color": row.get("color", ""),
                            "insert": row.get("insert", ""),
                            "batch": batch["batch_barcode"]
                        })

                method = _wax_method(items[0]["name"]) if items else "3d"

                # получаем ссылку на задание по методу
                task_ref = bridge.find_production_task_ref_by_method(method)
                if not task_ref:
                    QMessageBox.warning(self, "Не найдено задание", f"Нет задания для метода {method}")
                    continue

                try:
                    result = bridge.create_wax_order_from_task(task_ref, items)
                    created += 1
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка создания наряда", str(e))

        QMessageBox.information(self, "Готово", f"Создано {created} нарядов")
        self.refresh()

    # ------------------------------------------------------------------
    def _sync_job(self):
        code = self._selected_job_code()
        if not code:
            QMessageBox.warning(self, "Наряд", "Не выбран наряд")

        task_num, ok = QInputDialog.getText(self, "Создание нарядов", "Номер задания:")
        if not ok or not task_num:

            return
        try:
            count = bridge.create_multiple_wax_jobs_from_task(task_num)
            QMessageBox.information(self, "Готово", f"Создано {count} нарядов")
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


    # —──────────── дерево «Наряды» ─────────────
    def _fill_jobs_tree(self):
        self.tree_jobs.clear()

        # Группируем по wax_job (один наряд = одна запись)
        grouped = defaultdict(list)
        for job in WAX_JOBS_POOL:
            grouped[job["wax_job"]].append(job)

        for wax_code, rows in grouped.items():
            j0 = rows[0]  # первая строка — для заголовка
            method_label = METHOD_LABEL.get(j0["method"], j0["method"])
            total_qty = sum(r["qty"] for r in rows)
            total_weight = sum(r["weight"] for r in rows)

            # верхний уровень — сам наряд
            root = QTreeWidgetItem(self.tree_jobs, [
                f"{method_label} ({wax_code})",
                method_label,
                str(total_qty),
                f"{total_weight:.3f}",
                j0.get("status", ""),
                '✅' if j0.get('sync_doc_num') else ''
            ])
            root.setExpanded(True)

            # дочерние элементы — артикула с количеством
            for r in rows:
                QTreeWidgetItem(root, [
                    r["articles"],
                    "",
                    str(r["qty"]),
                    f"{r.get('weight', 0.0):.3f}",
                    "", ""
                ])

    # —──────────── дерево «Партии» ─────────────
    def _fill_parties_tree(self):
        self.tree_part.clear()

        # Информацию о партиях берём из ORDERS_POOL
        for pack in ORDERS_POOL:
            for b in pack["docs"].get("batches", []):
                root = QTreeWidgetItem(self.tree_part, [
                    f"Партия {b['batch_barcode']}  ({b['metal']} {b['hallmark']} {b['color']})",
                    str(b["qty"]), f"{b['total_w']:.3f}"
                ])
                root.setExpanded(True)

                agg = defaultdict(lambda: dict(qty=0, weight=0))
                for row in pack["order"]["rows"]:
                    if (row["metal"], row["hallmark"], row["color"]) == (
                        b["metal"], b["hallmark"], b["color"]):
                        k = (row["article"], row["size"])
                        agg[k]["qty"] += row["qty"]
                        agg[k]["weight"] += row["weight"]

                for (art, size), d in agg.items():
                    QTreeWidgetItem(root, [
                        f"{art}  (р-р {size})",
                        str(d["qty"]), f"{d['weight']:.3f}"
                    ])



# ----------------------------------------------------------------------
def _wax_method(article:str)->str:
    """Небольшая обёртка для определения метода по артикулу."""
    art = str(article).lower()
    if "д" in art or "d" in art:
        return "3d"
    return "rubber"
