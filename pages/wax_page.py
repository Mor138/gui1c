# wax_page.py • v0.7
# ─────────────────────────────────────────────────────────────────────────
from collections import defaultdict, Counter
from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QFont
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QPushButton, QMessageBox
)
from logic.production_docs import WAX_JOBS_POOL, ORDERS_POOL, METHOD_LABEL

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

    # ------------------------------------------------------------------
    def _ui(self):
        v = QVBoxLayout(self); v.setContentsMargins(40,30,40,30)

        hdr = QLabel("Воскование / 3-D печать")
        hdr.setFont(QFont("Arial",22,QFont.Bold)); v.addWidget(hdr)

        btn_row = QVBoxLayout()
        btn_new = QPushButton("Создать наряд")
        btn_new.clicked.connect(self._select_order_for_job)
        btn_ref = QPushButton("Обновить")
        btn_ref.clicked.connect(self.refresh)
        btn_row.addWidget(btn_new, alignment=Qt.AlignLeft)
        btn_row.addWidget(btn_ref, alignment=Qt.AlignLeft)
        v.addLayout(btn_row)

        # — дерево нарядов —
        lab1 = QLabel("Наряды (по методам)")
        lab1.setFont(QFont("Arial",16,QFont.Bold)); v.addWidget(lab1)

        self.tree_jobs = QTreeWidget()
        self.tree_jobs.setHeaderLabels(["Наименование","Qty","Вес"])
        self.tree_jobs.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_jobs.setStyleSheet(CSS_TREE)
        v.addWidget(self.tree_jobs,1)

        # — дерево партий —
        lab2 = QLabel("Партии (металл / проба / цвет)")
        lab2.setFont(QFont("Arial",16,QFont.Bold)); v.addWidget(lab2)

        self.tree_part = QTreeWidget()
        self.tree_part.setHeaderLabels(["Наименование","Qty","Вес"])
        self.tree_part.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_part.setStyleSheet(CSS_TREE)
        v.addWidget(self.tree_part,1)
        
    def _select_order_for_job(self):
        from PyQt5.QtWidgets import QInputDialog

        if not ORDERS_POOL:
            QMessageBox.warning(self, "Нет заказов", "Нет заказов для обработки")
            return

        # Получаем список заказов по номеру
        order_list = [o["docs"]["order_code"] for o in ORDERS_POOL]
        selected, ok = QInputDialog.getItem(self, "Выберите заказ", "Заказ:", order_list, editable=False)

        if ok and selected:
            selected_order = next((o for o in ORDERS_POOL if o["docs"]["order_code"] == selected), None)
            if selected_order:
                from logic.production_docs import WAX_JOBS_POOL, build_wax_jobs
                jobs = build_wax_jobs(selected_order["order"], selected_order["docs"]["batches"])
                WAX_JOBS_POOL.extend(jobs)
                QMessageBox.information(self, "Готово", f"Создано {len(jobs)} нарядов")
                self.refresh()
            else:
                QMessageBox.critical(self, "Ошибка", "Не удалось найти заказ")    

    # ------------------------------------------------------------------
    def _stub_create_job(self):
        QMessageBox.information(self,"Создать наряд",
            "Диалог выбора заказов появится на следующем этапе 🙂")

    # ------------------------------------------------------------------
    def refresh(self):
        self._fill_jobs_tree()
        self._fill_parties_tree()

    # —──────────── дерево «Наряды» ─────────────
    def _fill_jobs_tree(self):
        self.tree_jobs.clear()

        # метод → {article: Counter(size→qty, weight)}
        data = defaultdict(lambda: defaultdict(list))
        for pack in ORDERS_POOL:
            for row in pack["order"]["rows"]:
                method = _wax_method(row["article"])
                data[method][row["article"]].append(row)

        for m_key, arts in data.items():
            root = QTreeWidgetItem(self.tree_jobs,
                [f"Наряд {METHOD_LABEL[m_key]}", "", ""])
            root.setExpanded(True)

            for art, rows in arts.items():
                # агрегируем по размеру
                by_size = Counter()
                weight_sum = 0
                for r in rows:
                    by_size[r["size"]] += r["qty"]
                    weight_sum += r["weight"]
                qty_sum = sum(by_size.values())
                art_node = QTreeWidgetItem(root, [art, str(qty_sum), f"{weight_sum:.3f}"])
                for size, q in sorted(by_size.items()):
                    QTreeWidgetItem(art_node,
                        [f"р-р {size}", str(q), ""])

    # —──────────── дерево «Партии» ─────────────
    def _fill_parties_tree(self):
        self.tree_part.clear()

        # grouping by партия
        jobs_by_party = defaultdict(list)
        for j in WAX_JOBS_POOL:
            jobs_by_party[j["batch_code"]].append(j)

        for code, jobs in jobs_by_party.items():
            j0 = jobs[0]
            root = QTreeWidgetItem(self.tree_part, [
                f"Партия {code}  ({j0['metal']} {j0['hallmark']} {j0['color']})",
                str(j0["qty"]), f"{j0['weight']:.3f}"
            ])
            root.setExpanded(True)

            # article+size aggregated
            agg = defaultdict(lambda: dict(qty=0, weight=0))
            for pack in ORDERS_POOL:
                for row in pack["order"]["rows"]:
                    if (row["metal"],row["hallmark"],row["color"])==(
                        j0["metal"],j0["hallmark"],j0["color"]):
                        k = (row["article"], row["size"])
                        agg[k]["qty"] += row["qty"]
                        agg[k]["weight"] += row["weight"]

            for (art,size), d in agg.items():
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
