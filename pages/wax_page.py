# wax_page.py • v0.7
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

        btn_new = QPushButton("Создать наряд")
        btn_new.clicked.connect(self._select_order_for_job)

        btn_ref = QPushButton("Обновить")
        btn_ref.clicked.connect(self.refresh)

        btn_issue = QPushButton("🧾 Выдать")
        btn_issue.clicked.connect(self._give_job)

        btn_done = QPushButton("✅ Сдано")
        btn_done.clicked.connect(self._job_done)

        btn_accept = QPushButton("📥 Принято")
        btn_accept.clicked.connect(self._job_accept)

        btn_task = QPushButton("📋 Задание")
        btn_task.clicked.connect(self._create_task)

        btn_wax_job = QPushButton("📄 Наряд")
        btn_wax_job.clicked.connect(self._create_wax_job)

        btn_sync = QPushButton("🔄 В 1С")
        btn_sync.clicked.connect(self._sync_job)

        for b in [btn_new, btn_ref, btn_task, btn_wax_job, btn_issue, btn_done, btn_accept, btn_sync]:
            btn_row.addWidget(b, alignment=Qt.AlignLeft)

        t1.addLayout(btn_row)

        # — дерево нарядов —
        lab1 = QLabel("Наряды (по методам)")
        lab1.setFont(QFont("Arial", 16, QFont.Bold))
        t1.addWidget(lab1)

        self.tree_jobs = QTreeWidget()
        self.tree_jobs.setHeaderLabels(["Наименование", "Qty", "Вес", "Статус", "1С"])
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

        # ----- Tab 2: Process -----
        tab2 = QWidget(); t2 = QVBoxLayout(tab2)

        self.tree_process = QTreeWidget()
        self.tree_process.setHeaderLabels([
            "Наименование", "Статус", "Исполнитель", "Сдал", "Принял", "Вес, г"
        ])
        self.tree_process.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_process.setStyleSheet(CSS_TREE)
        t2.addWidget(self.tree_process, 1)

        self.tabs.addTab(tab2, "Процесс")
        
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
    def _selected_job_code(self):
        item = self.tree_jobs.currentItem()
        while item and not item.data(0, Qt.UserRole):
            item = item.parent()
        return item.data(0, Qt.UserRole) if item else None

    # ------------------------------------------------------------------
    def _give_job(self):
        from PyQt5.QtWidgets import QInputDialog
        code = self._selected_job_code()
        if not code:
            QMessageBox.warning(self, "Наряд", "Не выбран наряд")
            return
        name, ok = QInputDialog.getText(self, "Исполнитель", "Сотрудник:")
        if ok and name:
            from logic.production_docs import update_wax_job, log_event
            update_wax_job(code, {"assigned_to": name, "status": "given"})
            log_event(code, "given", name)
            self.refresh()

    # ------------------------------------------------------------------
    def _job_done(self):
        from PyQt5.QtWidgets import QInputDialog
        code = self._selected_job_code()
        if not code:
            QMessageBox.warning(self, "Наряд", "Не выбран наряд")
            return
        person, ok = QInputDialog.getText(self, "Выполнил", "Сотрудник:")
        if not (ok and person):
            return
        weight, ok_w = QInputDialog.getDouble(self, "Вес воска, г", "Вес:", 0, 0, 10000, 3)
        if not ok_w:
            weight = None
        from logic.production_docs import update_wax_job, log_event
        update_wax_job(code, {
            "completed_by": person,
            "weight_wax": weight,
            "status": "done"
        })
        log_event(code, "done", person, {"weight_wax": weight})
        self.refresh()

    # ------------------------------------------------------------------
    def _job_accept(self):
        from PyQt5.QtWidgets import QInputDialog
        code = self._selected_job_code()
        if not code:
            QMessageBox.warning(self, "Наряд", "Не выбран наряд")
            return
        name, ok = QInputDialog.getText(self, "Приёмка", "Сотрудник:")
        if ok and name:

            from logic.production_docs import update_wax_job, log_event, get_wax_job
            job = update_wax_job(code, {"accepted_by": name, "status": "accepted"})
            log_event(code, "accepted", name)
            if job:
                num = bridge.create_wax_job(job)
                if num:
                    update_wax_job(code, {"sync_doc_num": num})
            log_event(code, "synced_1c", name, {"doc_num": num})
        self.refresh()

    # ------------------------------------------------------------------
    def _create_task(self):
        from PyQt5.QtWidgets import QInputDialog
        if not ORDERS_POOL:
            QMessageBox.warning(self, "Нет данных", "Нет заказов для создания")
            return
        order_list = [o["docs"]["order_code"] for o in ORDERS_POOL]
        selected, ok = QInputDialog.getItem(self, "Выберите заказ", "Заказ:", order_list, editable=False)
        if ok and selected:
            order = next((o["order"] for o in ORDERS_POOL if o["docs"]["order_code"] == selected), None)
            if order:
                try:
                    num = bridge.create_task_from_order(order)
                    QMessageBox.information(self, "Готово", f"Задание №{num} создано")
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", str(e))

    # ------------------------------------------------------------------
    def _create_wax_job(self):
        from PyQt5.QtWidgets import QInputDialog
        task_num, ok = QInputDialog.getText(self, "Создание наряда", "Номер задания:")
        if ok and task_num:
            try:
                num = bridge.create_wax_job_from_task(task_num)
                QMessageBox.information(self, "Готово", f"Наряд №{num} создан")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", str(e))

    # ------------------------------------------------------------------
    def _sync_job(self):
        code = self._selected_job_code()
        if not code:
            QMessageBox.warning(self, "Наряд", "Не выбран наряд")
            return
        from logic.production_docs import get_wax_job, update_wax_job, log_event
        job = get_wax_job(code)
        if not job:
            QMessageBox.warning(self, "Наряд", "Не найден наряд")
            return
        num = bridge.create_wax_job(job)
        if num:
            update_wax_job(code, {"sync_doc_num": num})
            log_event(code, "synced_1c", None, {"doc_num": num})
        self.refresh()

    # ------------------------------------------------------------------


    def refresh(self):
        self._fill_jobs_tree()
        self._fill_parties_tree()
        self._fill_process_tree()

    # —──────────── дерево «Наряды» ─────────────
    def _fill_jobs_tree(self):
        self.tree_jobs.clear()

        jobs_by_method = defaultdict(list)
        for j in WAX_JOBS_POOL:
            jobs_by_method[j["method"]].append(j)

        for m_key, jobs in jobs_by_method.items():
            root = QTreeWidgetItem(self.tree_jobs, [METHOD_LABEL.get(m_key, m_key), "", "", "", ""])
            root.setExpanded(True)

            for j in jobs:
                item = QTreeWidgetItem(root, [
                    f"{j['operation']} ({j['wax_job']})",
                    str(j.get('qty', 0)),
                    f"{j.get('weight', 0.0):.3f}",
                    j.get('status', ''),
                    '✅' if j.get('sync_doc_num') else ''
                ])
                item.setData(0, Qt.UserRole, j['wax_job'])

    # —──────────── дерево «Партии» ─────────────
    def _fill_parties_tree(self):
        self.tree_part.clear()

        # grouping by партия
        jobs_by_party = defaultdict(list)
        for j in WAX_JOBS_POOL:
            jobs_by_party[j["batch_code"]].append(j)

        for code, jobs in jobs_by_party.items():
            j0 = jobs[0]
            wax_w = sum(j.get('weight_wax') or 0 for j in jobs)
            root = QTreeWidgetItem(self.tree_part, [
                f"Партия {code}  ({j0['metal']} {j0['hallmark']} {j0['color']})",
                str(j0["qty"]), f"{wax_w:.3f}"
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

    # —──────────── дерево «Процесс» ─────────────
    def _fill_process_tree(self):
        self.tree_process.clear()

        by_batch = defaultdict(list)
        for j in WAX_JOBS_POOL:
            by_batch[j["batch_code"]].append(j)

        for code, jobs in by_batch.items():
            j0 = jobs[0]
            root = QTreeWidgetItem(self.tree_process, [
                f"Партия {code} ({j0['metal']} {j0['hallmark']} {j0['color']})"
            ])
            root.setExpanded(True)
            for j in jobs:
                QTreeWidgetItem(root, [
                    f"{j['operation']} ({j['wax_job']})",
                    j.get('status', ''),
                    j.get('assigned_to') or '',
                    j.get('completed_by') or '',
                    j.get('accepted_by') or '',
                    f"{(j.get('weight_wax') or 0):.3f}"
                ])

# ----------------------------------------------------------------------
def _wax_method(article:str)->str:
    """Небольшая обёртка для определения метода по артикулу."""
    art = str(article).lower()
    if "д" in art or "d" in art:
        return "3d"
    return "rubber"
