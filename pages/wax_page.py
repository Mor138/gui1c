# wax_page.py • v0.7
# ─────────────────────────────────────────────────────────────────────────
from collections import defaultdict
from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QFont
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QPushButton, QMessageBox
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
        v = QVBoxLayout(self); v.setContentsMargins(40,30,40,30)

        hdr = QLabel("Воскование / 3-D печать")
        hdr.setFont(QFont("Arial",22,QFont.Bold)); v.addWidget(hdr)

        btn_row = QHBoxLayout()
        btn_new = QPushButton("Создать наряд")
        btn_new.clicked.connect(self._select_order_for_job)
        btn_ref = QPushButton("Обновить")
        btn_ref.clicked.connect(self.refresh)
        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
        btn_issue = QPushButton("Выдать")
        btn_issue.clicked.connect(self._give_job)
        btn_done = QPushButton("Сдано")
        btn_done.clicked.connect(self._job_done)
        btn_accept = QPushButton("Принято")
        btn_accept.clicked.connect(self._job_accept)
        btn_sync = QPushButton("В 1С")
        btn_sync.clicked.connect(self._sync_job)
        for b in [btn_new, btn_ref, btn_issue, btn_done, btn_accept, btn_sync]:
=======
        btn_issue = QPushButton("🧾 Выдать")
        btn_issue.clicked.connect(self._give_job)
        btn_done = QPushButton("✅ Сдано")
        btn_done.clicked.connect(self._job_done)
        btn_accept = QPushButton("📥 Принято")
        btn_accept.clicked.connect(self._job_accept)
        ep5fca-codex/реализация-логики-для-управления-нарядами-и-партиями
        btn_sync = QPushButton("🔄 В 1С")
        btn_sync.clicked.connect(self._sync_job)
        for b in [btn_new, btn_ref, btn_issue, btn_done, btn_accept, btn_sync]:
=======
        for b in [btn_new, btn_ref, btn_issue, btn_done, btn_accept]:
        main
        main
            btn_row.addWidget(b, alignment=Qt.AlignLeft)
        v.addLayout(btn_row)

        # — дерево нарядов —
        lab1 = QLabel("Наряды (по методам)")
        lab1.setFont(QFont("Arial",16,QFont.Bold)); v.addWidget(lab1)

        self.tree_jobs = QTreeWidget()
        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
        self.tree_jobs.setHeaderLabels(["Наименование","Qty","Вес","Статус","1С"])
=======
        ep5fca-codex/реализация-логики-для-управления-нарядами-и-партиями
        self.tree_jobs.setHeaderLabels(["Наименование","Qty","Вес","Статус","1С"])
=======
        self.tree_jobs.setHeaderLabels(["Наименование","Qty","Вес","Статус"])
        main
        main
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
        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
=======
        ep5fca-codex/реализация-логики-для-управления-нарядами-и-партиями
        main
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
        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
=======
=======
            from logic.production_docs import update_wax_job, log_event
            update_wax_job(code, {"accepted_by": name, "status": "accepted"})
            log_event(code, "accepted", name)
            self.refresh()

    # ------------------------------------------------------------------
        main
        main
    def refresh(self):
        self._fill_jobs_tree()
        self._fill_parties_tree()

    # —──────────── дерево «Наряды» ─────────────
    def _fill_jobs_tree(self):
        self.tree_jobs.clear()

        jobs_by_method = defaultdict(list)
        for j in WAX_JOBS_POOL:
            jobs_by_method[j["method"]].append(j)

        for m_key, jobs in jobs_by_method.items():
            root = QTreeWidgetItem(self.tree_jobs,
        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
                                  [METHOD_LABEL.get(m_key, m_key), "", "", "", ""])
=======
        ep5fca-codex/реализация-логики-для-управления-нарядами-и-партиями
                                  [METHOD_LABEL.get(m_key, m_key), "", "", "", ""])
=======
                                  [METHOD_LABEL.get(m_key, m_key), "", "", ""])
        main
        main
            root.setExpanded(True)

            for j in jobs:
                item = QTreeWidgetItem(root, [
                    f"{j['operation']} ({j['wax_job']})",
                    str(j['qty']),
                    f"{j['weight']:.3f}",
        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
                    j.get('status', ''),
                    ('OK' if j.get('sync_doc_num') else '')
=======
        ep5fca-codex/реализация-логики-для-управления-нарядами-и-партиями
                    j.get('status', ''),
                    '✅' if j.get('sync_doc_num') else ''
=======
                    j.get('status', '')
        main
        main
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

# ----------------------------------------------------------------------
def _wax_method(article:str)->str:
    """Небольшая обёртка для определения метода по артикулу."""
    art = str(article).lower()
    if "д" in art or "d" in art:
        return "3d"
    return "rubber"
