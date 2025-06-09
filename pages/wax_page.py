# wax_page.py • v0.8
# ─────────────────────────────────────────────────────────────────────────
from collections import defaultdict
from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QFont
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QPushButton, QMessageBox, QTabWidget
)
from logic.production_docs import (
    WAX_JOBS_POOL,
    ORDERS_POOL,
    METHOD_LABEL,
    process_new_order,
)
from pages.orders_page import parse_variant
from core.com_bridge import log
from config import BRIDGE as bridge, CSS_TREE

class WaxPage(QWidget):
    def __init__(self):
        super().__init__()
        self.last_created_task_ref = None
        self.jobs_page = None
        self._ui()
        self.refresh()

    def set_jobs_page(self, jobs_page):
        """Позволяет передать внешний объект страницы нарядов."""
        self.jobs_page = jobs_page

    def refresh(self):
        self._fill_jobs_tree()
        self._fill_parties_tree()
        self._fill_tasks_tree()
        self._fill_wax_jobs_tree()

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

        from PyQt5.QtWidgets import QComboBox

        # Список мастеров для задания
        self.combo_employee = QComboBox()
        self.combo_employee.setMinimumWidth(200)
        self.combo_employee.addItem("— выберите мастера —")

        # загружаем из 1С
        for item in bridge.list_catalog_items("ФизическиеЛица", limit=100):
            name = item.get("Description", "")
            if name:
                self.combo_employee.addItem(name)

        btn_row.addWidget(self.combo_employee)

        for b in [btn_create_task, btn_create_wax_jobs]:
            btn_row.addWidget(b, alignment=Qt.AlignLeft)

        # -------- выбор мастеров для нарядов --------
        label = QLabel("→ выберите мастеров")
        label.setStyleSheet("font-weight: bold; padding: 6px")

        self.combo_3d_master = QComboBox()
        self.combo_resin_master = QComboBox()

        employees = bridge.list_catalog_items("ФизическиеЛица", limit=100)
        names = [e.get("Description", "") for e in employees]
        self.combo_3d_master.addItems(names)
        self.combo_resin_master.addItems(names)

        h = QHBoxLayout()
        h.addWidget(QLabel("3D:"))
        h.addWidget(self.combo_3d_master)
        h.addWidget(QLabel("Пресс-форма:"))
        h.addWidget(self.combo_resin_master)

        t1.addWidget(label)
        t1.addLayout(h)

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
        self.tree_tasks.setHeaderLabels(["✓", "Номер", "Дата", "Участок", "Операция", "Ответственный"])
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

        # подключения сигналов
        self.tree_tasks.itemDoubleClicked.connect(self._on_task_double_click)
        self.tree_acts.itemDoubleClicked.connect(self._on_wax_job_double_click)

        self.tabs.addTab(tab3, "Наряды из 1С")
        
        btn_bar = QHBoxLayout()

        btn_refresh = QPushButton("🔄 Обновить")
        btn_post = QPushButton("✅ Провести отмеченные")
        btn_unpost = QPushButton("↩ Отменить проведение")
        btn_mark = QPushButton("🏷 Пометить")
        btn_unmark = QPushButton("🚫 Снять пометку")
        btn_delete = QPushButton("🗑 Удалить")

        btn_bar.addWidget(btn_refresh)
        btn_bar.addWidget(btn_post)
        btn_bar.addWidget(btn_unpost)
        btn_bar.addWidget(btn_mark)
        btn_bar.addWidget(btn_unmark)
        btn_bar.addWidget(btn_delete)

        btn_refresh.clicked.connect(self._fill_tasks_tree)
        btn_post.clicked.connect(self._post_selected_tasks)
        btn_unpost.clicked.connect(self._unpost_selected_tasks)
        btn_mark.clicked.connect(self._mark_selected_tasks)
        btn_unmark.clicked.connect(self._unmark_selected_tasks)
        btn_delete.clicked.connect(self._delete_selected_tasks)

        t2.addLayout(btn_bar)

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
        
    def _get_checked_tasks(self):
        result = []
        for i in range(self.tree_tasks.topLevelItemCount()):
            item = self.tree_tasks.topLevelItem(i)
            if item.checkState(0) == Qt.Checked:
                result.append(item.text(1))  # Номер
        print("[DEBUG] Отмечены задания:", result)
        return result
    
    def _post_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            bridge.post_task(num)
        self._fill_tasks_tree()

    def _unpost_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            bridge.undo_post_task(num)
        self._fill_tasks_tree()

    def _mark_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            bridge.mark_task_for_deletion(num)
        self._fill_tasks_tree()

    def _unmark_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            bridge.unmark_task_deletion(num)
        self._fill_tasks_tree()

    def _delete_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            bridge.delete_task(num)
        self._fill_tasks_tree() 

    def populate_jobs_tree(self, doc_num: str):
        self.tree_jobs.clear()
        self.tree_part.clear()

        try:
            # Попытка как наряд
            rows = bridge.get_wax_job_rows(doc_num)
            if rows:
                log(f"[populate_jobs_tree] Найдены строки наряда №{doc_num}: {len(rows)}")
                self.refresh()
                return
        except Exception as e:
            log(f"[populate_jobs_tree] Не удалось как наряд: {e}")

        try:
            # Попытка как задание на производство
            task = bridge._find_task_by_number(doc_num)
            if task and getattr(task, "ДокументОснование", None):
                base_ref = getattr(task, "ДокументОснование")
                base_obj = bridge.get_object_from_ref(base_ref)
                order_num = base_obj.Номер
                log(f"[populate_jobs_tree] Задание связано с заказом №{order_num}")

                # Загружаем строки заказа и формируем данные как для нового заказа
                order_lines = bridge.get_order_lines(order_num)
                order_json_rows = []
                for r in order_lines:
                    metal, hallmark, color = parse_variant(r.get("method", ""))
                    order_json_rows.append(
                        {
                            "article": r.get("article", ""),
                            "size": r.get("size", 0),
                            "qty": r.get("qty", 0),
                            "weight": r.get("weight", 0),
                            "metal": metal,
                            "hallmark": hallmark,
                            "color": color,
                        }
                    )

                # Перезаписываем ORDERS_POOL и используем существующую обработку
                ORDERS_POOL.clear()
                process_new_order({"number": order_num, "rows": order_json_rows})
                log(
                    f"[ORDERS_POOL] Добавлены строки и партии для заказа №{order_num}"
                )

                # Перерисовываем
                self.refresh()
            else:
                log("[populate_jobs_tree] Не найден документ основания")
        except Exception as ee:
            log(f"[populate_jobs_tree] Ошибка при получении заказа из задания: {ee}")

    def _on_task_double_click(self, item, column):
        num = item.text(1).strip() if item.columnCount() > 1 else item.text(0).strip()
        if not num:
            return

        task_ref = bridge._find_task_by_number(num)   # <-- это ссылка (Ref)
        self.last_created_task_ref = task_ref

        self.tabs.setCurrentIndex(0)
        self.populate_jobs_tree(num)                  # <-- передаём строковый номер
        log(f"[UI] Выбрано задание №{num}, переходим к созданию нарядов.")

    def load_task_data(self, task_obj):
        if not task_obj:
            log("[UI] ❌ Нет задания для отображения.")
            return

        self.last_created_task_ref = task_obj
        self._fill_jobs_table_from_task(task_obj)
        log(f"[UI] ✅ Загружены данные задания №{task_obj.Номер}")

    def _fill_jobs_table_from_task(self, task_obj):
        """Заполняет таблицу нарядов строками выбранного задания."""
        lines = bridge.get_task_lines(getattr(task_obj, "Номер", ""))
        log(f"[UI] Загрузка строк задания: {len(lines)}")

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
            status = ""
            if t.get("posted"):
                status = "✅"
            elif t.get("deleted"):
                status = "🗑"
            item = QTreeWidgetItem([
                status,
                t.get("num", ""),
                t.get("date", ""),
                t.get("section", ""),
                t.get("tech_op", ""),
                t.get("employee", "")
            ])
            item.setCheckState(0, Qt.Unchecked)
            from PyQt5.QtGui import QBrush, QColor

            if t.get("deleted"):
                for i in range(item.columnCount()):
                    item.setBackground(i, QBrush(QColor("#f87171")))  # красный
            elif t.get("posted"):
                for i in range(item.columnCount()):
                    item.setBackground(i, QBrush(QColor("#bbf7d0")))  # зелёный
            self.tree_tasks.addTopLevelItem(item)

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
            docs = o.get("docs", {})
            if docs.get("sync_task_num"):
                continue  # ⛔️ уже создано

            order_num = o.get("number") or docs.get("order_code")
            if not order_num:
                log("❌ У заказа нет номера")
                continue

            order_ref = bridge.get_doc_ref("ЗаказВПроизводство", order_num)
            if not order_ref:
                log(f"❌ Не удалось получить ссылку на заказ №{order_num}")
                continue

            rows = bridge.get_order_lines(order_num)
            # 🔄 обогащаем строку заказа
            for row in rows:
                # Пробуем вытащить пробу и цвет из варианта
                variant = row.get("method", "")
                if "585" in variant:
                    row["assay"] = "585"
                elif "925" in variant:
                    row["assay"] = "925"
                else:
                    row["assay"] = ""

                if "Красный" in variant:
                    row["color"] = "Красный"
                elif "Желтый" in variant:
                    row["color"] = "Желтый"
                elif "серебро" in variant.lower():
                    row["color"] = "Светлый"
                else:
                    row["color"] = ""

                if not row.get("weight"):
                    # Ставим None вместо 0, чтобы избежать подстановки фиктивного 1
                    row["weight"] = None

                # Мастер
                employee_name = self.combo_employee.currentText()
                if employee_name and employee_name != "— выберите мастера —":
                    row["employee"] = employee_name
                else:
                    QMessageBox.warning(self, "Мастер не выбран", "Пожалуйста, выберите мастера")
                    return
            # передаём выбранного мастера
            employee_name = self.combo_employee.currentText()
            if employee_name and employee_name != "— выберите мастера —":
                for r in rows:
                    r["employee"] = employee_name
            else:
                QMessageBox.warning(self, "Мастер не выбран", "Пожалуйста, выберите мастера (рабочий центр)")
                return
            if not rows:
                log(f"❌ В заказе №{order_num} нет строк для задания")
                continue

            try:
                result = bridge.create_production_task(order_ref, rows)
                docs["sync_task_num"] = result.get("Номер")  # ✅ запоминаем
                ref = result.get("Ref")
                self.last_created_task_ref = bridge.get_object_from_ref(ref) if ref else None
                log(f"✅ Создано задание №{result.get('Номер', '?')}")
            except Exception as e:
                log(f"❌ Ошибка создания задания: {e}")
                
            if not docs.get("batches"):
                from collections import defaultdict
                rows_by_batch = defaultdict(lambda: {"qty": 0, "total_w": 0.0})
                for row in rows:
                    key = (row["metal"], row["assay"], row["color"])
                    rows_by_batch[key]["qty"] += row["qty"]
                    rows_by_batch[key]["total_w"] += row["weight"]

                docs["batches"] = []
                for (metal, hallmark, color), data in rows_by_batch.items():
                    docs["batches"].append({
                        "batch_barcode": "AUTO",  # или сгенерируй штрихкод
                        "metal": metal,
                        "hallmark": hallmark,
                        "color": color,
                        "qty": data["qty"],
                        "total_w": round(data["total_w"], 3)
                    })    

    # ------------------------------------------------------------------
    def _create_wax_jobs(self):
        # Для создания наряда достаточно выбранного задания.
        # Проверка ORDERS_POOL мешала создавать наряды для уже существующих
        # заданий, поэтому её убрали.

        master_3d = self.combo_3d_master.currentText().strip()
        master_resin = self.combo_resin_master.currentText().strip()

        if not master_3d or not master_resin:
            QMessageBox.warning(self, "Ошибка", "Выберите мастеров для обоих методов.")
            return


        if not self.last_created_task_ref:

            QMessageBox.warning(self, "Ошибка", "Нет выбранного задания для создания нарядов.")
            return

        print("[DEBUG] last_created_task_ref =", self.last_created_task_ref)
        result = bridge.create_multiple_wax_jobs_from_task(
            self.last_created_task_ref,
            {"3D печать": master_3d, "Пресс-форма": master_resin}
        )

        if result:
            QMessageBox.information(self, "Успех", "Созданы наряды: " + ", ".join(result))
            self.refresh()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось создать наряды.")

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
def _wax_method(article: str) -> str:
    art = str(article).lower()
    if "д" in art or "d" in art:
        return "3D печать"
    return "Пресс-форма"
