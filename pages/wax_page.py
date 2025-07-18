# wax_page.py • v0.8
# ─────────────────────────────────────────────────────────────────────────
from collections import defaultdict
import re
from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QFont
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QPushButton, QMessageBox, QTabWidget, QInputDialog,
    QComboBox, QFormLayout, QTableWidget, QTableWidgetItem, QGroupBox
)
from logic.production_docs import (
    WAX_JOBS_POOL,
    ORDERS_POOL,
    METHOD_LABEL,
    process_new_order,
)
from pages.orders_page import parse_variant
from core.com_bridge import log
import config
from config import CSS_TREE
from widgets.production_task_form import ProductionTaskEditForm

class WaxPage(QWidget):
    def __init__(self):
        super().__init__()
        self.last_created_task_ref = None
        self.jobs_page = None
        self.close_job_refs = []
        self._task_select_callback = None
        self.warehouses = config.BRIDGE.list_catalog_items("Склады")
        self.norm_types = (
            config.BRIDGE.list_enum_values("ВидыНормативовНоменклатуры")
            or ["Номенклатура", "Комплектующее"]
        )
        self._ui()
        self.refresh()

    def set_jobs_page(self, jobs_page):
        """Позволяет передать внешний объект страницы нарядов."""
        self.jobs_page = jobs_page

    def goto_order_selection(self, callback=None):
        """Открывает страницу заказов и передаёт callback выбора."""
        main_win = self.window()
        if hasattr(main_win, "menu") and hasattr(main_win, "page_idx"):
            orders_idx = main_win.page_idx.get("orders")
            if orders_idx is not None:
                main_win.menu.setCurrentRow(orders_idx)
                orders_page = main_win.page_refs.get("orders")
                if orders_page and hasattr(orders_page, "set_selection_callback"):
                    orders_page.set_selection_callback(callback)
                    if hasattr(orders_page, "tabs"):
                        orders_page.tabs.setCurrentIndex(1)

    def select_task_for_wax_jobs(self):
        """Переключается на список заданий для выбора."""
        self._task_select_callback = self.load_task_data
        self.tabs.setCurrentWidget(self.tab_tasks)
        if hasattr(self, "tabs_tasks"):
            self.tabs_tasks.setCurrentIndex(1)

    def select_task_for_wax_close(self):
        """Выбор задания для закрытия нарядов."""
        self._task_select_callback = self.load_close_task_data
        self.tabs.setCurrentWidget(self.tab_tasks)
        if hasattr(self, "tabs_tasks"):
            self.tabs_tasks.setCurrentIndex(1)

    def refresh(self):
        """Обновляет данные текущей вкладки."""
        self._fill_tasks_tree()
        self._fill_jobs_tree()
        self._fill_parties_tree()
        self._fill_assembly_tree()

    # ------------------------------------------------------------------
    def _ui(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(40, 30, 40, 30)

        hdr = QLabel("Воскование / 3-D печать")
        hdr.setFont(QFont("Arial", 22, QFont.Bold))
        v.addWidget(hdr)

        self.tabs = QTabWidget()
        v.addWidget(self.tabs, 1)

        # ----- Tab: Задания на производство -----
        self.tab_tasks = QWidget()
        t_main = QVBoxLayout(self.tab_tasks)
        self.tabs_tasks = QTabWidget()
        t_main.addWidget(self.tabs_tasks, 1)

        # --- sub-tab: создание задания ---
        tab_task_new = QWidget(); t_new = QVBoxLayout(tab_task_new)
        self.task_form = ProductionTaskEditForm(config.BRIDGE)
        self.task_form.task_saved.connect(
            lambda ref: setattr(self, "last_created_task_ref", config.BRIDGE.get_object_from_ref(ref) if ref else None)
        )
        t_new.addWidget(self.task_form)

        tab_tasks_list = QWidget(); t1 = QVBoxLayout(tab_tasks_list)
        lbl2 = QLabel("Задания на производство")
        lbl2.setFont(QFont("Arial", 16, QFont.Bold))
        t1.addWidget(lbl2)

        self.tree_tasks = QTreeWidget()
        self.tree_tasks.setHeaderLabels([
            "✓", "Номер", "Дата", "Участок", "Операция", "Ответственный"
        ])
        self.tree_tasks.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_tasks.setStyleSheet(CSS_TREE)
        t1.addWidget(self.tree_tasks, 1)

        btn_bar = QHBoxLayout()

        btn_refresh = QPushButton("🔄 Обновить")
        btn_post = QPushButton("✅ Провести отмеченные")
        btn_unpost = QPushButton("↩ Отменить проведение")
        btn_mark = QPushButton("🏷 Пометить")
        btn_unmark = QPushButton("🚫 Снять пометку")
        btn_delete = QPushButton("🗑 Удалить")
        btn_to_work = QPushButton("📤 В работу")

        btn_bar.addWidget(btn_refresh)
        btn_bar.addWidget(btn_post)
        btn_bar.addWidget(btn_unpost)
        btn_bar.addWidget(btn_mark)
        btn_bar.addWidget(btn_unmark)
        btn_bar.addWidget(btn_delete)
        btn_bar.addWidget(btn_to_work)

        btn_refresh.clicked.connect(self._fill_tasks_tree)
        btn_post.clicked.connect(self._post_selected_tasks)
        btn_unpost.clicked.connect(self._unpost_selected_tasks)
        btn_mark.clicked.connect(self._mark_selected_tasks)
        btn_unmark.clicked.connect(self._unmark_selected_tasks)
        btn_delete.clicked.connect(self._delete_selected_tasks)
        btn_to_work.clicked.connect(self._send_task_to_work)

        t1.addLayout(btn_bar)

        self.tabs_tasks.addTab(tab_task_new, "Создание")
        self.tabs_tasks.addTab(tab_tasks_list, "Задания")
        self.tabs.addTab(self.tab_tasks, "Задания на производство")

        # ----- Tab: Наряды восковых изделий по методам -----
        self.tab_jobs = QWidget()
        j_main = QVBoxLayout(self.tab_jobs)
        self.tabs_jobs = QTabWidget()
        j_main.addWidget(self.tabs_jobs, 1)

        tab_jobs_new = QWidget(); j_new = QVBoxLayout(tab_jobs_new)
        lbl_new = QLabel("Создание нарядов")
        lbl_new.setFont(QFont("Arial", 16, QFont.Bold))
        j_new.addWidget(lbl_new)

        form = QFormLayout()
        self.combo_3d_master = QComboBox(); self.combo_3d_master.setEditable(True)
        self.combo_3d_master.addItems(config.EMPLOYEES)
        self.combo_form_master = QComboBox(); self.combo_form_master.setEditable(True)
        self.combo_form_master.addItems(config.EMPLOYEES)
        self.combo_warehouse = QComboBox(); self.combo_warehouse.setEditable(True)
        self.combo_warehouse.addItems([x["Description"] for x in self.warehouses])
        self.combo_norm_type = QComboBox(); self.combo_norm_type.setEditable(True)
        self.combo_norm_type.addItems(self.norm_types)
        if self.norm_types:
            self.combo_norm_type.setCurrentIndex(0)
        form.addRow("3D печать", self.combo_3d_master)
        form.addRow("Пресс-форма", self.combo_form_master)
        form.addRow("Склад", self.combo_warehouse)
        form.addRow("Вид норматива", self.combo_norm_type)
        j_new.addLayout(form)

        self.btn_select_task = QPushButton("Выбрать задание")
        self.btn_select_task.clicked.connect(self.select_task_for_wax_jobs)
        j_new.addWidget(self.btn_select_task, alignment=Qt.AlignLeft)


        self.lbl_task_number = QLabel("Задание: № —")
        self.lbl_task_date = QLabel("Дата: —")

        info_layout = QHBoxLayout()
        info_layout.addWidget(self.lbl_task_number)
        info_layout.addWidget(self.lbl_task_date)
        j_new.addLayout(info_layout)

        self.lbl_task_info = QLabel("Задание не выбрано")
        j_new.addWidget(self.lbl_task_info, alignment=Qt.AlignLeft)


        tables_layout = QHBoxLayout()

        group_3d = QGroupBox("3D печать")
        v_3d = QVBoxLayout(group_3d)
        self.tbl_3d = QTableWidget(0, 7)
        self.tbl_3d.setHorizontalHeaderLabels([
            "Номенклатура", "Артикул", "Размер", "Проба", "Цвет",
            "Кол-во", "Вес"
        ])
        self.tbl_3d.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_3d.verticalHeader().setVisible(False)
        v_3d.addWidget(self.tbl_3d)

        group_form = QGroupBox("Пресс-форма")
        v_form = QVBoxLayout(group_form)
        self.tbl_form = QTableWidget(0, 7)
        self.tbl_form.setHorizontalHeaderLabels([
            "Номенклатура", "Артикул", "Размер", "Проба", "Цвет",
            "Кол-во", "Вес"
        ])
        self.tbl_form.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_form.verticalHeader().setVisible(False)
        v_form.addWidget(self.tbl_form)

        tables_layout.addWidget(group_3d)
        tables_layout.addWidget(group_form)
        j_new.addLayout(tables_layout, 1)

        btn_create_jobs = QPushButton("Создать наряды")
        btn_create_jobs.clicked.connect(self._create_wax_jobs)
        j_new.addWidget(btn_create_jobs, alignment=Qt.AlignLeft)

        # ---- Панель управления нарядами (создание)
        bar_new_jobs = QHBoxLayout()
        self.btn_job_save = QPushButton("💾 Записать")
        self.btn_job_post = QPushButton("✅ Провести")
        self.btn_job_row_add = QPushButton("Строка +")
        self.btn_job_row_del = QPushButton("Строка -")
        self.btn_job_row_copy = QPushButton("Копировать строку")
        self.btn_job_new = QPushButton("Новый наряд")

        for b in [
            self.btn_job_save,
            self.btn_job_post,
            self.btn_job_row_add,
            self.btn_job_row_del,
            self.btn_job_row_copy,
            self.btn_job_new,
        ]:
            bar_new_jobs.addWidget(b)
            b.clicked.connect(self._not_implemented)
        j_new.addLayout(bar_new_jobs)

        # --- sub-tab: закрытие нарядов ---
        tab_jobs_close = QWidget(); c_layout = QVBoxLayout(tab_jobs_close)
        self.btn_close_select_task = QPushButton("Выбрать задание")
        self.btn_close_select_task.clicked.connect(self.select_task_for_wax_close)
        c_layout.addWidget(self.btn_close_select_task, alignment=Qt.AlignLeft)

        self.lbl_close_info = QLabel("Задание не выбрано")
        c_layout.addWidget(self.lbl_close_info, alignment=Qt.AlignLeft)

        close_tables = QHBoxLayout()

        group_c3d = QGroupBox("3D печать")
        vc3d = QVBoxLayout(group_c3d)
        self.tbl_close_3d = QTableWidget(0, 7)
        self.tbl_close_3d.setHorizontalHeaderLabels([
            "✓", "Номенклатура", "Размер", "Проба", "Цвет", "Кол-во", "Вес"
        ])
        self.tbl_close_3d.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_close_3d.verticalHeader().setVisible(False)
        vc3d.addWidget(self.tbl_close_3d)

        group_cform = QGroupBox("Пресс-форма")
        vcform = QVBoxLayout(group_cform)
        self.tbl_close_form = QTableWidget(0, 7)
        self.tbl_close_form.setHorizontalHeaderLabels([
            "✓", "Номенклатура", "Размер", "Проба", "Цвет", "Кол-во", "Вес"
        ])
        self.tbl_close_form.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_close_form.verticalHeader().setVisible(False)
        vcform.addWidget(self.tbl_close_form)

        close_tables.addWidget(group_c3d)
        close_tables.addWidget(group_cform)
        c_layout.addLayout(close_tables, 1)

        self.btn_close_jobs = QPushButton("Закрыть наряды")
        self.btn_close_jobs.clicked.connect(self._on_close_jobs)
        c_layout.addWidget(self.btn_close_jobs, alignment=Qt.AlignLeft)

        tab_jobs_list = QWidget(); j1 = QVBoxLayout(tab_jobs_list)
        lbl_jobs = QLabel("Наряды (восковые изделия)")
        lbl_jobs.setFont(QFont("Arial", 16, QFont.Bold))
        j1.addWidget(lbl_jobs)

        self.tree_jobs = QTreeWidget()
        self.tree_jobs.setHeaderLabels([
            "Номер",
            "Дата",
            "Закрыт",
            "Сотрудник",
            "Тех. операция",
            "Комментарий",
            "Склад",
            "Участок",
            "Организация",
            "Задание",
            "Ответственный",
        ])
        self.tree_jobs.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_jobs.setStyleSheet(CSS_TREE)
        j1.addWidget(self.tree_jobs, 1)

        btn_bar_jobs = QHBoxLayout()
        btn_jobs_refresh = QPushButton("🔄 Обновить")
        btn_jobs_post = QPushButton("✅ Провести отмеченные")
        btn_jobs_unpost = QPushButton("↩️ Отменить проведение")
        btn_jobs_mark = QPushButton("✏️ Пометить")
        btn_jobs_unmark = QPushButton("🚫 Снять пометку")
        btn_jobs_delete = QPushButton("🗑 Удалить")
        btn_jobs_to_work = QPushButton("🧑‍🔧 В работу")

        for b in [
            btn_jobs_refresh,
            btn_jobs_post,
            btn_jobs_unpost,
            btn_jobs_mark,
            btn_jobs_unmark,
            btn_jobs_delete,
            btn_jobs_to_work,
        ]:
            btn_bar_jobs.addWidget(b)

        btn_jobs_refresh.clicked.connect(self._fill_jobs_tree)
        btn_jobs_post.clicked.connect(self._post_selected_jobs)
        btn_jobs_unpost.clicked.connect(self._unpost_selected_jobs)
        btn_jobs_mark.clicked.connect(self._mark_selected_jobs)
        btn_jobs_unmark.clicked.connect(self._unmark_selected_jobs)
        btn_jobs_delete.clicked.connect(self._delete_selected_jobs)
        btn_jobs_to_work.clicked.connect(self._send_job_to_work)

        j1.addLayout(btn_bar_jobs)

        # --- sub-tab: сборка ёлок ---
        self.tab_tree = QWidget(); tree_layout = QVBoxLayout(self.tab_tree)
        lbl_tree = QLabel("Сборка ёлок")
        lbl_tree.setFont(QFont("Arial", 16, QFont.Bold))
        tree_layout.addWidget(lbl_tree)

        self.tree_assembly = QTreeWidget()
        self.tree_assembly.setHeaderLabels(["Параметры", "Кол-во", "Вес"])
        self.tree_assembly.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree_assembly.setStyleSheet(CSS_TREE)
        tree_layout.addWidget(self.tree_assembly, 1)

        bar_tree = QHBoxLayout()
        btn_form_trees = QPushButton("🌲 Сформировать ёлки")
        btn_clear_trees = QPushButton("🧹 Очистить список")
        btn_form_trees.clicked.connect(self._form_trees)
        btn_clear_trees.clicked.connect(self._clear_assembly_pool)
        for b in [btn_form_trees, btn_clear_trees]:
            bar_tree.addWidget(b)
        tree_layout.addLayout(bar_tree)

        self.tabs_jobs.addTab(tab_jobs_new, "Создание")
        self.tabs_jobs.addTab(tab_jobs_close, "Закрытие")
        self.tabs_jobs.addTab(tab_jobs_list, "Наряды (восковые изделия)")
        self.tabs_jobs.addTab(self.tab_tree, "Сборка ёлок")
        self.tabs.addTab(self.tab_jobs, "Наряды восковых изделий по методам")

        # подключения сигналов
        self.tree_tasks.itemDoubleClicked.connect(self._on_task_double_click)
        self.tree_jobs.itemDoubleClicked.connect(self._on_wax_job_double_click)

    def _show_wax_job_detail(self, item):
        from PyQt5.QtWidgets import QDialog, QTableWidget, QTableWidgetItem, QVBoxLayout
        num = item.text(0)
        rows = config.BRIDGE.get_wax_job_rows(num)

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
                "Номенклатура",
                "Размер",
                "Проба",
                "Цвет",
                "Количество",
                "Вес",
                "Партия",
                "Номер ёлки",
                "Состав набора",
            ]):
                val = r.get(k, "")
                if k == "Вес" and val != "":
                    val = f"{val:.{config.WEIGHT_DECIMALS}f}"
                tbl.setItem(i, j, QTableWidgetItem(str(val)))

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

    def _selected_job_code(self):
        """Возвращает код выбранного наряда или None."""
        item = self.tree_jobs.currentItem()
        if not item:
            return None

        # Если выбран подэлемент, поднимаемся к корневому уровню
        while item.parent():
            item = item.parent()

        text = item.text(0)
        m = re.search(r"\((WX[^)]+)\)", text)
        return m.group(1) if m else None
    
    def _post_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            config.BRIDGE.post_task(num)
        self._fill_tasks_tree()

    def _unpost_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            config.BRIDGE.undo_post_task(num)
        self._fill_tasks_tree()

    def _mark_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            config.BRIDGE.mark_task_for_deletion(num)
        self._fill_tasks_tree()

    def _unmark_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            config.BRIDGE.unmark_task_deletion(num)
        self._fill_tasks_tree()

    def _delete_selected_tasks(self):
        for num in self._get_checked_tasks():
            print(f"[DEBUG] Проведение: {num}")
            config.BRIDGE.delete_task(num)
        self._fill_tasks_tree()

    # ------------------------------------------------------------------
    def _get_checked_jobs(self):
        result = []
        for i in range(self.tree_jobs.topLevelItemCount()):
            item = self.tree_jobs.topLevelItem(i)
            if item.checkState(0) == Qt.Checked:
                result.append(item.text(0))
        return result

    def _post_selected_jobs(self):
        for num in self._get_checked_jobs():
            config.BRIDGE.post_wax_job(num)
        self._fill_jobs_tree()

    def _unpost_selected_jobs(self):
        for num in self._get_checked_jobs():
            config.BRIDGE.undo_post_wax_job(num)
        self._fill_jobs_tree()

    def _mark_selected_jobs(self):
        for num in self._get_checked_jobs():
            config.BRIDGE.mark_wax_job_for_deletion(num)
        self._fill_jobs_tree()

    def _unmark_selected_jobs(self):
        for num in self._get_checked_jobs():
            config.BRIDGE.unmark_wax_job_deletion(num)
        self._fill_jobs_tree()

    def _delete_selected_jobs(self):
        for num in self._get_checked_jobs():
            config.BRIDGE.delete_wax_job(num)
        self._fill_jobs_tree()

    def _send_job_to_work(self):
        """Перенос выбранные наряды на вкладку сборки ёлок."""
        jobs = self._get_checked_jobs()
        if not jobs:
            item = self.tree_jobs.currentItem()
            if item:
                jobs = [item.text(0).strip()]
        if not jobs:

            log("[UI] ❌ Не выбраны наряды для отправки в сборку")
            QMessageBox.warning(self, "Ошибка", "Выберите наряды")
            return

        log(f"[UI] Отправка нарядов в сборку: {', '.join(jobs)}")

        added = False
        for num in jobs:
            item_obj = None
            for i in range(self.tree_jobs.topLevelItemCount()):
                it = self.tree_jobs.topLevelItem(i)
                if it.text(0).strip() == num:
                    item_obj = it
                    break
            if item_obj and item_obj.text(2).strip() != "✅":

                log(f"[UI] Наряд {num} не закрыт")

                QMessageBox.warning(self, "Ошибка", f"Наряд {num} не закрыт")
                continue
            self._add_job_to_assembly(num)
            added = True
        if not added:

            log("[UI] ❌ Не удалось добавить наряды в сборку")

            return
        if hasattr(self, "tabs_jobs"):
            self.tabs_jobs.setCurrentWidget(self.tab_tree)
    def _send_task_to_work(self):
        """Переносит выбранное задание на вкладку создания нарядов."""
        item = self.tree_tasks.currentItem()
        if not item:
            QMessageBox.warning(self, "Ошибка", "Выберите задание")
            return
        num = item.text(1).strip() if item.columnCount() > 1 else item.text(0).strip()
        if not num:
            return
        task_obj = config.BRIDGE._find_task_by_number(num)
        if hasattr(task_obj, "GetObject"):
            task_obj = task_obj.GetObject()
        self.load_task_data(task_obj)
        self.tabs.setCurrentWidget(self.tab_jobs)
        if hasattr(self, "tabs_jobs"):
            self.tabs_jobs.setCurrentIndex(0)

    def populate_jobs_tree(self, doc_num: str):
        self.tree_jobs.clear()
        self.tree_part.clear()

        try:
            # Попытка как наряд
            rows = config.BRIDGE.get_wax_job_rows(doc_num)
            if rows:
                log(f"[populate_jobs_tree] Найдены строки наряда №{doc_num}: {len(rows)}")
                self.refresh()
                return
        except Exception as e:
            log(f"[populate_jobs_tree] Не удалось как наряд: {e}")

        try:
            # Попытка как задание на производство
            task = config.BRIDGE._find_task_by_number(doc_num)
            if task and getattr(task, "ДокументОснование", None):
                base_ref = getattr(task, "ДокументОснование")
                base_obj = config.BRIDGE.get_object_from_ref(base_ref)
                order_num = base_obj.Номер
                log(f"[populate_jobs_tree] Задание связано с заказом №{order_num}")

                # Загружаем строки заказа и формируем данные как для нового заказа
                order_lines = config.BRIDGE.get_order_lines(order_num)
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

        task_obj = config.BRIDGE._find_task_by_number(num)
        if hasattr(task_obj, "GetObject"):
            task_obj = task_obj.GetObject()
        self.last_created_task_ref = task_obj
        if self._task_select_callback:
            cb = self._task_select_callback
            self._task_select_callback = None
            if hasattr(self, "tabs_jobs"):
                if cb == self.load_close_task_data:
                    self.tabs_jobs.setCurrentIndex(1)
                else:
                    self.tabs_jobs.setCurrentIndex(0)
            self.tabs.setCurrentWidget(self.tab_jobs)
            cb(task_obj)
        else:
            self.tabs.setCurrentIndex(0)
            if hasattr(self, "tabs_tasks"):
                self.tabs_tasks.setCurrentIndex(0)
                self.task_form.load_task_object(task_obj)
        log(f"[UI] Выбрано задание №{num}.")

    def load_task_data(self, task_obj):
        if not task_obj:
            log("[UI] ❌ Нет задания для отображения.")
            return

        self.last_created_task_ref = task_obj
        self._fill_jobs_table_from_task(task_obj)


        if hasattr(self, "lbl_task_number") and hasattr(self, "lbl_task_date"):
            self.lbl_task_number.setText(f"Задание: № {task_obj.Номер}")
            try:
                d = task_obj.Дата.strftime("%d.%m.%Y")
            except Exception:
                d = ""
            self.lbl_task_date.setText(f"Дата: {d}")

        if hasattr(self, "lbl_task_info"):
            try:
                d = str(task_obj.Дата)[:10]
            except Exception:
                d = ""
            self.lbl_task_info.setText(f"Задание №{task_obj.Номер} от {d}")


        log(f"[UI] ✅ Загружены данные задания №{task_obj.Номер}")

    def load_close_task_data(self, task_obj):
        if not task_obj:
            log("[UI] ❌ Нет задания для закрытия.")
            return

        self.last_created_task_ref = task_obj
        if hasattr(self, "lbl_close_info"):
            try:
                d = str(task_obj.Дата)[:10]
            except Exception:
                d = ""
            self.lbl_close_info.setText(f"Задание №{task_obj.Номер} от {d}")

        self._fill_close_tables_from_task(task_obj)

    def _fill_jobs_table_from_task(self, task_obj):
        """Заполняет таблицу нарядов строками выбранного задания."""
        lines = config.BRIDGE.get_task_lines(getattr(task_obj, "Номер", ""))
        if not hasattr(self, "tbl_3d") or not hasattr(self, "tbl_form"):
            log("[UI] Таблицы для строк задания не инициализированы")
            return

        self.tbl_3d.setRowCount(0)
        self.tbl_form.setRowCount(0)

        for row in lines:
            art = str(row.get("article", "")).lower()
            is_3d = "д" in art or "d" in art
            table = self.tbl_3d if is_3d else self.tbl_form

            r = table.rowCount()
            table.insertRow(r)
            values = [
                row.get("nomen", ""),
                row.get("article", ""),
                row.get("size", ""),
                row.get("sample", ""),
                row.get("color", ""),
                row.get("qty", ""),
                f"{row['weight']:.{config.WEIGHT_DECIMALS}f}" if row.get("weight") not in ("", None) else "",
            ]
            for c, v in enumerate(values):
                table.setItem(r, c, QTableWidgetItem(str(v)))

        self.tbl_3d.resizeColumnsToContents()
        self.tbl_form.resizeColumnsToContents()
        log(f"[UI] Загрузка строк задания: {len(lines)}")

    def _fill_close_tables_from_task(self, task_obj):
        """Заполняет таблицы закрытия строками нарядов по заданию."""
        if not hasattr(self, "tbl_close_3d"):
            return
        self.tbl_close_3d.setRowCount(0)
        self.tbl_close_form.setRowCount(0)
        self.close_job_refs = []

        refs = config.BRIDGE.find_wax_jobs_by_task(task_obj.Ref)
        for ref in refs:
            job_obj = config.BRIDGE.get_object_from_ref(ref)
            if not job_obj:
                continue
            self.close_job_refs.append(ref)
            method_obj = getattr(job_obj, "ТехОперация", None)
            method = str(
                method_obj.Description if hasattr(method_obj, "Description") else method_obj
            ).lower()
            is_3d = "3d" in method or "3-д" in method or "3д" in method
            table = self.tbl_close_3d if is_3d else self.tbl_close_form
            rows = config.BRIDGE.get_wax_job_lines_by_ref(job_obj.Ref)
            for r_data in rows:
                r = table.rowCount()
                table.insertRow(r)
                chk = QTableWidgetItem()
                chk.setCheckState(Qt.Checked)
                chk.setData(Qt.UserRole, ref)
                table.setItem(r, 0, chk)
                values = [
                    r_data.get("nomen", ""),
                    r_data.get("size", ""),
                    r_data.get("sample", ""),
                    r_data.get("color", ""),
                    r_data.get("qty", ""),
                    f"{r_data['weight']:.{config.WEIGHT_DECIMALS}f}" if r_data.get("weight") not in ("", None) else "",
                ]
                for c, v in enumerate(values, start=1):
                    table.setItem(r, c, QTableWidgetItem(str(v)))

        self.tbl_close_3d.resizeColumnsToContents()
        self.tbl_close_form.resizeColumnsToContents()

    def _on_wax_job_double_click(self, item, column):
        num = item.text(0).split()[0].strip()
        if not num:
            return

        lines = config.BRIDGE.get_wax_job_lines(num)
        if not lines:
            QMessageBox.information(self, "Нет данных", f"В наряде {num} нет строк")
            return

        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTreeWidget, QTreeWidgetItem

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Строки наряда {num}")
        layout = QVBoxLayout(dlg)

        tree = QTreeWidget()
        tree.setHeaderLabels([
            "Номенклатура",
            "Размер",
            "Проба",
            "Цвет",
            "Кол-во",
            "Вес",
        ])
        for row in lines:
            QTreeWidgetItem(
                tree,
                [
                    row["nomen"],
                    str(row["size"]),
                    str(row["sample"]),
                    str(row["color"]),
                    str(row["qty"]),
                    f"{row['weight']:.{config.WEIGHT_DECIMALS}f}",
                ],
            )
        layout.addWidget(tree)
        dlg.resize(700, 400)
        dlg.exec_()

    def _fill_tasks_tree(self):
        self.tree_tasks.clear()
        for t in config.BRIDGE.list_tasks():
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
        if not hasattr(self, "tree_acts"):
            return
        self.tree_acts.clear()
        for t in config.BRIDGE.list_wax_jobs():
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


    # ------------------------------------------------------------------
    def _create_wax_jobs(self):
        # Для создания наряда достаточно выбранного задания.
        # Проверка ORDERS_POOL мешала создавать наряды для уже существующих
        # заданий, поэтому её убрали.

        if not hasattr(self, "combo_3d_master") or not hasattr(self, "combo_form_master"):
            QMessageBox.warning(self, "Ошибка", "Создание нарядов отключено")
            return

        master_3d = self.combo_3d_master.currentText().strip()
        master_resin = self.combo_form_master.currentText().strip()
        warehouse = self.combo_warehouse.currentText().strip()

        if not master_3d or not master_resin:
            QMessageBox.warning(self, "Ошибка", "Выберите мастеров для обоих методов.")
            return
        if not warehouse:
            QMessageBox.warning(self, "Ошибка", "Выберите склад.")
            return


        if not self.last_created_task_ref:

            QMessageBox.warning(self, "Ошибка", "Нет выбранного задания для создания нарядов.")
            return

        print("[DEBUG] last_created_task_ref =", self.last_created_task_ref)
        norm_type = self.combo_norm_type.currentText().strip()
        result = config.BRIDGE.create_wax_jobs_from_task(
            self.last_created_task_ref,
            master_3d,
            master_resin,
            warehouse,
            norm_type,
        )

        if result:
            QMessageBox.information(self, "Успех", "Созданы наряды: " + ", ".join(result))
            self.refresh()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось создать наряды.")

    # ------------------------------------------------------------------
    def _sync_job(self):
        if not hasattr(self, "tree_jobs"):
            QMessageBox.warning(self, "Ошибка", "Синхронизация отключена")
            return

        code = self._selected_job_code()
        if not code:
            QMessageBox.warning(self, "Наряд", "Не выбран наряд")

        task_num, ok = QInputDialog.getText(self, "Создание нарядов", "Номер задания:")
        if not ok or not task_num:

            return
        warehouse = self.combo_warehouse.currentText().strip()
        norm_type = self.combo_norm_type.currentText().strip()
        try:
            count = config.BRIDGE.create_wax_jobs_from_task(
                task_num,
                self.combo_3d_master.currentText().strip(),
                self.combo_form_master.currentText().strip(),
                warehouse,
                norm_type,
            )
            QMessageBox.information(self, "Готово", f"Создано {len(count)} нарядов")
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


    # —──────────── дерево «Наряды» ─────────────
    def _fill_jobs_tree(self):
        if not hasattr(self, "tree_jobs"):
            return
        self.tree_jobs.clear()

        jobs = config.BRIDGE.list_wax_jobs()
        for job in jobs:
            item = QTreeWidgetItem([
                job["Номер"],
                job["Дата"],
                job["Закрыт"],
                job["Сотрудник"],
                job["ТехОперация"],
                job["Комментарий"],
                job["Склад"],
                job["ПроизводственныйУчасток"],
                job["Организация"],
                job["Задание"],
                job["Ответственный"],
            ])
            item.setCheckState(0, Qt.Unchecked)
            from PyQt5.QtGui import QBrush, QColor
            if job.get("ПометкаУдаления"):
                for i in range(item.columnCount()):
                    item.setBackground(i, QBrush(QColor("#f87171")))
            elif job.get("Проведен"):
                for i in range(item.columnCount()):
                    item.setBackground(i, QBrush(QColor("#bbf7d0")))
            self.tree_jobs.addTopLevelItem(item)

    # —──────────── дерево «Партии» ─────────────
    def _fill_parties_tree(self):
        if not hasattr(self, "tree_part"):
            return
        self.tree_part.clear()

        # Информацию о партиях берём из ORDERS_POOL
        for pack in ORDERS_POOL:
            for b in pack["docs"].get("batches", []):
                root = QTreeWidgetItem(self.tree_part, [
                    f"Партия {b['batch_barcode']}  ({b['metal']} {b['hallmark']} {b['color']})",
                    str(b["qty"]), f"{b['total_w']:.{config.WEIGHT_DECIMALS}f}"
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
                str(d["qty"]), f"{d['weight']:.{config.WEIGHT_DECIMALS}f}"
            ])

    def _on_close_jobs(self):
        tables = [self.tbl_close_3d, self.tbl_close_form]
        job_refs = []
        for tbl in tables:
            for row in range(tbl.rowCount()):
                item = tbl.item(row, 0)
                if item and item.checkState() == Qt.Checked:
                    ref = item.data(Qt.UserRole)
                    if ref and ref not in job_refs:
                        job_refs.append(ref)

        if not job_refs:
            QMessageBox.warning(self, "Ошибка", "Нет выбранных нарядов")
            return

        result = config.BRIDGE.close_wax_jobs(job_refs)
        if result:
            QMessageBox.information(self, "Успех", "Закрыты наряды: " + ", ".join(result))
            self.refresh()
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось закрыть наряды")

    # ------------------------------------------------------------------
    def _add_job_to_assembly(self, job_num: str):
        """Добавляет наряд в очередь сборки ёлок."""
        from logic.state import ASSEMBLY_POOL

        added = False
        for pack in ORDERS_POOL:
            for j in pack["docs"].get("wax_jobs", []):
                if j.get("wax_job") == job_num and j not in ASSEMBLY_POOL:
                    ASSEMBLY_POOL.append(j.copy())
                    log(f"[UI] Добавлен наряд {job_num} в очередь сборки")
                    added = True
                    break
            if added:
                break

        if not added:
            for j in WAX_JOBS_POOL:
                if j.get("wax_job") == job_num and j not in ASSEMBLY_POOL:
                    ASSEMBLY_POOL.append(j.copy())
                    log(f"[UI] Добавлен наряд {job_num} в очередь сборки (поиск в WAX_JOBS_POOL)")
                    added = True
                    break

        if not added:
            log(f"[UI] ❌ Наряд {job_num} не найден для добавления в сборку")

        self._fill_assembly_tree()

    def _fill_assembly_tree(self):
        if not hasattr(self, "tree_assembly"):
            return
        from logic.state import ASSEMBLY_POOL
        self.tree_assembly.clear()

        grouped = defaultdict(lambda: dict(qty=0, weight=0, jobs=[]))
        for j in ASSEMBLY_POOL:
            key = (j.get("metal"), j.get("hallmark"), j.get("color"))
            grouped[key]["qty"] += j.get("qty", 0)
            grouped[key]["weight"] += j.get("weight", 0)
            grouped[key]["jobs"].append(j["wax_job"])

        for (metal, hallmark, color), data in grouped.items():
            root = QTreeWidgetItem(
                self.tree_assembly,
                [f"{metal} {hallmark} {color}", str(data["qty"]),
                 f"{data['weight']:.{config.WEIGHT_DECIMALS}f}"]
            )
            root.setExpanded(True)
            for j in [x for x in ASSEMBLY_POOL if (x.get("metal"), x.get("hallmark"), x.get("color")) == (metal, hallmark, color)]:
                QTreeWidgetItem(root, [j["wax_job"], str(j.get("qty", 0)), f"{j.get('weight', 0):.{config.WEIGHT_DECIMALS}f}"])

    def _clear_assembly_pool(self):
        from logic.state import ASSEMBLY_POOL, TREES_POOL
        ASSEMBLY_POOL.clear()
        TREES_POOL.clear()
        self._fill_assembly_tree()

    def _form_trees(self, return_to_jobs: bool = True):
        """Формирует ёлки из собранных нарядов.

        :param return_to_jobs: после формирования вернуться к списку нарядов
        """
        from logic.state import ASSEMBLY_POOL
        if not ASSEMBLY_POOL:
            QMessageBox.information(self, "Сборка", "Нет нарядов для сборки")
            return
        from logic import production_docs

        trees = production_docs.form_wax_trees(ASSEMBLY_POOL)
        ASSEMBLY_POOL.clear()
        self._fill_assembly_tree()

        count = len(trees)
        log(f"[UI] Сформировано {count} ёлок")
        if count:
            tree_codes = ", ".join(t["tree_code"] for t in trees)
            QMessageBox.information(
                self,
                "Сборка ёлок",
                f"Сформировано {count} ёлок: {tree_codes}",
            )
        else:
            QMessageBox.information(self, "Сборка ёлок", "Ёлки не сформированы")
        if return_to_jobs and hasattr(self, "tabs_jobs"):
            # после формирования возвращаемся к выбору нарядов
            self.tabs_jobs.setCurrentIndex(2)

    # ------------------------------------------------------------------
    def _not_implemented(self):
        QMessageBox.information(self, "Информация", "Функционал в разработке")

