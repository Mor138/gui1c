from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem

class CatalogTable(QWidget):
    def __init__(self, data: list[dict], columns: list[str]):
        super().__init__()
        layout = QVBoxLayout(self)
        table = QTableWidget(len(data), len(columns))
        table.setHorizontalHeaderLabels(columns)
        for row, item in enumerate(data):
            for col, col_name in enumerate(columns):
                value = item.get(col_name, "")
                table.setItem(row, col, QTableWidgetItem(str(value)))
        layout.addWidget(table)
