from PyQt5.QtWidgets import QWidget, QVBoxLayout
from form_renderer import render_form

class DynamicFormPage(QWidget):
    def __init__(self, tab_key):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.addWidget(render_form(tab_key))