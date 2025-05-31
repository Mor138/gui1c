from PyQt5.QtWidgets import QWidget, QVBoxLayout
from form_renderer import render_form

class CastingInPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        form_widget = render_form("casting_in")
        layout.addWidget(form_widget)
        self.setLayout(layout)
