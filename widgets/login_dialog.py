from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QDialogButtonBox, QComboBox,
    QLabel
)
from PyQt5.QtCore import Qt
import config


class LoginDialog(QDialog):
    """Простое диалоговое окно для ввода учётных данных 1С."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Вход в 1С")
        self.setStyleSheet(config.LOGIN_DIALOG_CSS)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("Рост Золото")
        title.setObjectName("title")
        layout.addWidget(title, alignment=Qt.AlignHCenter)

        form = QFormLayout()
        self.c_user = QComboBox(); self.c_user.setEditable(True)
        self.c_user.addItems(config.EMPLOYEE_LOGINS)
        self.ed_pass = QLineEdit(); self.ed_pass.setEchoMode(QLineEdit.Password)
        form.addRow("Пользователь", self.c_user)
        form.addRow("Пароль", self.ed_pass)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_credentials(self):
        return self.c_user.currentText().strip(), self.ed_pass.text()
