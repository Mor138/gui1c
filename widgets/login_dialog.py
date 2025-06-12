from PyQt5.QtWidgets import QDialog, QVBoxLayout, QFormLayout, QLineEdit, QDialogButtonBox
from PyQt5.QtCore import Qt


class LoginDialog(QDialog):
    """Простое диалоговое окно для ввода учётных данных 1С."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Вход в 1С")
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        form = QFormLayout()
        self.ed_user = QLineEdit(); self.ed_user.setPlaceholderText("Пользователь")
        self.ed_pass = QLineEdit(); self.ed_pass.setEchoMode(QLineEdit.Password)
        form.addRow("Пользователь", self.ed_user)
        form.addRow("Пароль", self.ed_pass)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_credentials(self):
        return self.ed_user.text().strip(), self.ed_pass.text()
