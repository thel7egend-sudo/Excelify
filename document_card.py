from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout, QGraphicsDropShadowEffect
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QMenu, QInputDialog
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QMenu, QInputDialog, QFrame, QLabel, QVBoxLayout
from PySide6.QtCore import Signal

class DocumentCard(QFrame):
    clicked = Signal(object)
    rename_requested = Signal(object)
    delete_requested = Signal(object)
    def __init__(self, document):
        super().__init__()
        self.document = document

        self.setFixedSize(176, 132)

        self.label = QLabel(document.name)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setWordWrap(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.addWidget(self.label)

        self._shadow_normal = (0, 2, 6, 22)
        self._shadow_hover = (0, 4, 10, 30)
        self._set_shadow(*self._shadow_normal)

        self.apply_dark_mode(False)

    def update_name(self):
        self.label.setText(self.document.name)

    def apply_dark_mode(self, enabled: bool):
        if not enabled:
            self._shadow_normal = (0, 2, 6, 22)
            self._shadow_hover = (0, 4, 10, 30)
            self._set_shadow(*self._shadow_normal)

            self.setStyleSheet("""
                QFrame {
                    background: #ffffff;
                    border-radius: 10px;
                    border: 1px solid #e5e7eb;
                }
                QFrame:hover {
                    border: 1px solid #d1d5db;
                    background: #ffffff;
                }
            """)
            self.label.setStyleSheet("color: #111827; font-size: 14px; font-weight: 500;")
            return

        self._shadow_normal = (0, 2, 6, 80)
        self._shadow_hover = (0, 4, 10, 110)
        self._set_shadow(*self._shadow_normal)

        self.setStyleSheet("""
            QFrame {
                background: #252525;
                border-radius: 10px;
                border: 1px solid rgba(255, 255, 255, 0.06);
            }
            QFrame:hover {
                border: 1px solid rgba(255, 255, 255, 0.10);
                background: #2e2e2e;
            }
        """)
        self.label.setStyleSheet("color: #eaeaea; font-size: 14px; font-weight: 500;")

    def _set_shadow(self, x_offset: int, y_offset: int, blur: int, alpha: int):
        effect = QGraphicsDropShadowEffect(self)
        effect.setOffset(x_offset, y_offset)
        effect.setBlurRadius(blur)
        effect.setColor(QColor(0, 0, 0, alpha))
        self.setGraphicsEffect(effect)

    def enterEvent(self, event):
        self._set_shadow(*self._shadow_hover)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._set_shadow(*self._shadow_normal)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.document)
        else:
            super().mousePressEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #f5f5f5;
                border: 1px solid #bdbdbd;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #333333;
            }
            QMenu::item:selected {
                background-color: #dcdcdc;
        """)

        rename_action = menu.addAction("Rename")
        delete_action = menu.addAction("Delete")

        action = menu.exec(event.globalPos())

        if action == rename_action:
            new_name, ok = QInputDialog.getText(
                self,
                "Rename Document",
                "New name:",
                text=self.document.name
            )
            if ok and new_name.strip():
                self.document.name = new_name.strip()
                self.update_name()
                self.rename_requested.emit(self.document)

        elif action == delete_action:
            self.delete_requested.emit(self.document)

   
