from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout
from PySide6.QtCore import Qt, Signal
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

        self.setFixedSize(160, 120)
        self.setStyleSheet("""
            QFrame {
                border-radius: 8px;
                border: 1px solid #bdbdbd;
            }
            QFrame:hover {
                border: 1px solid #8f8f8f;
            }
        """)


        self.label = QLabel(document.name)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setWordWrap(True)

        layout = QVBoxLayout(self)
        layout.addWidget(self.label)

    def update_name(self):
        self.label.setText(self.document.name)

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

   
