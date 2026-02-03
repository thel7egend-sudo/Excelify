from PySide6.QtWidgets import QTabBar, QMenu, QInputDialog
from PySide6.QtCore import Qt, Signal

class DocumentTabBar(QTabBar):
    rename_requested = Signal(int)

    def __init__(self):
        super().__init__()
        self.setMovable(True)
        self.setTabsClosable(True)

        self.setStyleSheet("""
        QTabBar::tab {
            background: #d9d9d9;
            padding: 6px 14px;
            margin-right: 6px;
            border-radius: 8px;
            font-size: 13px;
        }
        QTabBar::tab:selected {
            background: #c2c2c2;
        }
        """)

        self.tabCloseRequested.connect(self.removeTab)

    def contextMenuEvent(self, event):
        index = self.tabAt(event.pos())
        if index < 0:
            return

        menu = QMenu(self)
        rename_action = menu.addAction("Rename")

        if menu.exec(event.globalPos()) == rename_action:
            self.rename_requested.emit(index)
