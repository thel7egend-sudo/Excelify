from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QListWidget, QFrame, QSizePolicy
)

from PySide6.QtCore import Qt, Signal

class TopChrome(QWidget):
    home_clicked = Signal()
    search_selected = Signal(object)

    def __init__(self):
        super().__init__()
        self.setObjectName("topChrome")


        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ---------- ROW 1: HOME + TABS ----------
        row1 = QWidget()
        row1_layout = QHBoxLayout(row1)
        row1_layout.setContentsMargins(16, 8, 16, 8)
        row1_layout.setSpacing(0)

        self.back_label = QLabel("Home")
        self.back_label.setObjectName("homeLabel")

        self.back_label.setCursor(Qt.PointingHandCursor)

        def home_click(event):
            self.home_clicked.emit()

        self.back_label.mousePressEvent = home_click


        row1_layout.addWidget(self.back_label)
        row1_layout.addStretch()

        # ---------- ROW 2: SEARCH ----------
        row2 = QWidget()
        row2.setObjectName("searchRow")
        row2_layout = QHBoxLayout(row2)
        row2_layout.setContentsMargins(16, 6, 16, 6)
        row2_layout.setSpacing(0)

        self.search = QLineEdit()
        self.search.setObjectName("searchField")
        self.search.setPlaceholderText("Search documents")
        self.search.setFixedWidth(360)
        self.search_results = QListWidget()
        self.search_results.setObjectName("searchResults")
        self.search_results.setAutoFillBackground(True)
        self.search_results.setFrameShape(QFrame.NoFrame)
        self.search_results.setVisible(False)
        self.search_results.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.search_results.itemClicked.connect(self._emit_search_result)

        row2_layout.addStretch()
        row2_layout.addWidget(self.search)
        row2_layout.addStretch()

        layout.addWidget(row1)
        layout.addWidget(row2)
        layout.addWidget(self.search_results)
        self.apply_dark_mode(False)
    def update_search_results(self, documents, text):
        self.search_results.clear()

        if not text.strip():
            self.search_results.setVisible(False)
            return

        for doc in documents:
            if text.lower() in doc.name.lower():
                self.search_results.addItem(doc.name)

        self.search_results.setVisible(self.search_results.count() > 0)


    def _emit_search_result(self, item):
        self.search_selected.emit(item.text())
        self.search_results.setVisible(False)
    def show_home_mode(self):
        self.search.show()
        self.search_results.hide()


    def show_editor_mode(self):
        self.search.hide()
        self.search_results.hide()
    def apply_dark_mode(self, enabled: bool):
        if not enabled:
            self.setStyleSheet("""
            QWidget#topChrome {
                background-color: #d2d0c7;
                border-bottom: 1px solid #bbb8ad;
            }

            QWidget#topChrome > QWidget {
                background-color: #d2d0c7;
            }

            QWidget#searchRow {
                background-color: #d2d0c7;
            }

            QWidget#searchRow > QWidget {
                background-color: #d2d0c7;
            }

            QLabel#homeLabel {
                background-color: #f2f2f2;
                color: #232323;
                padding: 6px 14px;
                border-radius: 10px;
                font-weight: 600;
                border: 1px solid #b7b7b7;
            }

            QLabel#homeLabel:hover {
                background-color: #e6e6e6;
            }

            QLineEdit#searchField {
                background-color: #efefef;
                color: #222222;
                border: 1px solid #adadad;
                padding: 6px;
                border-radius: 6px;
                margin: 0px;
            }

            QLineEdit#searchField:focus {
                background-color: #f8f8f8;
                border: 1px solid #8f8f8f;
            }

            QListWidget#searchResults {
                background-color: #ececec;
                color: #222222;
                border: 1px solid #c2c2c2;
            }

            QListWidget#searchResults::viewport {
                background-color: #ececec;
            }

            QListWidget#searchResults::item {
                padding: 6px 10px;
                background-color: transparent;
            }

            QListWidget#searchResults::item:selected {
                background-color: #dcebdc;
            }
            """)
            return

        self.setStyleSheet("""
        QWidget#topChrome {
            background-color: #1e1e1e;
            border-bottom: 1px solid #3a3a3a;
        }

        QWidget#topChrome > QWidget {
            background-color: #1e1e1e;
        }

        QWidget#searchRow {
            background-color: #252526;
        }

        QWidget#searchRow > QWidget {
            background-color: #252526;
        }

        QLabel#homeLabel {
            background-color: #2d2d30;
            color: #e6e6e6;
            padding: 6px 14px;
            border-radius: 10px;
            font-weight: 600;
        }

        QLabel#homeLabel:hover {
            background-color: #3a3a3a;
        }

        QLineEdit#searchField {
            background-color: #1f2124;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            padding: 6px;
            border-radius: 6px;
            margin: 0px;
        }

        QLineEdit#searchField:focus {
            background-color: #1f2124;
            border: 1px solid #4a4f55;
        }

        QListWidget#searchResults {
            background-color: #1f2124;
            color: #e6e6e6;
            border: 1px solid #2f3237;
        }

        QListWidget#searchResults::viewport {
            background-color: #1f2124;
        }

        QListWidget#searchResults::item {
            padding: 6px 10px;
            background-color: transparent;
        }

        QListWidget#searchResults::item:selected {
            background-color: #2b2f35;
        }
        """)
    
