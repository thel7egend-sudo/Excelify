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
        row2_layout.setContentsMargins(16, 10, 16, 10)
        row2_layout.setSpacing(0)

        self.search = QLineEdit()
        self.search.setObjectName("searchField")
        self.search.setPlaceholderText("Search documents")
        self.search.setFixedWidth(420)
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
                background-color: #ffffff;
                border-bottom: 1px solid #e5e7eb;
            }

            QWidget#topChrome > QWidget {
                background-color: #ffffff;
            }

            QWidget#searchRow {
                background-color: #ffffff;
            }

            QWidget#searchRow > QWidget {
                background-color: #ffffff;
            }

            QLabel#homeLabel {
                background-color: #f9fafb;
                color: #111827;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 14px;
                font-weight: 600;
                border: 1px solid #e5e7eb;
            }

            QLabel#homeLabel:hover {
                background-color: #f3f4f6;
                border: 1px solid #d1d5db;
            }

            QLineEdit#searchField {
                background-color: #ffffff;
                color: #111827;
                border: 1px solid #c9ced6;
                padding: 9px 12px;
                border-radius: 9px;
                margin: 0px;
            }

            QLineEdit#searchField:focus {
                border: 2px solid #256d85;
                padding: 8px 11px;
            }

            QLineEdit#searchField::placeholder {
                color: #6b7280;
            }

            QListWidget#searchResults {
                background-color: #ffffff;
                color: #111827;
                border: 1px solid #e5e7eb;
            }

            QListWidget#searchResults::viewport {
                background-color: #ffffff;
            }

            QListWidget#searchResults::item {
                padding: 6px 10px;
                background-color: transparent;
            }

            QListWidget#searchResults::item:selected {
                background-color: #e8f2f5;
                color: #0f172a;
            }
            """)
            return

        self.setStyleSheet("""
        QWidget#topChrome {
            background-color: #111827;
            border-bottom: 1px solid #374151;
        }

        QWidget#topChrome > QWidget {
            background-color: #111827;
        }

        QWidget#searchRow {
            background-color: #111827;
        }

        QWidget#searchRow > QWidget {
            background-color: #111827;
        }

        QLabel#homeLabel {
            background-color: #1f2937;
            color: #e5e7eb;
            padding: 6px 14px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            border: 1px solid #374151;
        }

        QLabel#homeLabel:hover {
            background-color: #273449;
            border: 1px solid #4b5563;
        }

        QLineEdit#searchField {
            background-color: #1f2937;
            color: #e5e7eb;
            border: 1px solid #4b5563;
            padding: 9px 12px;
            border-radius: 9px;
            margin: 0px;
        }

        QLineEdit#searchField:focus {
            border: 2px solid #256d85;
            padding: 8px 11px;
        }

        QLineEdit#searchField::placeholder {
            color: #9ca3af;
        }

        QListWidget#searchResults {
            background-color: #1f2937;
            color: #e5e7eb;
            border: 1px solid #374151;
        }

        QListWidget#searchResults::viewport {
            background-color: #1f2937;
        }

        QListWidget#searchResults::item {
            padding: 6px 10px;
            background-color: transparent;
        }

        QListWidget#searchResults::item:selected {
            background-color: #1f4250;
            color: #ffffff;
        }
        """)
