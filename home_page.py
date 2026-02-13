from PySide6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QScrollArea, QGridLayout, QPushButton, QFrame
from document_card import DocumentCard
from document import Document
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QMessageBox
class HomePage(QWidget):
    open_document_requested = Signal(object)
    open_import_requested = Signal()
    save_requested = Signal()
    toggle_dark_mode_requested = Signal()

    def __init__(self, chrome):
        super().__init__()
        self.chrome = chrome
        self.chrome.search.textChanged.connect(self.on_search_text)
        self.chrome.search.returnPressed.connect(self.on_search_enter)
        self.chrome.search_selected.connect(self.open_document_by_name)

        self._dark_mode = False


        self.documents = []
        self.cards = {}

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # LEFT COLUMN (import area)
        left = QWidget()
        left.setObjectName("homeLeft")
        left.setFixedWidth(160)

        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(8)

        # spacer pushes button to bottom
        left_layout.addStretch()

        # dark mode button (UI only for now)
        self.dark_mode_btn = QPushButton("Dark Mode")
        self.dark_mode_btn.setCheckable(True)

        self.dark_mode_btn.setFixedHeight(36)
        self.dark_mode_btn.toggled.connect(
            lambda checked: self.toggle_dark_mode_requested.emit()
        )


        left_layout.addWidget(self.dark_mode_btn)


        # import button (UI only for now)
        self.import_btn = QPushButton("Import from Excel")
        self.import_btn.clicked.connect(self.request_import)

        self.import_btn.setFixedHeight(36)

        left_layout.addWidget(self.import_btn)




        # RIGHT DOCUMENT AREA
        self.grid = QGridLayout()
        self.grid.setContentsMargins(24, 24, 24, 24)
        self.grid.setHorizontalSpacing(24)
        self.grid.setVerticalSpacing(24)

        self.plus_btn = QPushButton("+")
        self.plus_btn.setFixedSize(160, 120)
        self.plus_btn.setObjectName("plusButton")
        self.plus_btn.clicked.connect(self.create_document)

        self.grid.addWidget(self.plus_btn, 0, 0)

        container = QWidget()
        container.setObjectName("homeContainer")
        container.setLayout(self.grid)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(container)
        scroll.setObjectName("homeScroll")
        scroll.setFrameShape(QFrame.NoFrame)

        layout.addWidget(left)
        layout.addWidget(scroll)
        self.apply_dark_mode(False)

    def create_document(self):
        doc = Document(f"Untitled {len(self.documents)+1}")
        self.documents.append(doc)

        card = DocumentCard(doc)
        card.apply_dark_mode(self._dark_mode)
        card.clicked.connect(lambda d=doc: self.open_document_requested.emit(d))
        card.rename_requested.connect(self.sync_rename)
        card.delete_requested.connect(self.request_delete)   # ✅ ADD THIS
        self.cards[doc] = card

        row = len(self.documents) // 4
        col = len(self.documents) % 4
        self.grid.addWidget(card, row, col)
        self.save_requested.emit()

    def sync_rename(self, document):
    # update the card text
        self.cards[document].update_name()
        self.save_requested.emit()
    def on_search_text(self, text):
        self.chrome.update_search_results(self.documents, text)


    def on_search_enter(self):
        if not self.documents:
            return

        text = self.chrome.search.text().lower()
        for doc in self.documents:
            if text in doc.name.lower():
                self.open_document_requested.emit(doc)
                self.chrome.search.clear()
                self.chrome.search_results.setVisible(False)
                break


    def open_document_by_name(self, name):
        for doc in self.documents:
            if doc.name == name:
                self.open_document_requested.emit(doc)
                break
    def add_existing_document(self, doc):
        card = DocumentCard(doc)
        card.apply_dark_mode(self._dark_mode)
        card.rename_requested.connect(self.sync_rename)
        card.delete_requested.connect(self.request_delete)
        card.clicked.connect(lambda d=doc: self.open_document_requested.emit(d))

        self.cards[doc] = card

        index = len(self.documents)
        row = index // 4
        col = index % 4
        self.grid.addWidget(card, row, col)

    def request_delete(self, document):
        reply = QMessageBox.question(
            self,
            "Delete Document",
            "Are you sure?\n\nThis will permanently delete this document.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # remove from internal list
        self.documents.remove(document)

        # remove card widget
        card = self.cards.pop(document)
        self.grid.removeWidget(card)
        card.deleteLater()

        # always put + button first
        self.grid.addWidget(self.plus_btn, 0, 0)

        # reflow documents AFTER the +
        for i, doc in enumerate(self.documents):
            index = i + 1  # offset because + occupies slot 0
            row = index // 4
            col = index % 4
            self.grid.addWidget(self.cards[doc], row, col)

        self.save_requested.emit()

    def request_import(self):
        self.open_import_requested.emit()

    def apply_dark_mode(self, enabled: bool):
        self._dark_mode = enabled

        if not enabled:
            self.setStyleSheet("""
            QWidget#homeLeft {
                background: #f7f8fa;
                border-right: 1px solid #e5e7eb;
            }

            QWidget#homeContainer {
                background: #f7f8fa;
            }

            QScrollArea#homeScroll {
                background: #f7f8fa;
                border: none;
            }

            QScrollArea#homeScroll::corner {
                background: #f7f8fa;
            }

            QScrollArea#homeScroll > QWidget {
                background: #f7f8fa;
            }

            QScrollArea#homeScroll > QWidget > QWidget {
                background: #f7f8fa;
            }

            QPushButton {
                background: #ffffff;
                color: #111827;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                padding: 0 12px;
                font-weight: 500;
            }

            QPushButton:hover {
                background: #f9fafb;
                border: 1px solid #d1d5db;
            }

            QPushButton:checked {
                background: #256d85;
                border: 1px solid #256d85;
                color: #ffffff;
            }

            QPushButton#plusButton {
                border: 1px dashed #d1d5db;
                border-radius: 10px;
                font-size: 34px;
                color: #6b7280;
                background: #ffffff;
            }

            QScrollBar:vertical {
                background: #f3f4f6;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #d1d5db;
                min-height: 30px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #9ca3af;
            }
            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {
                background: #f3f4f6;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
            }
            """)
        else:
            self.setStyleSheet("""
            QWidget#homeLeft {
                background: #111827;
                border-right: 1px solid #374151;
            }

            QWidget#homeContainer {
                background: #111827;
            }

            QScrollArea#homeScroll {
                background: #111827;
                border: none;
            }

            QScrollArea#homeScroll::corner {
                background: #111827;
            }

            QScrollArea#homeScroll > QWidget {
                background: #111827;
            }

            QScrollArea#homeScroll > QWidget > QWidget {
                background: #111827;
            }

            QPushButton {
                background: #1f2937;
                color: #e5e7eb;
                border: 1px solid #374151;
                border-radius: 8px;
                padding: 0 12px;
                font-weight: 500;
            }

            QPushButton:hover {
                background: #273449;
                border: 1px solid #4b5563;
            }

            QPushButton:checked {
                background: #256d85;
                border: 1px solid #256d85;
                color: #ffffff;
            }

            QPushButton#plusButton {
                border: 1px dashed #4b5563;
                border-radius: 10px;
                font-size: 34px;
                background: #1f2937;
                color: #9ca3af;
            }

            QScrollBar:vertical {
                background: #111827;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #374151;
                min-height: 30px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #4b5563;
            }
            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {
                background: #111827;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
            }
            """)

        self.plus_btn.setObjectName("plusButton")
        self.plus_btn.style().unpolish(self.plus_btn)
        self.plus_btn.style().polish(self.plus_btn)

        for card in self.cards.values():
            card.apply_dark_mode(enabled)
