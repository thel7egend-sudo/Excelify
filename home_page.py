from PySide6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QScrollArea, QGridLayout, QPushButton
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


        self.documents = []
        self.cards = {}

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # LEFT COLUMN (import area)
        left = QWidget()
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
        self.import_btn.setStyleSheet("""
            QPushButton {
                background:#f5f5f5;
                border:1px solid #bdbdbd;
                border-radius:6px;
                padding:0 8px;
                font-weight:500;
            }
            QPushButton:hover {
                background:#ebebeb;
            }
        """)

        left_layout.addWidget(self.import_btn)




        # RIGHT DOCUMENT AREA
        self.grid = QGridLayout()
        self.grid.setSpacing(16)

        self.plus_btn = QPushButton("+")
        self.plus_btn.setFixedSize(160, 120)
        self.plus_btn.setStyleSheet("""
            QPushButton {
                border:2px dashed #aaa;
                border-radius:8px;
                font-size:36px;
                background:#f2f2f2;
            }
        """)
        self.plus_btn.clicked.connect(self.create_document)

        self.grid.addWidget(self.plus_btn, 0, 0)

        container = QWidget()
        container.setLayout(self.grid)
        container.setStyleSheet("background:#f2f2f2;")

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(container)
        scroll.setStyleSheet("""
            QScrollBar:vertical {
                background: #e0e0e0;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #9e9e9e;
                min-height: 30px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #7f7f7f;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)

        layout.addWidget(left)
        layout.addWidget(scroll)

    def create_document(self):
        doc = Document(f"Untitled {len(self.documents)+1}")
        self.documents.append(doc)

        card = DocumentCard(doc)
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
