from PySide6.QtWidgets import (
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtCore import Signal


class DocEditorPage(QWidget):
    document_changed = Signal()

    def __init__(self, document):
        super().__init__()
        self.document = document
        self.pages = []
        self.text_edits = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.ribbon = QWidget()
        self.ribbon.setObjectName("docRibbon")
        self.ribbon.setFixedHeight(72)
        ribbon_layout = QHBoxLayout(self.ribbon)
        ribbon_layout.setContentsMargins(16, 0, 16, 0)

        placeholder_btn = QPushButton("Tools (Coming Soon)")
        placeholder_btn.setEnabled(False)
        placeholder_btn.setFixedHeight(32)
        ribbon_layout.addWidget(placeholder_btn)
        ribbon_layout.addStretch()

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setObjectName("docScroll")

        self.scroll_content = QWidget()
        self.scroll_content.setObjectName("docScrollContent")
        self.pages_layout = QVBoxLayout(self.scroll_content)
        self.pages_layout.setContentsMargins(0, 24, 0, 24)
        self.pages_layout.setSpacing(24)
        self.pages_layout.setAlignment(Qt.AlignHCenter | Qt.AlignTop)

        self.scroll.setWidget(self.scroll_content)

        layout.addWidget(self.ribbon)
        layout.addWidget(self.scroll)

        self.add_page()
        self._load_document_content()
        self.apply_grid_dark_mode(False)

    def add_page(self):
        page = QWidget()
        page.setObjectName("docPage")
        page.setFixedWidth(800)
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(28, 28, 28, 28)

        text_edit = QTextEdit()
        text_edit.textChanged.connect(self._sync_document_content)
        page_layout.addWidget(text_edit)

        shadow = QGraphicsDropShadowEffect(page)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 3)
        shadow.setColor(QColor(0, 0, 0, 35))
        page.setGraphicsEffect(shadow)

        self.pages_layout.addWidget(page)
        self.pages.append(page)
        self.text_edits.append(text_edit)

    def _load_document_content(self):
        if self.text_edits:
            self.text_edits[0].setPlainText(self.document.content)

    def _sync_document_content(self):
        if self.text_edits:
            self.document.content = self.text_edits[0].toPlainText()
            self.document_changed.emit()

    def apply_grid_dark_mode(self, enabled: bool):
        if not enabled:
            self.setStyleSheet("""
                QWidget {
                    background: #f7f8fa;
                    color: #111827;
                }
                QWidget#docRibbon {
                    background-color: #f9fafb;
                    border-bottom: 1px solid #e5e7eb;
                }
                QScrollArea#docScroll,
                QWidget#docScrollContent {
                    background: #f7f8fa;
                    border: none;
                }
                QWidget#docPage {
                    background: #ffffff;
                    border: 1px solid #e5e7eb;
                    border-radius: 6px;
                }
                QWidget#docPage QTextEdit {
                    background: #ffffff;
                    color: #111827;
                    border: none;
                    font-size: 14px;
                    line-height: 1.5;
                }
            """)
            return

        self.setStyleSheet("""
            QWidget {
                background: #202124;
                color: #eaeaea;
            }
            QWidget#docRibbon {
                background-color: #252525;
                border-bottom: 1px solid rgba(255, 255, 255, 0.06);
            }
            QScrollArea#docScroll,
            QWidget#docScrollContent {
                background: #202124;
                border: none;
            }
            QWidget#docPage {
                background: #2a2a2a;
                border: 1px solid rgba(255, 255, 255, 0.10);
                border-radius: 6px;
            }
            QWidget#docPage QTextEdit {
                background: #2a2a2a;
                color: #f3f4f6;
                border: none;
                font-size: 14px;
                line-height: 1.5;
            }
        """)
