from PySide6.QtCore import Qt, Signal, QSignalBlocker
from PySide6.QtGui import QColor, QTextDocument
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


class DocEditorPage(QWidget):
    document_changed = Signal()

    PAGE_WIDTH = 800
    PAGE_HEIGHT = 1100
    PAGE_MARGIN = 40

    def __init__(self, document):
        super().__init__()
        self.document = document
        self.pages = []
        self.text_edits = []
        self._syncing_pages = False

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
        self.pages_layout.setContentsMargins(0, 28, 0, 28)
        self.pages_layout.setSpacing(48)
        self.pages_layout.setAlignment(Qt.AlignHCenter | Qt.AlignTop)

        self.scroll.setWidget(self.scroll_content)

        layout.addWidget(self.ribbon)
        layout.addWidget(self.scroll)

        self._set_page_texts(self._paginate_text(self.document.content))
        self.apply_grid_dark_mode(False)

    def _create_page(self):
        page = QWidget()
        page.setObjectName("docPage")
        page.setFixedSize(self.PAGE_WIDTH, self.PAGE_HEIGHT)

        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(
            self.PAGE_MARGIN,
            self.PAGE_MARGIN,
            self.PAGE_MARGIN,
            self.PAGE_MARGIN,
        )

        text_edit = QTextEdit()
        text_edit.setFrameShape(QFrame.NoFrame)
        text_edit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        text_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        text_edit.textChanged.connect(self._on_any_page_text_changed)
        page_layout.addWidget(text_edit)

        shadow = QGraphicsDropShadowEffect(page)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 3)
        shadow.setColor(QColor(0, 0, 0, 35))
        page.setGraphicsEffect(shadow)

        self.pages_layout.addWidget(page)
        self.pages.append(page)
        self.text_edits.append(text_edit)

    def _ensure_page_count(self, count):
        while len(self.text_edits) < count:
            self._create_page()

        while len(self.text_edits) > count:
            edit = self.text_edits.pop()
            page = self.pages.pop()
            edit.deleteLater()
            self.pages_layout.removeWidget(page)
            page.deleteLater()

    def _on_any_page_text_changed(self):
        if self._syncing_pages:
            return

        full_text = "".join(edit.toPlainText() for edit in self.text_edits)
        page_texts = self._paginate_text(full_text)
        self._set_page_texts(page_texts)
        self.document.content = "".join(page_texts)
        self.document_changed.emit()

    def _paginate_text(self, text):
        if not text:
            return [""]

        probe = QTextDocument()
        probe.setDocumentMargin(0)
        probe.setTextWidth(self.PAGE_WIDTH - (self.PAGE_MARGIN * 2))

        raw_lines = text.splitlines(keepends=True)
        if not raw_lines:
            raw_lines = [text]

        pages = []
        current = ""
        max_height = self.PAGE_HEIGHT - (self.PAGE_MARGIN * 2)

        for line in raw_lines:
            candidate = current + line
            probe.setPlainText(candidate)
            if probe.size().height() <= max_height:
                current = candidate
                continue

            if current:
                pages.append(current)
                current = line
            else:
                pages.append(line)
                current = ""

        pages.append(current)
        return pages or [""]

    def _set_page_texts(self, page_texts):
        self._syncing_pages = True
        self._ensure_page_count(max(1, len(page_texts)))

        for i, edit in enumerate(self.text_edits):
            text = page_texts[i] if i < len(page_texts) else ""
            with QSignalBlocker(edit):
                edit.setPlainText(text)

        self._syncing_pages = False

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
                    background: #dfe1e5;
                    border: none;
                }
                QWidget#docPage {
                    background: #ffffff;
                    border: 1px solid #e5e7eb;
                    border-radius: 4px;
                }
                QWidget#docPage QTextEdit {
                    background: #ffffff;
                    color: #111827;
                    border: none;
                    font-size: 14px;
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
                background: #1b1c1e;
                border: none;
            }
            QWidget#docPage {
                background: #2b2c2f;
                border: 1px solid rgba(255, 255, 255, 0.10);
                border-radius: 4px;
            }
            QWidget#docPage QTextEdit {
                background: #2b2c2f;
                color: #f3f4f6;
                border: none;
                font-size: 14px;
            }
        """)
