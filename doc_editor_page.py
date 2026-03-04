from docx import Document as DocxDocument
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
    QFileDialog,
    QMessageBox,
)


class DocEditorPage(QWidget):
    document_changed = Signal()
    export_requested = Signal(object)

    PAGE_WIDTH = 800
    PAGE_HEIGHT = 1100
    PAGE_MARGIN = 40

    @property
    def _usable_page_width(self):
        return self.PAGE_WIDTH - (self.PAGE_MARGIN * 2)

    @property
    def _usable_page_height(self):
        return self.PAGE_HEIGHT - (self.PAGE_MARGIN * 2)

    def __init__(self, document):
        super().__init__()
        self.document = document
        self.pages = []
        self.text_edits = []
        self._page_texts = [self.document.content or ""]
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

        self.export_btn = QPushButton("Export to Docs")
        self.export_btn.setFixedHeight(36)
        self.export_btn.clicked.connect(lambda: self.export_requested.emit(self.document))
        ribbon_layout.addWidget(self.export_btn)

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
        text_edit.setFixedSize(self._usable_page_width, self._usable_page_height)
        text_edit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        text_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        text_edit.document().setDocumentMargin(0)
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

        sender = self.sender()
        sender_index = self.text_edits.index(sender) if sender in self.text_edits else None

        full_text, global_cursor_pos = self._build_full_text_from_change(sender_index)
        page_texts = self._paginate_text(full_text)
        self._set_page_texts(page_texts, global_cursor_pos)
        self.document.content = "".join(page_texts)
        self.document_changed.emit()

    def _build_full_text_from_change(self, sender_index):
        if sender_index is None:
            return "".join(edit.toPlainText() for edit in self.text_edits), None

        old_page_texts = self._page_texts or ["" for _ in self.text_edits]
        old_full_text = "".join(old_page_texts)

        start_offset = sum(len(chunk) for chunk in old_page_texts[:sender_index])
        old_chunk_length = len(old_page_texts[sender_index]) if sender_index < len(old_page_texts) else 0
        new_chunk = self.text_edits[sender_index].toPlainText()

        full_text = (
            old_full_text[:start_offset]
            + new_chunk
            + old_full_text[start_offset + old_chunk_length :]
        )

        local_cursor = self.text_edits[sender_index].textCursor().position()
        global_cursor_pos = start_offset + local_cursor
        return full_text, global_cursor_pos

    def _paginate_text(self, text):
        if not text:
            return [""]

        probe = QTextDocument()
        probe.setDocumentMargin(0)
        if self.text_edits:
            probe.setDefaultFont(self.text_edits[0].font())
        probe.setTextWidth(self._usable_page_width)
        max_height = self._usable_page_height
        pages = []
        remaining = text

        while remaining:
            probe.setPlainText(remaining)
            if probe.documentLayout().documentSize().height() <= max_height:
                pages.append(remaining)
                break

            low = 1
            high = len(remaining)
            fit = 1

            while low <= high:
                mid = (low + high) // 2
                chunk = remaining[:mid]
                probe.setPlainText(chunk)
                if probe.documentLayout().documentSize().height() <= max_height:
                    fit = mid
                    low = mid + 1
                else:
                    high = mid - 1

            split_at = fit
            while split_at > 1 and not remaining[split_at - 1].isspace():
                split_at -= 1

            if split_at <= 1:
                split_at = fit

            pages.append(remaining[:split_at])
            remaining = remaining[split_at:]

        return pages or [""]

    def _set_page_texts(self, page_texts, global_cursor_pos=None):
        self._syncing_pages = True
        self._ensure_page_count(max(1, len(page_texts)))
        self._page_texts = page_texts[:]

        for i, edit in enumerate(self.text_edits):
            text = page_texts[i] if i < len(page_texts) else ""
            with QSignalBlocker(edit):
                edit.setPlainText(text)

        if global_cursor_pos is not None and self.text_edits:
            self._restore_cursor(global_cursor_pos)

        self._syncing_pages = False

    def _restore_cursor(self, global_cursor_pos):
        global_cursor_pos = max(0, min(global_cursor_pos, len("".join(self._page_texts))))

        offset = 0
        target_index = 0
        local_pos = 0
        for i, chunk in enumerate(self._page_texts):
            next_offset = offset + len(chunk)
            if global_cursor_pos <= next_offset:
                target_index = i
                local_pos = global_cursor_pos - offset
                break
            offset = next_offset
        else:
            target_index = len(self._page_texts) - 1
            local_pos = len(self._page_texts[target_index])

        target_edit = self.text_edits[target_index]
        cursor = target_edit.textCursor()
        cursor.setPosition(local_pos)
        target_edit.setTextCursor(cursor)
        target_edit.setFocus()
        self.scroll.ensureWidgetVisible(target_edit)

    def get_full_text(self):
        return "\n".join(edit.toPlainText() for edit in self.text_edits)

    def export_to_docx(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export to Docs",
            f"{self.document.name}.docx",
            "Word Documents (*.docx)"
        )

        if not path:
            return

        if not path.lower().endswith(".docx"):
            path = f"{path}.docx"

        try:
            text = self.get_full_text()
            doc = DocxDocument()
            for paragraph in text.split("\n"):
                doc.add_paragraph(paragraph)
            doc.save(path)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Export Failed",
                f"Could not export file:\n{e}"
            )
            return

        QMessageBox.information(
            self,
            "Export Complete",
            "Word document exported successfully."
        )

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
