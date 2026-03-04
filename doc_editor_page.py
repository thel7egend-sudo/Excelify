from docx import Document as DocxDocument
from docx.shared import Inches
from PySide6.QtCore import QTimer, Qt, QRectF, QSizeF, Signal
from PySide6.QtGui import QColor, QPainter, QTextDocument
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class WordStyleEditor(QTextEdit):
    PAGE_WIDTH = 800
    PAGE_HEIGHT = 1100
    PAGE_MARGIN = 56

    BACKGROUND = QColor("#dfe1e5")
    PAGE_COLOR = QColor("#ffffff")
    PAGE_BORDER = QColor("#d8dde6")
    SHADOW = QColor(0, 0, 0, 18)

    def __init__(self, initial_text=""):
        super().__init__()
        self._page_count = 1
        self._metrics_guard = False
        self._pending_sync = False
        self._last_viewport_margins = None

        self.setFrameShape(QFrame.NoFrame)
        self.setAcceptRichText(False)
        self.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.document().setDocumentMargin(self.PAGE_MARGIN)

        self.setPlainText(initial_text)

        self.document().documentLayout().documentSizeChanged.connect(self._schedule_metrics_sync)
        self.cursorPositionChanged.connect(self._keep_cursor_visible)
        self._schedule_metrics_sync()

    @property
    def usable_page_width(self):
        return self.PAGE_WIDTH - (self.PAGE_MARGIN * 2)

    @property
    def usable_page_height(self):
        return self.PAGE_HEIGHT - (self.PAGE_MARGIN * 2)

    @property
    def _page_size(self):
        return QSizeF(self.PAGE_WIDTH, self.PAGE_HEIGHT)

    def _schedule_metrics_sync(self):
        if self._pending_sync:
            return
        self._pending_sync = True
        QTimer.singleShot(0, self._sync_metrics)

    def _sync_metrics(self):
        if self._metrics_guard:
            return

        self._pending_sync = False
        self._metrics_guard = True
        try:
            if self.document().pageSize() != self._page_size:
                self.document().setPageSize(self._page_size)

            if self.lineWrapColumnOrWidth() != self.usable_page_width:
                self.setLineWrapColumnOrWidth(self.usable_page_width)

            self._page_count = max(1, self.document().pageCount())
            self._update_viewport_margins()
            self.viewport().update()
        finally:
            self._metrics_guard = False

    def _update_viewport_margins(self):
        content_width = self.viewport().width()
        page_x = max(18, (content_width - self.PAGE_WIDTH) / 2)
        side_pad = max(0, int(page_x))
        top_pad = 24
        bottom_pad = 24
        margins = (side_pad, top_pad, side_pad, bottom_pad)
        if margins != self._last_viewport_margins:
            self._last_viewport_margins = margins
            self.setViewportMargins(*margins)

    def _keep_cursor_visible(self):
        cursor_rect = self.cursorRect().adjusted(0, -28, 0, 28)
        viewport_rect = self.viewport().rect()
        bar = self.verticalScrollBar()

        if cursor_rect.bottom() > viewport_rect.bottom():
            bar.setValue(bar.value() + (cursor_rect.bottom() - viewport_rect.bottom()))
        elif cursor_rect.top() < viewport_rect.top():
            bar.setValue(bar.value() - (viewport_rect.top() - cursor_rect.top()))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._schedule_metrics_sync()

    def paintEvent(self, event):
        painter = QPainter(self.viewport())
        painter.fillRect(self.viewport().rect(), self.BACKGROUND)

        content_width = self.viewport().width()
        page_x = max(18, (content_width - self.PAGE_WIDTH) / 2)

        for idx in range(self._page_count):
            y = 24 + (idx * self.PAGE_HEIGHT)
            shadow_rect = QRectF(page_x + 3, y + 4, self.PAGE_WIDTH, self.PAGE_HEIGHT)
            page_rect = QRectF(page_x, y, self.PAGE_WIDTH, self.PAGE_HEIGHT)

            painter.fillRect(shadow_rect, self.SHADOW)
            painter.fillRect(page_rect, self.PAGE_COLOR)
            painter.setPen(self.PAGE_BORDER)
            painter.drawRect(page_rect)

        super().paintEvent(event)


class DocEditorPage(QWidget):
    document_changed = Signal()
    export_requested = Signal(object)

    def __init__(self, document):
        super().__init__()
        self.document = document

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

        self.editor = WordStyleEditor(self.document.content or "")
        self.editor.textChanged.connect(self._on_text_changed)

        layout.addWidget(self.ribbon)
        layout.addWidget(self.editor)

        self.apply_grid_dark_mode(False)

    @property
    def _usable_page_width(self):
        return self.editor.usable_page_width

    @property
    def _usable_page_height(self):
        return self.editor.usable_page_height

    def _on_text_changed(self):
        self.document.content = self.editor.toPlainText()
        self.document_changed.emit()

    def _paginate_text(self, text):
        if not text:
            return [""]

        probe = QTextDocument()
        probe.setDocumentMargin(0)
        probe.setDefaultFont(self.editor.font())
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

    def export_to_docx(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export to Docs",
            f"{self.document.name}.docx",
            "Word Documents (*.docx)",
        )

        if not path:
            return

        if not path.lower().endswith(".docx"):
            path = f"{path}.docx"

        try:
            text = self.editor.toPlainText()
            page_chunks = self._paginate_text(text)

            doc = DocxDocument()
            section = doc.sections[0]
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.left_margin = Inches(0.7)
            section.right_margin = Inches(0.7)
            section.top_margin = Inches(0.7)
            section.bottom_margin = Inches(0.7)

            for page_index, chunk in enumerate(page_chunks):
                paragraphs = chunk.split("\n")
                for paragraph in paragraphs:
                    doc.add_paragraph(paragraph)

                if page_index < len(page_chunks) - 1:
                    doc.add_page_break()

            doc.save(path)
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", f"Could not export file:\n{e}")
            return

        QMessageBox.information(self, "Export Complete", "Word document exported successfully.")

    def apply_grid_dark_mode(self, enabled: bool):
        if not enabled:
            self.setStyleSheet(
                """
                QWidget {
                    background: #f7f8fa;
                    color: #111827;
                }
                QWidget#docRibbon {
                    background-color: #f9fafb;
                    border-bottom: 1px solid #e5e7eb;
                }
                QTextEdit {
                    background: #dfe1e5;
                    color: #111827;
                    border: none;
                    font-size: 14px;
                }
            """
            )
            return

        self.setStyleSheet(
            """
            QWidget {
                background: #202124;
                color: #eaeaea;
            }
            QWidget#docRibbon {
                background-color: #252525;
                border-bottom: 1px solid rgba(255, 255, 255, 0.06);
            }
            QTextEdit {
                background: #1b1c1e;
                color: #f3f4f6;
                border: none;
                font-size: 14px;
            }
        """
        )
