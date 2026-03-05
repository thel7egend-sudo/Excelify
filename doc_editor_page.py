from docx import Document as DocxDocument
from docx.shared import Inches
from PySide6.QtCore import QSizeF, Qt, Signal
from PySide6.QtGui import QColor, QKeyEvent, QTextCursor, QTextDocument
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class PageTextEdit(QTextEdit):
    backspace_at_start = Signal()
    return_pressed = Signal()

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter) and event.modifiers() == Qt.NoModifier:
            self.return_pressed.emit()
            return

        if (
            event.key() == Qt.Key_Backspace
            and not self.textCursor().hasSelection()
            and self.textCursor().position() == 0
        ):
            self.backspace_at_start.emit()
            return
        super().keyPressEvent(event)

    def wheelEvent(self, event):
        # Keep each page fixed like a printed sheet: wheel scrolling should
        # scroll the outer document area, not create hidden scroll inside a page.
        event.ignore()


class PageWidget(QFrame):
    PAGE_WIDTH = 800
    PAGE_HEIGHT = 1100
    PAGE_MARGIN = 56

    def __init__(self):
        super().__init__()
        self.setObjectName("docPage")
        self.setFixedSize(self.PAGE_WIDTH, self.PAGE_HEIGHT)
        self.setFrameShape(QFrame.NoFrame)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(28)
        shadow.setXOffset(0)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 45))
        self.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(self.PAGE_MARGIN, self.PAGE_MARGIN, self.PAGE_MARGIN, self.PAGE_MARGIN)
        layout.setSpacing(0)

        self.editor = PageTextEdit()
        self.editor.setFrameShape(QFrame.NoFrame)
        self.editor.setAcceptRichText(False)
        self.editor.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.editor.setLineWrapMode(QTextEdit.WidgetWidth)
        self.editor.setFixedHeight(self.usable_page_height)
        layout.addWidget(self.editor)

    @property
    def usable_page_width(self):
        return self.PAGE_WIDTH - (self.PAGE_MARGIN * 2)

    @property
    def usable_page_height(self):
        return self.PAGE_HEIGHT - (self.PAGE_MARGIN * 2)


class WordStyleEditor(QWidget):
    textChanged = Signal()

    PAGE_GAP = 48
    LIGHT_WORKSPACE = "#dfe1e5"
    DARK_WORKSPACE = "#13161c"

    def __init__(self, initial_text=""):
        super().__init__()
        self._is_reflowing = False
        self._active_page_idx = 0
        self._workspace_color = self.LIGHT_WORKSPACE
        self._pages = []

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setFrameShape(QFrame.NoFrame)

        self.container = QWidget()
        self.pages_layout = QVBoxLayout(self.container)
        self.pages_layout.setContentsMargins(24, 24, 24, 24)
        self.pages_layout.setSpacing(self.PAGE_GAP)
        self.pages_layout.setAlignment(Qt.AlignHCenter | Qt.AlignTop)

        self.scroll_area.setWidget(self.container)
        root.addWidget(self.scroll_area)

        self._apply_theme(False)
        self.setPlainText(initial_text)

    @property
    def usable_page_width(self):
        return PageWidget.PAGE_WIDTH - (PageWidget.PAGE_MARGIN * 2)

    @property
    def usable_page_height(self):
        return PageWidget.PAGE_HEIGHT - (PageWidget.PAGE_MARGIN * 2)

    def _create_page(self, text=""):
        page = PageWidget()
        page.editor.document().setDocumentMargin(0)
        page.editor.setPlainText(text)
        page.editor.textChanged.connect(lambda p=page: self._on_page_text_changed(p))
        page.editor.backspace_at_start.connect(lambda p=page: self._merge_with_previous(p))
        page.editor.return_pressed.connect(lambda p=page: self._handle_return_pressed(p))
        page.editor.selectionChanged.connect(lambda p=page: self._track_active_page(p))
        page.editor.cursorPositionChanged.connect(lambda p=page: self._track_active_page(p))
        page.editor.verticalScrollBar().rangeChanged.connect(lambda *_args, e=page.editor: e.verticalScrollBar().setValue(0))
        return page

    def _append_page(self, text=""):
        page = self._create_page(text)
        self._pages.append(page)
        self.pages_layout.addWidget(page)
        return page

    def _insert_page_after(self, idx, text=""):
        page = self._create_page(text)
        self._pages.insert(idx + 1, page)
        self.pages_layout.insertWidget(idx + 1, page)
        return page

    def _remove_page(self, idx):
        if len(self._pages) <= 1:
            return
        page = self._pages.pop(idx)
        self.pages_layout.removeWidget(page)
        page.deleteLater()

    def _track_active_page(self, page):
        if page in self._pages:
            self._active_page_idx = self._pages.index(page)

    def _text_fits(self, text):
        probe = QTextDocument()
        probe.setDefaultFont(self._pages[0].editor.font())
        probe.setDocumentMargin(0)
        probe.setTextWidth(self.usable_page_width)
        probe.setPageSize(QSizeF(self.usable_page_width, self.usable_page_height))
        probe.setPlainText(text)
        doc_height = probe.documentLayout().documentSize().height()
        # Small epsilon avoids allowing partially clipped final lines.
        return doc_height <= (self.usable_page_height - 0.5)

    def _handle_return_pressed(self, page):
        if self._is_reflowing or page not in self._pages:
            return

        editor = page.editor
        cursor = editor.textCursor()
        if cursor.hasSelection():
            cursor.removeSelectedText()

        idx = self._pages.index(page)
        text = editor.toPlainText()
        pos = cursor.position()
        candidate = f"{text[:pos]}\n{text[pos:]}"

        self._is_reflowing = True
        if self._text_fits(candidate):
            cursor.insertBlock()
            editor.setTextCursor(cursor)
            self._is_reflowing = False
            self.textChanged.emit()
            return

        before = text[:pos]
        after = text[pos:]

        editor.blockSignals(True)
        editor.setPlainText(before)
        editor.blockSignals(False)

        if idx + 1 >= len(self._pages):
            self._append_page("")

        next_editor = self._pages[idx + 1].editor
        next_text = next_editor.toPlainText()
        next_editor.blockSignals(True)
        next_editor.setPlainText(f"\n{after}{next_text}")
        next_editor.blockSignals(False)

        self._rebalance_from(idx)
        self._is_reflowing = False

        next_cursor = next_editor.textCursor()
        next_cursor.setPosition(1)
        next_editor.setTextCursor(next_cursor)
        next_editor.setFocus()
        self.textChanged.emit()

    def _fitting_index(self, text):
        if not text:
            return 0
        low, high = 0, len(text)
        fit = 0
        while low <= high:
            mid = (low + high) // 2
            chunk = text[:mid]
            if self._text_fits(chunk):
                fit = mid
                low = mid + 1
            else:
                high = mid - 1

        split_at = fit
        while 1 < split_at < len(text) and not text[split_at - 1].isspace() and text[split_at].isalnum():
            split_at -= 1

        return max(1, split_at)

    def _rebalance_from(self, start_idx):
        idx = max(0, start_idx)
        while idx < len(self._pages):
            current = self._pages[idx].editor
            text = current.toPlainText()
            if self._text_fits(text):
                idx += 1
                continue

            split_at = self._fitting_index(text)
            leading = text[:split_at]
            trailing = text[split_at:]
            old_pos = current.textCursor().position()

            current.blockSignals(True)
            current.setPlainText(leading)
            current.blockSignals(False)

            if idx + 1 >= len(self._pages):
                self._append_page("")

            next_editor = self._pages[idx + 1].editor
            next_text = next_editor.toPlainText()
            next_editor.blockSignals(True)
            next_editor.setPlainText(trailing + next_text)
            next_editor.blockSignals(False)

            if old_pos > len(leading):
                cursor = next_editor.textCursor()
                cursor.setPosition(old_pos - len(leading))
                next_editor.setTextCursor(cursor)
                next_editor.setFocus()
                self._active_page_idx = idx + 1
            else:
                cursor = current.textCursor()
                cursor.setPosition(old_pos)
                current.setTextCursor(cursor)

        self._pull_text_up(max(0, start_idx - 1))

    def _pull_text_up(self, start_idx):
        idx = max(0, start_idx)
        while idx < len(self._pages) - 1:
            current = self._pages[idx].editor
            nxt = self._pages[idx + 1].editor
            current_text = current.toPlainText()
            next_text = nxt.toPlainText()

            if not next_text:
                self._remove_page(idx + 1)
                continue

            if not self._text_fits(current_text + next_text[:1]):
                idx += 1
                continue

            low, high, fit = 1, len(next_text), 0
            while low <= high:
                mid = (low + high) // 2
                candidate = current_text + next_text[:mid]
                if self._text_fits(candidate):
                    fit = mid
                    low = mid + 1
                else:
                    high = mid - 1

            moved = next_text[:fit]
            remainder = next_text[fit:]

            current.blockSignals(True)
            current.setPlainText(current_text + moved)
            current.blockSignals(False)

            nxt.blockSignals(True)
            nxt.setPlainText(remainder)
            nxt.blockSignals(False)

            if not remainder:
                self._remove_page(idx + 1)
            else:
                idx += 1

    def _merge_with_previous(self, page):
        if page not in self._pages:
            return
        idx = self._pages.index(page)
        if idx == 0:
            return

        prev_editor = self._pages[idx - 1].editor
        this_editor = self._pages[idx].editor

        prev_text = prev_editor.toPlainText()
        this_text = this_editor.toPlainText()

        self._is_reflowing = True
        prev_editor.blockSignals(True)
        prev_editor.setPlainText(prev_text + this_text)
        prev_editor.blockSignals(False)
        self._remove_page(idx)
        self._rebalance_from(idx - 1)
        self._is_reflowing = False

        cursor = prev_editor.textCursor()
        cursor.setPosition(len(prev_text))
        prev_editor.setTextCursor(cursor)
        prev_editor.setFocus()

        self.textChanged.emit()

    def _on_page_text_changed(self, page):
        if self._is_reflowing:
            return
        if page not in self._pages:
            return

        self._is_reflowing = True
        idx = self._pages.index(page)
        self._rebalance_from(idx)
        self._is_reflowing = False
        for pg in self._pages:
            pg.editor.verticalScrollBar().setValue(0)
        self.textChanged.emit()

    def toPlainText(self):
        return "".join(page.editor.toPlainText() for page in self._pages)

    def setPlainText(self, text):
        self._is_reflowing = True
        for page in self._pages:
            self.pages_layout.removeWidget(page)
            page.deleteLater()
        self._pages = []

        self._append_page(text or "")
        self._rebalance_from(0)
        self._is_reflowing = False

        if self._pages:
            self._pages[0].editor.moveCursor(QTextCursor.Start)
            self._pages[0].editor.setFocus()

    def _apply_theme(self, enabled: bool):
        workspace = self.DARK_WORKSPACE if enabled else self.LIGHT_WORKSPACE
        page_text_color = "#eaeaea" if enabled else "#111827"
        page_bg = "#1d1f24" if enabled else "#ffffff"
        page_border = "rgba(255,255,255,0.05)" if enabled else "#d8dde6"

        self.container.setStyleSheet(f"background: {workspace};")
        self.setStyleSheet(
            f"""
            QScrollArea {{
                background: {workspace};
                border: none;
            }}
            QFrame#docPage {{
                background: {page_bg};
                border: 1px solid {page_border};
            }}
            QTextEdit {{
                background: transparent;
                color: {page_text_color};
                border: none;
                font-size: 14px;
            }}
            """
        )

    def set_dark_mode(self, enabled: bool):
        self._apply_theme(enabled)


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
            """
            )
            self.editor.set_dark_mode(False)
            return

        self.editor.set_dark_mode(True)
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
        """
        )
