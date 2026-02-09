from PySide6.QtWidgets import QTableView, QApplication
from PySide6.QtCore import Qt, Signal
from PySide6.QtCore import QItemSelectionModel
from PySide6.QtGui import QPainter, QColor, QKeySequence

class TableView(QTableView):
    drag_swap_requested = Signal(object, object)
    block_swap_requested = Signal(tuple, tuple)
    selection_finalized = Signal()
    # ((r1,c1,r2,c2), (dest_r1,dest_c1))
    def __init__(self):
        super().__init__()

        self.setAlternatingRowColors(False)
        self.setShowGrid(True)
        self.swap_mode = None

        # Excel-like behavior
        self.horizontalHeader().setStretchLastSection(False)
        self.horizontalHeader().setSectionsMovable(False)
        self.horizontalHeader().setDefaultSectionSize(100)

        self.verticalHeader().setDefaultSectionSize(24)

        self.setSelectionMode(QTableView.SingleSelection)
        self.setSelectionBehavior(QTableView.SelectItems)
        self.setSelectionMode(QTableView.ExtendedSelection)
        self.setSelectionBehavior(QTableView.SelectItems)
        self._saved_selection_mode = None
        self._ghost_active = False
        self._ghost_rect = None
        self._drag_start_pos = None
        
        self._drag_start_index = None
        self.zoom_box = None
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)

    def mousePressEvent(self, event):
        index = self.indexAt(event.pos())

        mode = self.get_swap_mode()

        if (
            event.button() == Qt.LeftButton
            and mode is not None
            and index.isValid()
            and (
                mode != "rectangle"                    or self.selectionModel().isSelected(index)
            )
        ):
            selection = self.selectionModel()
            if selection is not None and not selection.isSelected(index):
                selection.setCurrentIndex(
                    index,
                    QItemSelectionModel.ClearAndSelect
                )

            # 🔒 HARD SWITCH: disable Qt selection
            self._saved_selection_mode = self.selectionMode()
            self.setSelectionMode(QTableView.NoSelection)

            self._ghost_active = True
            self._drag_start_pos = event.pos()
            self._ghost_rect = None

            event.accept()
            return  # ⛔ STOP Qt completely

        if event.button() == Qt.LeftButton:
            self.setFocus(Qt.MouseFocusReason)

        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        super().mouseDoubleClickEvent(event)
        if event.button() == Qt.LeftButton and self.zoom_box and self.zoom_box.isVisible():
            self.zoom_box.setFocus(Qt.MouseFocusReason)
            self.zoom_box.selectAll()

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            self._copy_selection_to_clipboard()
            return

        if event.matches(QKeySequence.Paste):
            self._paste_clipboard_to_selection()
            return

        if event.matches(QKeySequence.Cut):
            self._cut_selection_to_clipboard()
            return

        super().keyPressEvent(event)

    def mouseMoveEvent(self, event):
        if self._ghost_active:
            self._update_ghost(event.pos())
            self.viewport().update()
            event.accept()
            return

        super().mouseMoveEvent(event)





    def paintEvent(self, event):
        super().paintEvent(event)

        if self._ghost_active and self._ghost_rect:
            painter = QPainter(self.viewport())
            painter.setPen(QColor(0, 120, 215))
            painter.setBrush(QColor(0, 120, 215, 60))
            painter.drawRect(self._ghost_rect)



    def mouseReleaseEvent(self, event):
        if self._ghost_active:
            self._ghost_active = False

            # restore selection mode
            self.setSelectionMode(self._saved_selection_mode)

            end_index = self.indexAt(event.pos())
            selected = self.selectionModel().selectedIndexes()
            mode = self.get_swap_mode()

            if end_index.isValid() and selected:
                rows = [i.row() for i in selected]
                cols = [i.column() for i in selected]

                src_rect = (
                    min(rows), min(cols),
                    max(rows), max(cols)
                )

                dest_top_left = (
                    end_index.row(),
                    end_index.column()
                )

                if mode == "rectangle":
                    self.block_swap_requested.emit(src_rect, dest_top_left)

                elif mode == "row":
                    r = src_rect[0]
                    self.block_swap_requested.emit(
                        (r, 0, r, self.model().columnCount() - 1),
                        (end_index.row(), 0)
                    )

                elif mode == "column":
                    c = src_rect[1]
                    self.block_swap_requested.emit(
                        (0, c, self.model().rowCount() - 1, c),
                        (0, end_index.column())
                    )

                elif mode == "cell":
                    r, c = src_rect[0], src_rect[1]
                    self.block_swap_requested.emit(
                        (r, c, r, c),
                        (end_index.row(), end_index.column())
                    )


            # full cleanup
            self._ghost_rect = None
            self._drag_start_pos = None
            self.clearSelection()
            if mode == "rectangle":
                self.clear_swap_mode()
            self.viewport().update()
            event.accept()
            return

        super().mouseReleaseEvent(event)
        if event.button() == Qt.LeftButton:
            self.selection_finalized.emit()
    

    def _update_ghost(self, pos):
        index = self.indexAt(pos)
        if not index.isValid():
            self._ghost_rect = None
            return

        mode = self.get_swap_mode()

        if mode == "rectangle":
            rows = [i.row() for i in self.selectionModel().selectedIndexes()]
            cols = [i.column() for i in self.selectionModel().selectedIndexes()]

            if not rows or not cols:
                self._ghost_rect = None
                return

            h = max(rows) - min(rows) + 1
            w = max(cols) - min(cols) + 1

            top_left = index
            bottom_right = self.model().index(
                index.row() + h - 1,
                index.column() + w - 1
            )

        elif mode == "row":
            r = index.row()
            top_left = self.model().index(r, 0)
            bottom_right = self.model().index(
                r, self.model().columnCount() - 1
            )

        elif mode == "column":
            c = index.column()
            top_left = self.model().index(0, c)
            bottom_right = self.model().index(
                self.model().rowCount() - 1, c
            )

        elif mode == "cell":
            top_left = index
            bottom_right = index

        else:
            self._ghost_rect = None
            return

        self._ghost_rect = self.visualRect(top_left).united(
            self.visualRect(bottom_right)
        )

    def _show_context_menu(self, pos):
        index = self.indexAt(pos)
        if not index.isValid():
            return

        selection = self.selectionModel()
        if selection is None or not selection.isSelected(index):
            return

        from PySide6.QtWidgets import QMenu
        menu = QMenu(self)

        swap_action = menu.addAction("Swap Rectangle")
        copy_action = menu.addAction("Copy")
        paste_action = menu.addAction("Paste")
        cut_action = menu.addAction("Cut")
        action = menu.exec(self.viewport().mapToGlobal(pos))

        if action == swap_action:
            parent = self.parent()
            if hasattr(parent, "arm_rectangle_swap"):
                parent.arm_rectangle_swap()
        elif action == copy_action:
            self._copy_selection_to_clipboard()
        elif action == paste_action:
            self._paste_clipboard_to_selection()
        elif action == cut_action:
            self._cut_selection_to_clipboard()
    
    def clear_swap_mode(self):
        self.swap_mode = None

    def set_zoom_box(self, zoom_box):
        self.zoom_box = zoom_box

    def _selected_rect(self):
        selection = self.selectionModel()
        if selection is None:
            return None

        selected = selection.selectedIndexes()
        if not selected:
            index = self.currentIndex()
            if not index.isValid():
                return None
            return index.row(), index.column(), index.row(), index.column()

        rows = [i.row() for i in selected]
        cols = [i.column() for i in selected]
        return min(rows), min(cols), max(rows), max(cols)

    def _copy_selection_to_clipboard(self):
        rect = self._selected_rect()
        if rect is None:
            return

        r1, c1, r2, c2 = rect
        lines = []
        for r in range(r1, r2 + 1):
            row_values = []
            for c in range(c1, c2 + 1):
                value = self.model().data(self.model().index(r, c), Qt.DisplayRole)
                row_values.append("" if value is None else str(value))
            lines.append("\t".join(row_values))

        QApplication.clipboard().setText("\n".join(lines))

    def _paste_clipboard_to_selection(self):
        text = QApplication.clipboard().text()
        if text == "":
            return

        rect = self._selected_rect()
        if rect is None:
            return

        start_row, start_col, _, _ = rect
        rows = text.splitlines() or [""]
        for r_offset, row_text in enumerate(rows):
            cols = row_text.split("\t")
            for c_offset, value in enumerate(cols):
                index = self.model().index(start_row + r_offset, start_col + c_offset)
                if index.isValid():
                    self.model().setData(index, value, Qt.EditRole)

    def _cut_selection_to_clipboard(self):
        rect = self._selected_rect()
        if rect is None:
            return

        self._copy_selection_to_clipboard()
        r1, c1, r2, c2 = rect
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                index = self.model().index(r, c)
                if index.isValid():
                    self.model().setData(index, "", Qt.EditRole)

        start_row, start_col, _, _ = rect
        rows = text.splitlines() or [""]
        for r_offset, row_text in enumerate(rows):
            cols = row_text.split("\t")
            for c_offset, value in enumerate(cols):
                index = self.model().index(start_row + r_offset, start_col + c_offset)
                if index.isValid():
                    self.model().setData(index, value, Qt.EditRole)
