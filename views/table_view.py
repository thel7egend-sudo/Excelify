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
        if event.button() == Qt.LeftButton and self.zoom_box:
            self.zoom_box.setFocus(Qt.MouseFocusReason)
            self.zoom_box.selectAll()

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            index = self.currentIndex()
            if index.isValid():
                text = self.model().data(index, Qt.DisplayRole)
                QApplication.clipboard().setText("" if text is None else str(text))
                return

        if event.matches(QKeySequence.Paste):
            index = self.currentIndex()
            if index.isValid():
                text = QApplication.clipboard().text()
                self.model().setData(index, text, Qt.EditRole)
                return

        if self.zoom_box:
            if event.matches(QKeySequence.Paste):
                self.zoom_box.setFocus(Qt.OtherFocusReason)
                self.zoom_box.selectAll()
                self.zoom_box.paste()
                return

            text = event.text()
            if text and not (event.modifiers() & (Qt.ControlModifier | Qt.AltModifier | Qt.MetaModifier)):
                self.zoom_box.setFocus(Qt.OtherFocusReason)
                self.zoom_box.selectAll()
                self.zoom_box.insert(text)
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

                mode = self.get_swap_mode()

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

        if not self.selectionModel().isSelected(index):
            return

        from PySide6.QtWidgets import QMenu
        menu = QMenu(self)

        swap_action = menu.addAction("Swap Rectangle")
        copy_action = menu.addAction("Copy")
        paste_action = menu.addAction("Paste")
        action = menu.exec(self.viewport().mapToGlobal(pos))

        if action == swap_action:
            parent = self.parent()
            if hasattr(parent, "arm_rectangle_swap"):
                parent.arm_rectangle_swap()
        elif action == copy_action:
            self._copy_selection_to_clipboard()
        elif action == paste_action:
            self._paste_clipboard_to_selection()
    
    def clear_swap_mode(self):
        self.swap_mode = None

    def set_zoom_box(self, zoom_box):
        self.zoom_box = zoom_box

