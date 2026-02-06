from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit
from models.table_model import TableModel
from views.table_view import TableView
from PySide6.QtCore import Signal, Qt, QItemSelectionModel
from document import Sheet
from PySide6.QtWidgets import QRadioButton


class EditorPage(QWidget):
    export_requested = Signal(object)  # document
    document_changed = Signal()
    
    def __init__(self, document):
        super().__init__()

        self.document = document
        self.sheet_buttons = []
        self.swap_mode = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        # ---------- TOOL RIBBON ----------
        tool_ribbon = QWidget()
        tool_ribbon.setObjectName("editorRibbon")
        tool_ribbon.setFixedHeight(72)

        ribbon_layout = QHBoxLayout(tool_ribbon)
        ribbon_layout.setContentsMargins(16, 0, 16, 0)

        # 🔹 Swap toggles (UI only)
        self.swap_cell_btn = QPushButton("Swap Cell")
        self.swap_row_btn = QPushButton("Swap Row")
        self.swap_col_btn = QPushButton("Swap Column")
        self.swap_rectangle_armed = False
        self.swap_cell_btn.toggled.connect(
            lambda checked: self._set_swap_mode("cell", checked)
        )

        self.swap_row_btn.toggled.connect(
            lambda checked: self._set_swap_mode("row", checked)
        )

        self.swap_col_btn.toggled.connect(
            lambda checked: self._set_swap_mode("column", checked)
        )

        for btn in (self.swap_cell_btn, self.swap_row_btn, self.swap_col_btn):
            btn.setCheckable(True)
            btn.setFixedHeight(32)

        ribbon_layout.addWidget(self.swap_cell_btn)
        ribbon_layout.addWidget(self.swap_row_btn)
        ribbon_layout.addWidget(self.swap_col_btn)

        # 🔹 THIS IS THE FIX
        ribbon_layout.addStretch()

        # 🔹 Export button (RIGHT)
        self.export_btn = QPushButton("Export to Excel")
        self.export_btn.setFixedHeight(36)

        self.export_btn.clicked.connect(
            lambda: self.export_requested.emit(self.document)
        )

        ribbon_layout.addWidget(self.export_btn)


        layout.addWidget(tool_ribbon)


        self.model = TableModel(document)
        self.view = TableView()
        self.view.get_swap_mode = lambda: self.swap_mode
        self.view.clear_swap_mode = self.clear_swap_mode


        self.view.setModel(self.model)
        self.zoom_box = QPlainTextEdit()
        self.zoom_box.setPlaceholderText("Zoom box")
        self.zoom_box.setFixedHeight(90)
        self._syncing_zoom_box = False

        self.view.set_zoom_box(self.zoom_box)
        self.view.selectionModel().currentChanged.connect(
            self._sync_zoom_box_from_selection
        )
        self.model.dataChanged.connect(self._sync_zoom_box_from_model)
        self.zoom_box.textChanged.connect(self._apply_zoom_box_edit)
        self.view.block_swap_requested.connect(self.handle_block_swap)

        self.view.drag_swap_requested.connect(self.handle_drag_swap)

        layout.addWidget(self.view)
        layout.addWidget(self.zoom_box)

# ---------- SHEET BAR (UI ONLY) ----------

        sheet_bar = QWidget()
        sheet_bar.setObjectName("sheetBar")
        sheet_bar.setFixedHeight(36)
        sheet_layout = QHBoxLayout(sheet_bar)
        sheet_layout.setContentsMargins(8, 0, 8, 0)
        sheet_layout.setSpacing(6)

        self.sheet_layout = sheet_layout
        self.refresh_sheet_buttons()
        add_sheet_btn = QPushButton("+")
        add_sheet_btn.setFixedSize(26, 26)

        add_sheet_btn.clicked.connect(self.add_sheet)

        sheet_layout.addWidget(add_sheet_btn)
        sheet_layout.addStretch()

        layout.addWidget(sheet_bar)
    def refresh_sheet_buttons(self):
        # remove old buttons
        for btn in self.sheet_buttons:
            self.sheet_layout.removeWidget(btn)
            btn.deleteLater()

        self.sheet_buttons.clear()

        # recreate buttons from document sheets
        for i, sheet in enumerate(self.document.sheets):
            btn = QPushButton(sheet.name)
            btn.setProperty("sheetButton", True)
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(
                lambda pos, idx=i: self.show_sheet_context_menu(idx, btn)
            )

            btn.setFixedHeight(26)
            btn.setCheckable(True)
            btn.setChecked(i == self.document.active_sheet_index)

            btn.clicked.connect(lambda _, idx=i: self.switch_sheet(idx))

            self.sheet_layout.insertWidget(i, btn)
            self.sheet_buttons.append(btn)


    def add_sheet(self):
        count = len(self.document.sheets) + 1
        self.document.sheets.append(Sheet(f"Sheet{count}"))
        self.document.active_sheet_index = len(self.document.sheets) - 1
        self.model.layoutChanged.emit()
        self.refresh_sheet_buttons()
        self.document_changed.emit()
    def switch_sheet(self, index):
        self.document.active_sheet_index = index
        self.model.layoutChanged.emit()
        self.refresh_sheet_buttons()
    def show_sheet_context_menu(self, index, button):
        from PySide6.QtWidgets import QMenu, QInputDialog, QMessageBox

        menu = QMenu(self)
        rename_action = menu.addAction("Rename")
        delete_action = menu.addAction("Delete")

        action = menu.exec(button.mapToGlobal(button.rect().bottomLeft()))

        if action == rename_action:
            self.rename_sheet(index)

        elif action == delete_action:
            self.delete_sheet(index)
    def rename_sheet(self, index):
        from PySide6.QtWidgets import QInputDialog

        sheet = self.document.sheets[index]

        new_name, ok = QInputDialog.getText(
            self,
            "Rename Sheet",
            "New sheet name:",
            text=sheet.name
        )

        if not ok or not new_name.strip():
            return

        sheet.name = new_name.strip()
        self.refresh_sheet_buttons()
        self.document_changed.emit() 
    def delete_sheet(self, index):
        from PySide6.QtWidgets import QMessageBox

        if len(self.document.sheets) == 1:
            QMessageBox.warning(
                self,
                "Cannot Delete Sheet",
                "A document must have at least one sheet."
            )
            return

        reply = QMessageBox.question(
            self,
            "Delete Sheet",
            "Are you sure you want to delete this sheet?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        del self.document.sheets[index]

        # adjust active sheet index
        if self.document.active_sheet_index >= len(self.document.sheets):
            self.document.active_sheet_index = len(self.document.sheets) - 1

        self.model.layoutChanged.emit()
        self.refresh_sheet_buttons()
        self.document_changed.emit() 
    def handle_drag_swap(self, start_index, end_index):
        r1, c1 = start_index.row(), start_index.column()
        r2, c2 = end_index.row(), end_index.column()

        # CELL SWAP
        if self.swap_cell_btn.isChecked():
            self.model.swap_cells(r1, c1, r2, c2)

        # ROW SWAP
        if self.swap_row_btn.isChecked():
            self.model.swap_rows(r1, r2)

        # COLUMN SWAP
        if self.swap_col_btn.isChecked():
            self.model.swap_columns(c1, c2)

        self.document_changed.emit()
    def handle_block_swap(self, src_rect, dest_top_left):
        r1, c1, r2, c2 = src_rect
        dr, dc = dest_top_left

        height = r2 - r1 + 1
        width = c2 - c1 + 1

        dest_r2 = dr + height - 1
        dest_c2 = dc + width - 1

        # prevent self-swap
        if (r1, c1) == (dr, dc):
            return

        self.model.swap_block(
            r1, c1, r2, c2,
            dr, dc, dest_r2, dest_c2
        )

        # 🔑 RESET SELECTION SAFELY (AFTER layoutChanged)
        view = self.view
        model = self.model

        new_index = model.index(dr, dc)

        sel = view.selectionModel()
        sel.clearSelection()
        sel.setCurrentIndex(
            new_index,
            QItemSelectionModel.ClearAndSelect
        )

        self.document_changed.emit()

    def _set_zoom_box_text(self, text):
        self._syncing_zoom_box = True
        self.zoom_box.setPlainText(text)
        self.zoom_box.selectAll()
        self._syncing_zoom_box = False

    def _sync_zoom_box_from_selection(self, current, previous):
        if not current.isValid():
            self._set_zoom_box_text("")
            return

        value = self.model.data(current, Qt.EditRole)
        self._set_zoom_box_text(value or "")

    def _sync_zoom_box_from_model(self, top_left, bottom_right, roles=None):
        current = self.view.currentIndex()
        if not current.isValid():
            return

        if (
            top_left.row() <= current.row() <= bottom_right.row()
            and top_left.column() <= current.column() <= bottom_right.column()
        ):
            value = self.model.data(current, Qt.EditRole)
            self._set_zoom_box_text(value or "")

    def _apply_zoom_box_edit(self):
        if self._syncing_zoom_box:
            return

        index = self.view.currentIndex()
        if not index.isValid():
            return

        text = self.zoom_box.toPlainText()
        self.model.setData(index, text, Qt.EditRole)

    def disarm_swap_rectangle(self):
        self.swap_rectangle_armed = False
        self.view.clearSelection()
    def arm_swap_rectangle(self):
        self.swap_rectangle_armed = True

    def _set_swap_mode(self, mode, checked):
        if checked:
            self.swap_mode = mode
        else:
            if self.swap_mode == mode:
                self.swap_mode = None
    def arm_rectangle_swap(self):
        self.swap_mode = "rectangle"
    def clear_swap_mode(self):
        # clear logical mode
        self.swap_mode = None

        # clear UI toggles safely
        for btn in (
            getattr(self, "swap_cell_btn", None),
            getattr(self, "swap_row_btn", None),
            getattr(self, "swap_col_btn", None),
        ):
            if btn is not None:
                btn.blockSignals(True)
                btn.setChecked(False)
                btn.blockSignals(False)
    def apply_grid_dark_mode(self, enabled: bool):
        if not enabled:
            self.setStyleSheet("")
            return

        self.setStyleSheet("""
        /* ===============================
        GLOBAL EDITOR
        =============================== */
        QWidget {
            background-color: #1e1e1e;
            color: #e6e6e6;
        }

        /* ===============================
        TOOL RIBBON
        =============================== */
        QWidget#editorRibbon {
            background-color: #252526;
        }

        QWidget#editorRibbon QPushButton {
            background-color: #2d2d30;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 6px;
            padding: 4px 12px;
        }

        QWidget#editorRibbon QPushButton:hover {
            background-color: #3a3a3a;
        }

        QWidget#editorRibbon QPushButton:checked {
            background-color: #094771;
            color: #ffffff;
            border: 1px solid #1a6fb3;
        }

        QWidget#editorRibbon QPushButton:disabled {
            color: #9e9e9e;
            background-color: #2a2a2a;
        }

        /* ===============================
        SHEET BUTTONS (NOT QTabBar!)
        =============================== */
        QPushButton[sheetButton="true"] {
            background-color: #2b2b2b;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 4px;
            padding: 0 12px;
        }

        QPushButton[sheetButton="true"]:checked {
            background-color: #094771;
            color: #ffffff;
            border: 1px solid #1a6fb3;
        }

        QPushButton[sheetButton="true"]:hover {
            background-color: #333333;
        }

        QWidget#sheetBar QPushButton {
            background-color: #2b2b2b;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 4px;
        }

        /* ===============================
        TABLE GRID
        =============================== */
        QTableView {
            background-color: #1e1e1e;
            gridline-color: #3a3a3a;
            color: #e6e6e6;
            selection-background-color: #094771;
            selection-color: #ffffff;
        }

        QTableView::item:selected {
            background-color: #094771;
            color: #ffffff;
        }

        /* ===============================
        HEADERS
        =============================== */
        QHeaderView::section {
            background-color: #252526;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            padding: 4px;
        }

        QTableCornerButton::section {
            background-color: #252526;
            border: 1px solid #3a3a3a;
        }

        /* ===============================
        SHEET BAR
        =============================== */
        QWidget#sheetBar {
            background-color: #1e1e1e;
            border-top: 1px solid #2b2b2b;
        }

        /* ===============================
        SCROLLBARS
        =============================== */
        QScrollBar:vertical, QScrollBar:horizontal {
            background: #1e1e1e;
            height: 10px;
            width: 10px;
            margin: 0px;
        }
        QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
            background: #3a3a3a;
            min-height: 24px;
            min-width: 24px;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
            background: #4a4a4a;
        }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical,
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {
            height: 0px;
            width: 0px;
        }
        """)
