from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, Signal

class TableModel(QAbstractTableModel):
    save_requested = Signal()
    history_changed = Signal(bool, bool)
    undo_state_changed = Signal(bool, bool)
    MAX_ROWS = 2000
    MAX_COLUMNS = 200

    def __init__(self, document):
        super().__init__()
        self.document = document

        self.rows = self.MAX_ROWS
        self.columns = self.MAX_COLUMNS
        self._undo_stack = []
        self._redo_stack = []
        self._macro_depth = 0
        self._macro_before = {}
        self._macro_after = {}
        self._suspend_history = False




    # ---------- REQUIRED OVERRIDES ----------

    def rowCount(self, parent=QModelIndex()):
        return self.rows

    def columnCount(self, parent=QModelIndex()):
        return self.columns


    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled

    # ---------- HEADERS ----------

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None

        if orientation == Qt.Horizontal:
            return self._column_name(section)

        return str(section + 1)

    def _column_name(self, index):
        """Convert 0 → A, 1 → B, ... 25 → Z, 26 → AA"""
        name = ""
        while index >= 0:
            name = chr(index % 26 + 65) + name
            index = index // 26 - 1
        return name
    
    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            row, col = index.row(), index.column()

            cells = self.document.active_sheet.cells
            before = cells.get((row, col), "")
            after = "" if value is None else str(value)

            if after == before:
                return False

            if after == "":
                cells.pop((row, col), None)
            else:
                cells[(row, col)] = after

            after = cells.get((row, col), "")
            self._push_change({(row, col): before}, {(row, col): after})
            self.dataChanged.emit(index, index)
            self.save_requested.emit()
            return True

        return False

    def set_cells_batch(self, changes):
        if not changes:
            return False

        cells = self.document.active_sheet.cells
        min_row = min(r for r, _ in changes.keys())
        max_row = max(r for r, _ in changes.keys())
        min_col = min(c for _, c in changes.keys())
        max_col = max(c for _, c in changes.keys())

        for (row, col), value in changes.items():
            if value == "":
                cells.pop((row, col), None)
            else:
                cells[(row, col)] = value

        self.dataChanged.emit(
            self.index(min_row, min_col),
            self.index(max_row, max_col)
        )
        self.save_requested.emit()
        return True

    def set_cells_batch(self, changes):
        if not changes:
            return False

        cells = self.document.active_sheet.cells
        min_row = min(r for r, _ in changes.keys())
        max_row = max(r for r, _ in changes.keys())
        min_col = min(c for _, c in changes.keys())
        max_col = max(c for _, c in changes.keys())

        for (row, col), value in changes.items():
            if value == "":
                cells.pop((row, col), None)
            else:
                cells[(row, col)] = value

        self.dataChanged.emit(
            self.index(min_row, min_col),
            self.index(max_row, max_col)
        )
        self.save_requested.emit()
        return True

    def set_cells_batch(self, changes):
        if not changes:
            return False

        cells = self.document.active_sheet.cells
        min_row = min(r for r, _ in changes.keys())
        max_row = max(r for r, _ in changes.keys())
        min_col = min(c for _, c in changes.keys())
        max_col = max(c for _, c in changes.keys())

        for (row, col), value in changes.items():
            if value == "":
                cells.pop((row, col), None)
            else:
                cells[(row, col)] = value

        self.dataChanged.emit(
            self.index(min_row, min_col),
            self.index(max_row, max_col)
        )
        self.save_requested.emit()
        return True


    
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        if role in (Qt.DisplayRole, Qt.EditRole):
            return self.document.active_sheet.cells.get(
                (index.row(), index.column()),
                ""
            )
        return None


    def clear_cells(self, positions):
        unique_positions = list(dict.fromkeys(positions))
        if not unique_positions:
            return False

        cells = self.document.active_sheet.cells
        before = {}
        after = {}
        rows = []
        cols = []

        for row, col in unique_positions:
            previous = cells.get((row, col), "")
            if previous == "":
                continue

            before[(row, col)] = previous
            after[(row, col)] = ""
            cells.pop((row, col), None)
            rows.append(row)
            cols.append(col)

        if not before:
            return False

        self._push_change(before, after)
        self.dataChanged.emit(
            self.index(min(rows), min(cols)),
            self.index(max(rows), max(cols))
        )
        self.save_requested.emit()
        return True

    @property
    def cells(self):
        return self.document.active_sheet.cells
    def swap_cells(self, r1, c1, r2, c2):
        positions = [(r1, c1), (r2, c2)]
        before = self._snapshot_positions(positions)
        cells = self.cells
        v1 = cells.get((r1, c1), "")
        v2 = cells.get((r2, c2), "")

        if v1:
            cells[(r2, c2)] = v1
        else:
            cells.pop((r2, c2), None)

        if v2:
            cells[(r1, c1)] = v2
        else:
            cells.pop((r1, c1), None)

        self.dataChanged.emit(
            self.index(min(r1, r2), min(c1, c2)),
            self.index(max(r1, r2), max(c1, c2))
        )
        after = self._snapshot_positions(positions)
        self._push_change(before, after)
    def swap_rows(self, r1, r2):
        if r1 == r2:
            return

        cells = self.cells
        cols = set(c for (_, c) in cells.keys())
        positions = [(r1, c) for c in cols] + [(r2, c) for c in cols]
        before = self._snapshot_positions(positions)

        for c in cols:
            v1 = cells.get((r1, c), "")
            v2 = cells.get((r2, c), "")

            if v1:
                cells[(r2, c)] = v1
            else:
                cells.pop((r2, c), None)

            if v2:
                cells[(r1, c)] = v2
            else:
                cells.pop((r1, c), None)

        after = self._snapshot_positions(positions)
        self._push_change(before, after)
        self.layoutChanged.emit()
        self.save_requested.emit()
        self._record_action(self._diff_cells(before, after))
    def swap_columns(self, c1, c2):
        if c1 == c2:
            return

        cells = self.cells
        rows = set(r for (r, _) in cells.keys())
        positions = [(r, c1) for r in rows] + [(r, c2) for r in rows]
        before = self._snapshot_positions(positions)

        for r in rows:
            v1 = cells.get((r, c1), "")
            v2 = cells.get((r, c2), "")

            if v1:
                cells[(r, c2)] = v1
            else:
                cells.pop((r, c2), None)

            if v2:
                cells[(r, c1)] = v2
            else:
                cells.pop((r, c1), None)

        after = self._snapshot_positions(positions)
        self._push_change(before, after)
        self.layoutChanged.emit()
        self.save_requested.emit()
        self._record_action(self._diff_cells(before, after))
    def swap_block(self, r1, c1, r2, c2, dr1, dc1, dr2, dc2):
        cells = self.document.active_sheet.cells

        src_h = r2 - r1
        src_w = c2 - c1

        if (dr2 - dr1) != src_h or (dc2 - dc1) != src_w:
            return  # shape mismatch

        positions = []
        for r in range(src_h + 1):
            for c in range(src_w + 1):
                positions.append((r1 + r, c1 + c))
                positions.append((dr1 + r, dc1 + c))
        before = self._snapshot_positions(positions)

        # snapshot source
        src = {}
        dst = {}

        for r in range(src_h + 1):
            for c in range(src_w + 1):
                src[(r, c)] = cells.get((r1 + r, c1 + c), "")
                dst[(r, c)] = cells.get((dr1 + r, dc1 + c), "")

        # write swapped
        for r in range(src_h + 1):
            for c in range(src_w + 1):
                if dst[(r, c)]:
                    cells[(r1 + r, c1 + c)] = dst[(r, c)]
                else:
                    cells.pop((r1 + r, c1 + c), None)

                if src[(r, c)]:
                    cells[(dr1 + r, dc1 + c)] = src[(r, c)]
                else:
                    cells.pop((dr1 + r, dc1 + c), None)

        after = self._snapshot_positions(positions)
        self._push_change(before, after)
        self.layoutChanged.emit()

    def begin_macro(self):
        self._macro_depth += 1
        if self._macro_depth == 1:
            self._macro_before = {}
            self._macro_after = {}

    def end_macro(self):
        if self._macro_depth == 0:
            return
        self._macro_depth -= 1
        if self._macro_depth > 0:
            return
        if self._macro_before or self._macro_after:
            self._undo_stack.append((self._macro_before, self._macro_after))
            self._redo_stack.clear()
            self._emit_history_state()
        self._macro_before = {}
        self._macro_after = {}

    def undo(self):
        if not self._undo_stack:
            return
        before, after = self._undo_stack.pop()
        self._redo_stack.append((before, after))
        self._apply_change(before)
        self._emit_history_state()

    def redo(self):
        if not self._redo_stack:
            return
        before, after = self._redo_stack.pop()
        self._undo_stack.append((before, after))
        self._apply_change(after)
        self._emit_history_state()

    def can_undo(self):
        return bool(self._undo_stack)

    def can_redo(self):
        return bool(self._redo_stack)

    def _snapshot_positions(self, positions):
        cells = self.cells
        snapshot = {}
        for pos in positions:
            if pos in snapshot:
                continue
            snapshot[pos] = cells.get(pos, "")
        return snapshot

    def _push_change(self, before, after):
        if self._suspend_history:
            return
        if before == after:
            return
        if self._macro_depth > 0:
            for pos, value in before.items():
                if pos not in self._macro_before:
                    self._macro_before[pos] = value
            for pos, value in after.items():
                self._macro_after[pos] = value
            return
        self._undo_stack.append((before, after))
        self._redo_stack.clear()
        self._emit_history_state()

    def _emit_history_state(self):
        can_undo = self.can_undo()
        can_redo = self.can_redo()
        self.history_changed.emit(can_undo, can_redo)
        self.undo_state_changed.emit(can_undo, can_redo)

    def _apply_change(self, values):
        if not values:
            return
        self._suspend_history = True
        cells = self.cells
        rows = []
        cols = []
        for (row, col), value in values.items():
            if value == "":
                cells.pop((row, col), None)
            else:
                cells[(row, col)] = value
            rows.append(row)
            cols.append(col)
        self._suspend_history = False
        if rows and cols:
            self.dataChanged.emit(
                self.index(min(rows), min(cols)),
                self.index(max(rows), max(cols))
            )
            self.save_requested.emit()
