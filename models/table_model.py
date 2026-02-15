from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, Signal

class TableModel(QAbstractTableModel):
    save_requested = Signal()
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
        self._compound_changes = None
        self._suspend_undo = False

        self._history_by_sheet = {}
        self._history_sheet_key = None
        self._ensure_history_for_active_sheet()




    def _active_sheet_history_key(self):
        return id(self.document.active_sheet)

    def _save_history_state(self, key):
        self._history_by_sheet[key] = {
            "undo_stack": self._undo_stack,
            "redo_stack": self._redo_stack,
            "macro_depth": self._macro_depth,
            "macro_before": self._macro_before,
            "macro_after": self._macro_after,
            "suspend_history": self._suspend_history,
        }

    def _ensure_history_for_active_sheet(self):
        key = self._active_sheet_history_key()
        if self._history_sheet_key == key:
            return

        if self._history_sheet_key is not None:
            self._save_history_state(self._history_sheet_key)

        state = self._history_by_sheet.get(key)
        if state is None:
            state = {
                "undo_stack": [],
                "redo_stack": [],
                "macro_depth": 0,
                "macro_before": {},
                "macro_after": {},
                "suspend_history": False,
            }
            self._history_by_sheet[key] = state

        self._undo_stack = state["undo_stack"]
        self._redo_stack = state["redo_stack"]
        self._macro_depth = state["macro_depth"]
        self._macro_before = state["macro_before"]
        self._macro_after = state["macro_after"]
        self._suspend_history = state["suspend_history"]
        self._history_sheet_key = key

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
        self._ensure_history_for_active_sheet()
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
        self.save_requested.emit()
        self._record_action([
            ((r1, c1), v1, cells.get((r1, c1), "")),
            ((r2, c2), v2, cells.get((r2, c2), "")),
        ])
    def swap_rows(self, r1, r2):
        if r1 == r2:
            return

        cells = self.cells
        cols = set(c for (_, c) in cells.keys())
        before = {(r1, c): cells.get((r1, c), "") for c in cols}
        before.update({(r2, c): cells.get((r2, c), "") for c in cols})

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

        after = {(r1, c): cells.get((r1, c), "") for c in cols}
        after.update({(r2, c): cells.get((r2, c), "") for c in cols})
        self.layoutChanged.emit()
        self.save_requested.emit()
        self._record_action(self._diff_cells(before, after))
    def swap_columns(self, c1, c2):
        if c1 == c2:
            return

        cells = self.cells
        rows = set(r for (r, _) in cells.keys())
        before = {(r, c1): cells.get((r, c1), "") for r in rows}
        before.update({(r, c2): cells.get((r, c2), "") for r in rows})

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

        after = {(r, c1): cells.get((r, c1), "") for r in rows}
        after.update({(r, c2): cells.get((r, c2), "") for r in rows})
        self.layoutChanged.emit()
        self.save_requested.emit()
        self._record_action(self._diff_cells(before, after))
    def swap_block(self, r1, c1, r2, c2, dr1, dc1, dr2, dc2):
        cells = self.document.active_sheet.cells

        src_h = r2 - r1
        src_w = c2 - c1

        if (dr2 - dr1) != src_h or (dc2 - dc1) != src_w:
            return  # shape mismatch

        before = {}
        for r in range(src_h + 1):
            for c in range(src_w + 1):
                before[(r1 + r, c1 + c)] = cells.get((r1 + r, c1 + c), "")
                before[(dr1 + r, dc1 + c)] = cells.get((dr1 + r, dc1 + c), "")

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

        after = {}
        for r in range(src_h + 1):
            for c in range(src_w + 1):
                after[(r1 + r, c1 + c)] = cells.get((r1 + r, c1 + c), "")
                after[(dr1 + r, dc1 + c)] = cells.get((dr1 + r, dc1 + c), "")
        self.layoutChanged.emit()
        self.save_requested.emit()
        self._record_action(self._diff_cells(before, after))

    def begin_macro(self):
        self._ensure_history_for_active_sheet()
        self._macro_depth += 1
        if self._macro_depth == 1:
            self._macro_before = {}
            self._macro_after = {}

    def end_macro(self):
        self._ensure_history_for_active_sheet()
        if self._macro_depth == 0:
            return
        changes = self._compound_changes
        self._compound_changes = None
        if not changes:
            return
        self._push_action(changes)

    def undo(self):
        self._ensure_history_for_active_sheet()
        if not self._undo_stack:
            return
        changes = self._undo_stack.pop()
        self._apply_changes(changes, use_new=False)
        self._redo_stack.append(changes)
        self._emit_undo_state()

    def redo(self):
        self._ensure_history_for_active_sheet()
        if not self._redo_stack:
            return
        changes = self._redo_stack.pop()
        self._apply_changes(changes, use_new=True)
        self._undo_stack.append(changes)
        self._emit_undo_state()

    def can_undo(self):
        self._ensure_history_for_active_sheet()
        return bool(self._undo_stack)

    def can_redo(self):
        self._ensure_history_for_active_sheet()
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
        self._ensure_history_for_active_sheet()
        if self._suspend_history:
            return
        if self._compound_changes is not None:
            self._compound_changes.extend(changes)
            return
        self._push_action(changes)

    def _push_action(self, changes):
        self._undo_stack.append(changes)
        self._redo_stack.clear()
        self._emit_undo_state()

    def _push_change(self, before, after):
        """Backward-compatible helper used by older undo code paths.

        Accepts dicts keyed by (row, col) with old/new values and records a
        single undo action equivalent to `_record_action` payloads.
        """
        if before is None:
            before = {}
        if after is None:
            after = {}

        # If a caller already passes the normalized tuple payload,
        # accept it directly.
        if isinstance(before, list) and all(isinstance(item, tuple) and len(item) == 3 for item in before):
            self._record_action(before)
            return

        if not isinstance(before, dict):
            before = dict(before)
        if not isinstance(after, dict):
            after = dict(after)

        changes = []
        keys = set(before.keys()) | set(after.keys())
        for key in keys:
            old_value = before.get(key, "")
            new_value = after.get(key, "")
            if old_value != new_value:
                changes.append((key, old_value, new_value))
        if changes:
            self._record_action(changes)

    def _emit_undo_state(self):
        self.undo_state_changed.emit(self.can_undo(), self.can_redo())

    def _apply_changes(self, changes, use_new: bool):
        cells = self.document.active_sheet.cells
        self._suspend_undo = True
        rows = []
        cols = []
        for (row, col), old_value, new_value in changes:
            value = new_value if use_new else old_value
            if value == "":
                cells.pop((row, col), None)
            else:
                cells[(row, col)] = value
            rows.append(row)
            cols.append(col)
        self._suspend_undo = False
        if rows and cols:
            top_left = self.index(min(rows), min(cols))
            bottom_right = self.index(max(rows), max(cols))
            self.dataChanged.emit(top_left, bottom_right)
        else:
            self.layoutChanged.emit()
        self.save_requested.emit()

    def _diff_cells(self, before, after):
        changes = []
        for key, old_value in before.items():
            new_value = after.get(key, "")
            if old_value != new_value:
                changes.append((key, old_value, new_value))
        return changes
