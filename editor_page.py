from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QToolButton,
    QToolTip,
    QVBoxLayout,
    QWidget,
    QTableView,
)
from PySide6.QtCore import (
    Property,
    QItemSelectionModel,
    QPropertyAnimation,
    QThread,
    QEasingCurve,
    Signal,
    Qt,
    QPoint,
    QTimer,
)
from PySide6.QtGui import QColor, QFont, QPainter, QTextCursor, QPen
from PySide6.QtWidgets import QStyle, QStyleOptionToolButton

from models.table_model import TableModel
from views.table_view import TableView
from voice.voice_controller import VoiceController, TranscriptionTarget
from document import Sheet


class DictateToolButton(QToolButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._pulse_scale = 1.0

    def _get_pulse_scale(self):
        return self._pulse_scale

    def _set_pulse_scale(self, value):
        self._pulse_scale = value
        self.update()

    pulse_scale = Property(float, _get_pulse_scale, _set_pulse_scale)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        center = self.rect().center()
        painter.translate(center)
        painter.scale(self._pulse_scale, self._pulse_scale)
        painter.translate(-center)

        option = QStyleOptionToolButton()
        self.initStyleOption(option)
        self.style().drawComplexControl(QStyle.CC_ToolButton, option, painter, self)




class LoadingSpinner(QWidget):
    def __init__(self, parent=None, diameter=14, line_count=12):
        super().__init__(parent)
        self._angle = 0
        self._line_count = line_count
        self._diameter = diameter
        self.setFixedSize(diameter, diameter)
        self._timer = QTimer(self)
        self._timer.setInterval(70)
        self._timer.timeout.connect(self._tick)
        self.hide()

    def start(self):
        self.show()
        if not self._timer.isActive():
            self._timer.start()

    def stop(self):
        self._timer.stop()
        self.hide()

    def _tick(self):
        self._angle = (self._angle + 1) % self._line_count
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        center = self.rect().center()
        radius = min(self.width(), self.height()) / 2.0 - 1
        line_length = radius * 0.45
        line_width = max(2.0, radius * 0.16)

        for i in range(self._line_count):
            alpha = int(255 * ((i + 1) / self._line_count))
            color = QColor(0, 0, 0, alpha)
            pen = QPen(color)
            pen.setWidthF(line_width)
            pen.setCapStyle(Qt.RoundCap)
            painter.setPen(pen)

            painter.save()
            painter.translate(center)
            painter.rotate((360.0 / self._line_count) * ((i + self._angle) % self._line_count))
            painter.drawLine(0, -radius + line_length, 0, -radius + 1)
            painter.restore()


class MicListWorker(QThread):
    devices_ready = Signal(object)

    def __init__(self, list_devices_callable, parent=None):
        super().__init__(parent)
        self._list_devices_callable = list_devices_callable

    def run(self):
        devices = self._list_devices_callable()
        self.devices_ready.emit(devices)


class ZoomBoxEdit(QPlainTextEdit):
    commit_requested = Signal()
    move_requested = Signal(int, int)
    jump_next_row_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._markers = []
        self.setContextMenuPolicy(Qt.NoContextMenu)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.commit_requested.emit()
            return

        if event.key() == Qt.Key_Down and (event.modifiers() & Qt.ShiftModifier):
            self.jump_next_row_requested.emit()
            return

        if event.key() in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down):
            if event.key() == Qt.Key_Left:
                self.move_requested.emit(0, -1)
            elif event.key() == Qt.Key_Right:
                self.move_requested.emit(0, 1)
            elif event.key() == Qt.Key_Up:
                self.move_requested.emit(-1, 0)
            elif event.key() == Qt.Key_Down:
                self.move_requested.emit(1, 0)
            return

        super().keyPressEvent(event)

    def mousePressEvent(self, event):
        if (
            event.button() == Qt.LeftButton
            and (event.modifiers() & Qt.ShiftModifier)
        ):
            cursor = self.cursorForPosition(event.pos())
            self._add_marker(cursor.position())
            event.accept()
            return
        if (
            event.button() == Qt.RightButton
            and (event.modifiers() & Qt.ShiftModifier)
        ):
            cursor = self.cursorForPosition(event.pos())
            self._remove_marker(cursor.position())
            event.accept()
            return

        super().mousePressEvent(event)

    def contextMenuEvent(self, event):
        event.accept()

    def paintEvent(self, event):
        super().paintEvent(event)

        if not self._markers:
            return

        painter = QPainter(self.viewport())
        painter.setPen(QColor(90, 170, 255))

        for cursor in self._markers:
            rect = self.cursorRect(cursor)
            if rect.isNull():
                continue
            x = rect.x()
            y = rect.y()
            h = rect.height()
            painter.drawLine(x, y, x, y + h)
            painter.drawLine(x + 3, y, x + 3, y + h)

    def _add_marker(self, position):
        for cursor in self._markers:
            if cursor.position() == position:
                return
        cursor = QTextCursor(self.document())
        cursor.setPosition(position)
        self._markers.append(cursor)
        self.viewport().update()

    def _remove_marker(self, position):
        for i, cursor in enumerate(self._markers):
            if cursor.position() == position:
                del self._markers[i]
                self.viewport().update()
                return True
        return False

    def clear_markers(self):
        if not self._markers:
            return
        self._markers.clear()
        self.viewport().update()

    def marker_positions(self):
        if not self._markers:
            return []
        return sorted({cursor.position() for cursor in self._markers})


class EditorPage(QWidget):
    export_requested = Signal(object)  # document
    document_changed = Signal()
    
    def __init__(self, document):
        super().__init__()

        self.document = document
        self.sheet_buttons = []
        self.swap_mode = None
        self.model = TableModel(document)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
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
        self.zoom_box_btn = QPushButton("Zoom Box")
        self.zoom_box_btn.setCheckable(True)
        self.zoom_box_btn.setFixedHeight(32)
        self.zoom_box_btn.toggled.connect(self._toggle_zoom_box)
        ribbon_layout.addWidget(self.zoom_box_btn)

        ribbon_layout.addStretch()

        self.undo_btn = QPushButton("↶ Undo")
        self.undo_btn.setFixedHeight(36)
        self.undo_btn.clicked.connect(self._undo_action)
        self.redo_btn = QPushButton("↷ Redo")
        self.redo_btn.setFixedHeight(36)
        self.redo_btn.clicked.connect(self._redo_action)

        ribbon_layout.addWidget(self.undo_btn)
        ribbon_layout.addWidget(self.redo_btn)

        self.dictate_btn = DictateToolButton()
        dictate_font = QFont("Segoe UI Emoji", self.dictate_btn.font().pointSize())
        self.dictate_btn.setFont(dictate_font)
        self.dictate_btn.setText("🎙️Dictate")
        self.dictate_btn.setFixedHeight(36)
        self.dictate_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.dictate_btn.setPopupMode(QToolButton.MenuButtonPopup)
        self.dictate_btn.clicked.connect(self._toggle_dictate)

        self.dictate_menu = QMenu(self.dictate_btn)
        self.dictate_menu.aboutToShow.connect(self._populate_mic_menu)
        self.dictate_btn.setMenu(self.dictate_menu)
        self._mic_actions = []
        self._mic_worker = None

        self._dictate_pulse = QPropertyAnimation(self.dictate_btn, b"pulse_scale")
        self._dictate_pulse.setStartValue(1.0)
        self._dictate_pulse.setEndValue(1.12)
        self._dictate_pulse.setDuration(700)
        self._dictate_pulse.setEasingCurve(QEasingCurve.InOutSine)
        self._dictate_pulse.setLoopCount(-1)

        self._dictate_idle_style = (
            "QToolButton {"
            "margin: 0px;"
            "padding: 4px 10px 4px 12px;"
            "}"
            "QToolButton::menu-button {"
            "width: 18px;"
            "subcontrol-origin: padding;"
            "subcontrol-position: right center;"
            "border-left: 1px solid rgba(120, 120, 120, 0.4);"
            "}"
        )
        self._dictate_recording_style = (
            "QToolButton {"
            "margin: 0px;"
            "padding: 4px 10px 4px 12px;"
            "background-color: #d9534f;"
            "color: white;"
            "border-radius: 6px;"
            "}"
            "QToolButton::menu-button {"
            "width: 18px;"
            "subcontrol-origin: padding;"
            "subcontrol-position: right center;"
            "border-left: 1px solid rgba(255, 255, 255, 0.35);"
            "}"
        )
        self.dictate_btn.setStyleSheet(self._dictate_idle_style)

        ribbon_layout.addWidget(self.dictate_btn, 0, Qt.AlignVCenter)

        self.dictate_transcribing_spinner = LoadingSpinner(self, diameter=14)
        self.dictate_transcribing_spinner.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.dictate_transcribing_label = QWidget(self)
        transcribe_status_layout = QHBoxLayout(self.dictate_transcribing_label)
        transcribe_status_layout.setContentsMargins(0, 0, 0, 0)
        transcribe_status_layout.setSpacing(6)
        self.dictate_transcribing_text = QLabel("Transcribing please wait.")
        self.dictate_transcribing_text.setStyleSheet("QLabel { background: transparent; }")
        transcribe_status_layout.addWidget(self.dictate_transcribing_spinner)
        transcribe_status_layout.addWidget(self.dictate_transcribing_text)
        self.dictate_transcribing_label.hide()

        self.dictate_menu_spinner = LoadingSpinner(self, diameter=14)
        self.dictate_menu_spinner.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        # 🔹 Export button (RIGHT)
        self.export_btn = QPushButton("Export to Excel")
        self.export_btn.setFixedHeight(36)

        self.export_btn.clicked.connect(
            lambda: self.export_requested.emit(self.document)
        )

        ribbon_layout.addWidget(self.export_btn)

        layout.addWidget(tool_ribbon)


        self.model = TableModel(document)
        if not hasattr(self, "view"):
            self.view = TableView()
        self.view.get_swap_mode = lambda: self.swap_mode
        self.view.clear_swap_mode = self.clear_swap_mode


        self.view.setModel(self.model)
        self.view.block_swap_requested.connect(self.handle_block_swap)

        self.view.drag_swap_requested.connect(self.handle_drag_swap)
        self._restoring_sizes = False
        self._default_row_height = self.view.verticalHeader().defaultSectionSize()
        self._default_col_width = self.view.horizontalHeader().defaultSectionSize()
        self.view.verticalHeader().sectionResized.connect(self._on_row_resized)
        self.view.horizontalHeader().sectionResized.connect(self._on_col_resized)
        self._apply_sheet_sizes()

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self.voice_controller = VoiceController(max_duration_s=90, model_name="base")
        self.voice_controller.recording_started.connect(self._on_dictate_started)
        self.voice_controller.recording_stopped.connect(self._on_dictate_stopped)
        self.voice_controller.transcription_ready.connect(self._on_dictate_transcription_ready)
        self.voice_controller.transcription_error.connect(self._on_dictate_error)
        self.voice_controller.hint_requested.connect(self._show_dictate_hint)
        self.voice_controller.level_changed.connect(self._on_dictate_level)

        self._zoom_box_geometry = None
        self._zoom_box_ratio = (0.7, 0.1)
        self._zoom_syncing = False
        self._zoom_internal_edit = False
        self._saved_edit_triggers = self.view.editTriggers()
        self._enter_moves_right = True
        self._dictate_level_ema = 0.0
        self._dictate_noise_floor = 0.0
        self._dictate_peak_level = 0.0

        self.zoom_box = ZoomBoxEdit(self)
        self.zoom_box.setObjectName("zoomBox")
        self.zoom_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._apply_zoom_box_font_size()

        self.zoom_box.textChanged.connect(self._on_zoom_text_changed)
        self.zoom_box.commit_requested.connect(self._commit_zoom_box)
        self.zoom_box.move_requested.connect(self._move_current_by)
        self.zoom_box.jump_next_row_requested.connect(self._jump_next_row)

        self.zoom_box_host = QWidget()
        self.zoom_box_host.setObjectName("zoomBoxHost")
        zoom_host_layout = QHBoxLayout(self.zoom_box_host)
        zoom_host_layout.setContentsMargins(0, 6, 12, 10)
        zoom_host_layout.addStretch()
        zoom_host_layout.addWidget(self.zoom_box)
        self.enter_toggle = QCheckBox("Press Enter to change Column")
        self.enter_toggle.setChecked(True)
        self.enter_toggle.toggled.connect(self._on_enter_toggle_changed)
        zoom_host_layout.addWidget(self.enter_toggle)
        self.zoom_box_host.hide()
        self.view.set_zoom_box(self.zoom_box)

        sel = self.view.selectionModel()
        sel.currentChanged.connect(self._on_current_changed)
        self.view.selection_finalized.connect(self._sync_zoom_box_to_current)
        self.model.dataChanged.connect(self._on_model_data_changed)
        self.model.layoutChanged.connect(self._on_model_layout_changed)
        self.model.undo_state_changed.connect(self._update_undo_redo_state)
        self._update_undo_redo_state(self.model.can_undo(), self.model.can_redo())

        layout.addWidget(self.view)
        layout.addWidget(self.zoom_box_host)

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

    def _apply_zoom_box_font_size(self):
        font = self.zoom_box.font()
        point_size = font.pointSizeF()
        if point_size <= 0:
            point_size = 10.0
        font.setPointSizeF(point_size + 4.0)
        self.zoom_box.setFont(font)
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
        self._deactivate_swaps()
        self._deactivate_zoom_box()
        self._apply_sheet_sizes()
        self.document_changed.emit()
    def switch_sheet(self, index):
        self.document.active_sheet_index = index
        self.model.layoutChanged.emit()
        self.refresh_sheet_buttons()
        self._deactivate_swaps()
        self._deactivate_zoom_box()
        self._apply_sheet_sizes()
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
        self._deactivate_swaps()
        self._deactivate_zoom_box()
        self._apply_sheet_sizes()
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

    def disarm_swap_rectangle(self):
        self.swap_rectangle_armed = False
        self.view.clearSelection()
    def arm_swap_rectangle(self):
        self.swap_rectangle_armed = True

    def _set_swap_mode(self, mode, checked):
        if checked:
            for other_mode, btn in (
                ("cell", self.swap_cell_btn),
                ("row", self.swap_row_btn),
                ("column", self.swap_col_btn),
            ):
                if other_mode != mode and btn.isChecked():
                    btn.blockSignals(True)
                    btn.setChecked(False)
                    btn.blockSignals(False)
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

    def _toggle_zoom_box(self, checked):
        if checked:
            self._show_zoom_box()
        else:
            self._hide_zoom_box()

    def _deactivate_zoom_box(self):
        if not self.zoom_box_btn.isChecked():
            return
        self.zoom_box_btn.blockSignals(True)
        self.zoom_box_btn.setChecked(False)
        self.zoom_box_btn.blockSignals(False)
        self._hide_zoom_box()

    def _deactivate_swaps(self):
        self.clear_swap_mode()

    def _show_zoom_box(self):
        self._saved_edit_triggers = self.view.editTriggers()
        self.view.setEditTriggers(QTableView.NoEditTriggers)
        self._update_zoom_box_size_from_ratio()
        self.zoom_box_host.show()
        self._ensure_current_index()
        self._sync_zoom_box_to_current()
        self.zoom_box.setFocus()

    def _hide_zoom_box(self):
        if self.zoom_box_host.isVisible():
            self._store_zoom_box_ratio()
            self.zoom_box_host.hide()

        if self._saved_edit_triggers is not None:
            self.view.setEditTriggers(self._saved_edit_triggers)

    def _apply_sheet_sizes(self):
        sheet = self.document.active_sheet
        v_header = self.view.verticalHeader()
        h_header = self.view.horizontalHeader()
        self._restoring_sizes = True
        try:
            self._reset_header_sizes(
                v_header, self.model.rowCount(), self._default_row_height
            )
            self._reset_header_sizes(
                h_header, self.model.columnCount(), self._default_col_width
            )
            for row, height in sheet.row_heights.items():
                v_header.resizeSection(row, height)
            for col, width in sheet.col_widths.items():
                h_header.resizeSection(col, width)
        finally:
            self._restoring_sizes = False

    def _reset_header_sizes(self, header, count, default_size):
        header.setDefaultSectionSize(default_size)
        for index in range(count):
            header.resizeSection(index, default_size)

    def _on_row_resized(self, logical_index, old_size, new_size):
        if self._restoring_sizes:
            return
        sheet = self.document.active_sheet
        if new_size == self._default_row_height:
            sheet.row_heights.pop(logical_index, None)
        else:
            sheet.row_heights[logical_index] = new_size
        self.document_changed.emit()

    def _on_col_resized(self, logical_index, old_size, new_size):
        if self._restoring_sizes:
            return
        sheet = self.document.active_sheet
        if new_size == self._default_col_width:
            sheet.col_widths.pop(logical_index, None)
        else:
            sheet.col_widths[logical_index] = new_size
        self.document_changed.emit()

    def _update_zoom_box_size_from_ratio(self):
        viewport = self.view.viewport().size()
        if viewport.width() <= 0 or viewport.height() <= 0:
            return
        width_ratio, height_ratio = self._zoom_box_ratio
        width = max(240, int(viewport.width() * width_ratio))
        height = max(90, int(viewport.height() * height_ratio))
        width = min(width, max(240, viewport.width() - 40))
        height = min(height, max(90, viewport.height() - 40))
        self.zoom_box.setFixedSize(width, height)
        self.zoom_box_host.setFixedHeight(height + 16)

    def _store_zoom_box_ratio(self):
        viewport = self.view.viewport().size()
        if viewport.width() <= 0 or viewport.height() <= 0:
            return
        self._zoom_box_ratio = (
            self.zoom_box.width() / viewport.width(),
            self.zoom_box.height() / viewport.height()
        )

    def _on_current_changed(self, current, previous):
        if self.voice_controller.is_recording:
            target_index = previous if previous.isValid() else current
            if target_index.isValid():
                target = TranscriptionTarget(target_index.row(), target_index.column())
            else:
                target = TranscriptionTarget(0, 0)
            self.voice_controller.stop_recording(target)

        if not self.zoom_box_host.isVisible():
            return

        if QApplication.mouseButtons() & Qt.LeftButton:
            return

        self._sync_zoom_box_to_index(current)

    def _sync_zoom_box_to_current(self):
        if not self.zoom_box_host.isVisible():
            return
        self._sync_zoom_box_to_index(self.view.currentIndex())
        self.zoom_box.setFocus(Qt.MouseFocusReason)

    def _sync_zoom_box_to_index(self, index):
        if not self.zoom_box_host.isVisible():
            return
        if not index.isValid():
            self._zoom_syncing = True
            self.zoom_box.clear_markers()
            self.zoom_box.setPlainText("")
            self._zoom_syncing = False
            return

        value = self.model.data(index, Qt.EditRole) or ""
        if self.zoom_box.toPlainText() == value:
            return

        self._zoom_syncing = True
        self.zoom_box.clear_markers()
        self.zoom_box.setPlainText(value)
        self._zoom_syncing = False

    def _on_zoom_text_changed(self):
        if not self.zoom_box_host.isVisible():
            return
        if self._zoom_syncing:
            return
        self._push_zoom_text_to_model()

    def _push_zoom_text_to_model(self):
        index = self._ensure_current_index()
        if not index.isValid():
            return

        text = self.zoom_box.toPlainText()
        current = self.model.data(index, Qt.EditRole) or ""
        if text == current:
            return

        self._zoom_internal_edit = True
        self.model.setData(index, text, Qt.EditRole)
        self._zoom_internal_edit = False

    def _commit_zoom_box(self):
        if not self.zoom_box_host.isVisible():
            return
        # Keep active cell text synchronized before any Enter behavior.
        self._push_zoom_text_to_model()
        marker_positions = self.zoom_box.marker_positions()
        if marker_positions:
            self._commit_zoom_box_segments(marker_positions)
            self.zoom_box.clear_markers()
            return

        if self._enter_moves_right:
            self._advance_to_next_cell()
        else:
            self._advance_to_next_row()

    def _advance_to_next_cell(self):
        index = self._ensure_current_index()
        if not index.isValid():
            return

        row, col = index.row(), index.column()
        if col + 1 >= self.model.columnCount():
            if row + 1 < self.model.rowCount():
                col = 0
                row = row + 1
        else:
            col += 1

        self._set_current_index(row, col)

    def _advance_to_next_row(self):
        index = self._ensure_current_index()
        if not index.isValid():
            return
        row, col = index.row(), index.column()
        if row + 1 < self.model.rowCount():
            row = row + 1
        self._set_current_index(row, col)

    def _move_current_by(self, row_delta, col_delta):
        if not self.zoom_box_host.isVisible():
            return
        self._push_zoom_text_to_model()
        index = self._ensure_current_index()
        if not index.isValid():
            return
        self._set_current_index(index.row() + row_delta, index.column() + col_delta)

    def _jump_next_row(self):
        if not self.zoom_box_host.isVisible():
            return
        self._push_zoom_text_to_model()
        index = self._ensure_current_index()
        if not index.isValid():
            return
        self._set_current_index(index.row() + 1, 0)

    def _commit_zoom_box_segments(self, marker_positions):
        text = self.zoom_box.toPlainText()
        length = len(text)
        positions = self._normalized_marker_positions(marker_positions, length)
        if not positions:
            self._push_zoom_text_to_model()
            return

        positions = sorted(set(positions))
        segments = []
        start = 0
        for pos in positions:
            segments.append((start, pos))
            start = pos
        segments.append((start, length))

        index = self._ensure_current_index()
        if not index.isValid():
            return

        targets = self._segment_targets(index, len(segments))
        if not targets:
            return

        if self._targets_have_data(targets):
            reply = QMessageBox.question(
                self,
                "Overwrite Cells?",
                "Cells already contain data. Overwrite?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        self.model.begin_compound_action()
        try:
            for row, col, value in values:
                self.model.setData(self.model.index(row, col), value, Qt.EditRole)
        finally:
            self.model.end_compound_action()

        last_row, last_col = targets[-1]
        self._set_current_index(last_row, last_col)

    def _normalized_marker_positions(self, marker_positions, length):
        if length <= 1:
            return []

        normalized = []
        for pos in marker_positions:
            try:
                p = int(pos)
            except (TypeError, ValueError):
                continue
            # Clamp to valid split points so one Enter consistently works.
            p = max(1, min(p, length - 1))
            normalized.append(p)

        return sorted(set(normalized))

    def _segment_targets(self, index, segment_count):
        row, col = index.row(), index.column()
        if self._enter_moves_right:
            max_cols = self.model.columnCount() - col
            count = min(segment_count, max_cols)
            return [(row, col + i) for i in range(count)]

        max_rows = self.model.rowCount() - row
        count = min(segment_count, max_rows)
        return [(row + i, col) for i in range(count)]

    def _targets_need_overwrite_confirmation(self, values, source_index=None):
        # Overwrite prompts are intentionally disabled for zoom-box segment commits.
        # Keep this no-op helper for backward compatibility with any stale call sites.
        return False

    def _targets_have_data(self, targets):
        # Backward-compatible helper for older call sites that still invoke
        # _targets_have_data during segment commit checks.
        for row, col in targets:
            idx = self.model.index(row, col)
            existing = self.model.data(idx, Qt.EditRole) or ""
            if existing != "":
                return True
        return False

    def _on_enter_toggle_changed(self, checked):
        self._enter_moves_right = checked
        if checked:
            self.enter_toggle.setText("Press Enter to change Column")
        else:
            self.enter_toggle.setText("Press Enter to change Row")

    def _set_current_index(self, row, col):
        row = max(0, min(row, self.model.rowCount() - 1))
        col = max(0, min(col, self.model.columnCount() - 1))
        new_index = self.model.index(row, col)
        sel = self.view.selectionModel()
        sel.setCurrentIndex(
            new_index,
            QItemSelectionModel.ClearAndSelect
        )
        self.view.scrollTo(new_index)

    def _ensure_current_index(self):
        index = self.view.currentIndex()
        if index.isValid():
            return index

        index = self.model.index(0, 0)
        sel = self.view.selectionModel()
        sel.setCurrentIndex(
            index,
            QItemSelectionModel.ClearAndSelect
        )
        return index

    def _on_model_data_changed(self, top_left, bottom_right, roles=None):
        if not self.zoom_box_host.isVisible():
            return
        if self._zoom_internal_edit:
            return

        current = self.view.currentIndex()
        if not current.isValid():
            return

        in_range = (
            top_left.row() <= current.row() <= bottom_right.row()
            and top_left.column() <= current.column() <= bottom_right.column()
        )
        if in_range:
            self._sync_zoom_box_to_index(current)

    def _on_model_layout_changed(self):
        if not self.zoom_box_host.isVisible():
            return
        self._sync_zoom_box_to_current()

    def _undo_action(self):
        self.model.undo()

    def _redo_action(self):
        self.model.redo()

    def _update_undo_redo_state(self, can_undo, can_redo):
        if hasattr(self, "undo_btn"):
            self.undo_btn.setEnabled(can_undo)
        if hasattr(self, "redo_btn"):
            self.redo_btn.setEnabled(can_redo)

    def _toggle_dictate(self):
        if self.voice_controller.is_transcribing:
            self._show_dictate_hint("Transcribing... please wait.")
            return

        index = self._ensure_current_index()
        target = TranscriptionTarget(index.row(), index.column())

        if self.voice_controller.is_recording:
            self._show_transcribing_status(True)
            self.voice_controller.stop_recording(target)
        else:
            self._show_transcribing_status(False)
            self.voice_controller.start_recording(target)

    def _populate_mic_menu(self):
        if self._mic_worker is not None and self._mic_worker.isRunning():
            return

        self._show_dictate_menu_loading(True)
        self.dictate_menu.clear()
        loading_action = self.dictate_menu.addAction("Loading microphones...")
        loading_action.setEnabled(False)

        self._mic_worker = MicListWorker(self.voice_controller.list_devices, self)
        self._mic_worker.devices_ready.connect(self._on_mic_devices_loaded)
        self._mic_worker.finished.connect(self._on_mic_worker_finished)
        self._mic_worker.start()

    def _on_mic_devices_loaded(self, devices):
        self.dictate_menu.clear()
        self._mic_actions = []

        default_action = self.dictate_menu.addAction("Default system mic")
        default_action.setCheckable(True)
        default_action.setData(None)
        if self.voice_controller.selected_device_id is None:
            default_action.setChecked(True)
        default_action.triggered.connect(
            lambda checked=False, action=default_action: self._select_microphone_action(None, action)
        )
        self._mic_actions.append(default_action)

        if not devices:
            no_mics = self.dictate_menu.addAction("No mics connected")
            no_mics.setEnabled(False)
        else:
            for device in devices:
                action = self.dictate_menu.addAction(device.name)
                action.setCheckable(True)
                action.setData(device.device_id)
                if device.device_id == self.voice_controller.selected_device_id:
                    action.setChecked(True)
                action.triggered.connect(
                    lambda checked=False, device_id=device.device_id, action=action: (
                        self._select_microphone_action(device_id, action)
                    )
                )
                self._mic_actions.append(action)

        if self.voice_controller.is_recording:
            for action in self.dictate_menu.actions():
                action.setEnabled(False)

        self._show_dictate_menu_loading(False)

    def _on_mic_worker_finished(self):
        if self._mic_worker is not None:
            self._mic_worker.deleteLater()
            self._mic_worker = None

    def _select_microphone(self, device_id):
        self.voice_controller.set_selected_device(device_id)

    def _select_microphone_action(self, device_id, action):
        self._select_microphone(device_id)
        self._update_mic_checks(action)

    def _update_mic_checks(self, selected_action):
        for action in self._mic_actions:
            action.setChecked(action is selected_action)

    def _on_dictate_started(self):
        self.dictate_btn.setText("⏺Dictate")
        self.dictate_btn.setStyleSheet(self._dictate_recording_style)
        self._dictate_pulse.stop()
        self._dictate_level_ema = 0.0
        self._dictate_noise_floor = 0.0
        self._dictate_peak_level = 0.0

    def _on_dictate_stopped(self):
        self._dictate_pulse.stop()
        self.dictate_btn.setStyleSheet(self._dictate_idle_style)
        self.dictate_btn.setText("🎙️Dictate")
        self.dictate_btn.pulse_scale = 1.0
        self._dictate_level_ema = 0.0
        self._dictate_noise_floor = 0.0
        self._dictate_peak_level = 0.0

    def _on_dictate_transcription_ready(self, text, target):
        self._show_transcribing_status(False)
        self._append_text_to_cell(text, target)

    def _on_dictate_error(self, message):
        self._show_transcribing_status(False)
        self._show_dictate_hint(f"Transcription failed: {message}")

    def _show_transcribing_status(self, visible):
        if visible:
            self.dictate_transcribing_spinner.start()
            self.dictate_transcribing_label.show()
            self.dictate_transcribing_label.raise_()
        else:
            self.dictate_transcribing_spinner.stop()
            self.dictate_transcribing_label.hide()
        self._position_dictate_status_widgets()

    def _show_dictate_menu_loading(self, visible):
        if visible:
            self.dictate_menu_spinner.start()
            self.dictate_menu_spinner.raise_()
            QApplication.processEvents()
        else:
            self.dictate_menu_spinner.stop()
        self._position_dictate_status_widgets()

    def _position_dictate_status_widgets(self):
        if not hasattr(self, "dictate_btn"):
            return

        btn_top_left = self.dictate_btn.mapTo(self, QPoint(0, 0))
        btn_rect = self.dictate_btn.geometry()
        gap_y = btn_top_left.y() + btn_rect.height() + 6

        if hasattr(self, "dictate_transcribing_label"):
            self.dictate_transcribing_label.adjustSize()
            trans_w = self.dictate_transcribing_label.width()
            x = btn_top_left.x() + max(0, (btn_rect.width() - trans_w) // 2)
            self.dictate_transcribing_label.move(x, gap_y)

        if hasattr(self, "dictate_menu_spinner"):
            spinner_x = btn_top_left.x() + btn_rect.width() - self.dictate_menu_spinner.width() - 2
            spinner_y = btn_top_left.y() + btn_rect.height() + 6
            self.dictate_menu_spinner.move(spinner_x, spinner_y)

    def _show_dictate_hint(self, message):
        rect = self.dictate_btn.rect()
        pos = self.dictate_btn.mapToGlobal(
            QPoint(rect.center().x(), rect.bottom() + 12)
        )
        QToolTip.showText(pos, message, self.dictate_btn)

    def _on_dictate_level(self, level):
        if not self.voice_controller.is_recording:
            return

        self._dictate_level_ema = (self._dictate_level_ema * 0.7) + (level * 0.3)

        if self._dictate_noise_floor <= 0.0:
            self._dictate_noise_floor = self._dictate_level_ema
        else:
            self._dictate_noise_floor = (
                self._dictate_noise_floor * 0.97
                + self._dictate_level_ema * 0.03
            )

        self._dictate_peak_level = max(
            self._dictate_peak_level * 0.98,
            self._dictate_level_ema,
        )

        dynamic_span = max(self._dictate_peak_level - self._dictate_noise_floor, 1e-6)
        relative_level = (self._dictate_level_ema - self._dictate_noise_floor) / dynamic_span
        is_voice_active = (
            self._dictate_level_ema > 0.0012
            and relative_level > 0.25
        )

        if is_voice_active:
            if self._dictate_pulse.state() != QPropertyAnimation.Running:
                self._dictate_pulse.start()
        else:
            if self._dictate_pulse.state() == QPropertyAnimation.Running:
                self._dictate_pulse.stop()
                self.dictate_btn.pulse_scale = 1.0

    def _append_text_to_cell(self, text, target):
        if not text:
            return
        index = self.model.index(target.row, target.column)
        if not index.isValid():
            return
        current = self.model.data(index, Qt.EditRole) or ""
        separator = ""
        if current and text and not current.endswith((" ", "\n")) and not text.startswith(" "):
            separator = " "
        updated = f"{current}{separator}{text}"
        self.model.setData(index, updated, Qt.EditRole)

    def showEvent(self, event):
        super().showEvent(event)
        if self.zoom_box_btn.isChecked():
            self._show_zoom_box()
        self._position_dictate_status_widgets()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.zoom_box_host.isVisible():
            self._update_zoom_box_size_from_ratio()
        self._position_dictate_status_widgets()

    def hideEvent(self, event):
        if self.zoom_box_host.isVisible():
            self._store_zoom_box_ratio()
            self.zoom_box_host.hide()
        if self._saved_edit_triggers is not None:
            self.view.setEditTriggers(self._saved_edit_triggers)
        super().hideEvent(event)

    def apply_grid_dark_mode(self, enabled: bool):
        if not enabled:
            self.setStyleSheet("""
            QWidget {
                background-color: #f7f8fa;
                color: #111827;
            }

            QWidget#editorRibbon {
                background-color: #f9fafb;
                border-bottom: 1px solid #e5e7eb;
            }

            QWidget#editorRibbon QPushButton,
            QWidget#editorRibbon QToolButton {
                background-color: #ffffff;
                color: #111827;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                padding: 6px 14px;
                font-weight: 500;
            }

            QWidget#editorRibbon QPushButton:hover,
            QWidget#editorRibbon QToolButton:hover {
                background-color: #f3f4f6;
                border: 1px solid #cbd5e1;
            }

            QWidget#editorRibbon QPushButton:focus,
            QWidget#editorRibbon QToolButton:focus {
                border: 1px solid #256d85;
            }

            QWidget#editorRibbon QPushButton:checked {
                background-color: #256d85;
                color: #ffffff;
                border: 1px solid #256d85;
            }

            QPushButton[sheetButton="true"] {
                background-color: #f3f4f6;
                color: #4b5563;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                padding: 0 12px;
                font-weight: 400;
            }

            QPushButton[sheetButton="true"]:checked {
                background-color: #256d85;
                color: #ffffff;
                border: 1px solid #256d85;
                font-weight: 500;
            }

            QPushButton[sheetButton="true"]:hover {
                background-color: #ffffff;
                border: 1px solid #d1d5db;
            }

            QWidget#sheetBar QPushButton {
                background-color: #f3f4f6;
                color: #4b5563;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
            }

            QTableView {
                background-color: #ffffff;
                gridline-color: #e5e7eb;
                color: #111827;
                selection-background-color: transparent;
                selection-color: #111827;
                border: 1px solid #e5e7eb;
            }

            QTableView::item:selected {
                background-color: transparent;
                color: #111827;
            }

            QAbstractScrollArea::viewport {
                background-color: #ffffff;
            }

            QAbstractScrollArea::corner {
                background: #f7f8fa;
            }

            QHeaderView {
                background-color: #f9fafb;
            }

            QHeaderView::section {
                background-color: #f3f4f6;
                color: #374151;
                border: 1px solid #e5e7eb;
                border-bottom: 1px solid #d1d5db;
                padding: 4px;
            }

            QTableCornerButton::section,
            QTableView QTableCornerButton::section {
                background-color: #f9fafb;
                border: 1px solid #e5e7eb;
            }

            QWidget#sheetBar {
                background-color: #f7f8fa;
                border-top: 1px solid #e5e7eb;
            }

            QWidget#zoomBoxHost {
                background-color: #f3f4f6;
                border-top: 1px solid #e5e7eb;
            }

            QPlainTextEdit#zoomBox {
                background-color: #ffffff;
                color: #111827;
                border: 1px solid #d1d5db;
                border-radius: 8px;
            }

            QScrollBar:vertical, QScrollBar:horizontal {
                background: #f3f4f6;
                height: 10px;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
                background: #d1d5db;
                min-height: 24px;
                min-width: 24px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
                background: #9ca3af;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: #f3f4f6;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical,
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                height: 0px;
                width: 0px;
            }
            """)
            return

        self.setStyleSheet("""
        QWidget {
            background-color: #202124;
            color: #eaeaea;
        }

        QWidget#editorRibbon {
            background-color: #252525;
            border-bottom: 1px solid rgba(255, 255, 255, 0.06);
        }

        QWidget#editorRibbon QPushButton,
        QWidget#editorRibbon QToolButton {
            background-color: #2a2a2a;
            color: #eaeaea;
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 8px;
            padding: 6px 14px;
            font-weight: 500;
        }

        QWidget#editorRibbon QToolButton {
            background-color: #2d2d30;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 6px;
            padding: 4px 12px;
        }

        QWidget#editorRibbon QToolButton {
            background-color: #2d2d30;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 6px;
            padding: 4px 12px;
        }

        QWidget#editorRibbon QToolButton {
            background-color: #2d2d30;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 6px;
            padding: 4px 12px;
        }

        QWidget#editorRibbon QToolButton {
            background-color: #2d2d30;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 6px;
            padding: 4px 12px;
        }

        QWidget#editorRibbon QToolButton {
            background-color: #2d2d30;
            color: #e6e6e6;
            border: 1px solid #3a3a3a;
            border-radius: 6px;
            padding: 4px 12px;
        }

        QWidget#editorRibbon QPushButton:hover {
            background-color: #3a3a3a;
        }

        QWidget#editorRibbon QToolButton:hover {
            background-color: #3a3a3a;
        }

        QWidget#editorRibbon QPushButton:checked {
            background-color: #256d85;
            color: #ffffff;
            border: 1px solid #256d85;
        }

        QWidget#editorRibbon QPushButton:disabled,
        QWidget#editorRibbon QToolButton:disabled {
            color: #a0a0a0;
            background-color: #2a2a2a;
        }

        QWidget#editorRibbon QToolButton:disabled {
            color: #9e9e9e;
            background-color: #2a2a2a;
        }

        QWidget#editorRibbon QToolButton:disabled {
            color: #9e9e9e;
            background-color: #2a2a2a;
        }

        QWidget#editorRibbon QToolButton:disabled {
            color: #9e9e9e;
            background-color: #2a2a2a;
        }

        QWidget#editorRibbon QToolButton:disabled {
            color: #9e9e9e;
            background-color: #2a2a2a;
        }

        QWidget#editorRibbon QToolButton:disabled {
            color: #9e9e9e;
            background-color: #2a2a2a;
        }

        /* ===============================
        SHEET BUTTONS (NOT QTabBar!)
        =============================== */
        QPushButton[sheetButton="true"] {
            background-color: #2a2a2a;
            color: #a0a0a0;
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 6px;
            padding: 0 12px;
            font-weight: 400;
        }

        QPushButton[sheetButton="true"]:checked {
            background-color: #256d85;
            color: #ffffff;
            border: 1px solid #256d85;
            font-weight: 500;
        }

        QPushButton[sheetButton="true"]:hover {
            background-color: #2e2e2e;
            border: 1px solid rgba(255, 255, 255, 0.10);
        }

        QWidget#sheetBar QPushButton {
            background-color: #2a2a2a;
            color: #a0a0a0;
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 6px;
        }

        QTableView {
            background-color: #252525;
            gridline-color: rgba(255, 255, 255, 0.06);
            color: #eaeaea;
            selection-background-color: transparent;
            selection-color: #eaeaea;
            border: 1px solid rgba(255, 255, 255, 0.06);
        }

        QTableView::item:selected {
            background-color: transparent;
            color: #eaeaea;
        }

        QAbstractScrollArea::viewport {
            background-color: #252525;
        }

        QAbstractScrollArea::corner {
            background: #202124;
        }

        QHeaderView {
            background-color: #202124;
        }

        QHeaderView::section {
            background-color: #2a2a2a;
            color: #eaeaea;
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-bottom: 1px solid rgba(255, 255, 255, 0.10);
            padding: 4px;
        }

        QTableCornerButton::section,
        QTableView QTableCornerButton::section {
            background-color: #202124;
            border: 1px solid rgba(255, 255, 255, 0.06);
        }

        QWidget#sheetBar {
            background-color: #202124;
            border-top: 1px solid rgba(255, 255, 255, 0.06);
        }

        QWidget#zoomBoxHost {
            background-color: #2a2a2a;
            border-top: 1px solid rgba(255, 255, 255, 0.06);
        }

        QPlainTextEdit#zoomBox {
            background-color: #202124;
            color: #eaeaea;
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 8px;
        }

        QScrollBar:vertical, QScrollBar:horizontal {
            background: #202124;
            height: 10px;
            width: 10px;
            margin: 0px;
        }
        QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
            background: #2e2e2e;
            min-height: 24px;
            min-width: 24px;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
            background: #3a3a3a;
        }
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
            background: #202124;
        }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical,
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {
            height: 0px;
            width: 0px;
        }
        """)
