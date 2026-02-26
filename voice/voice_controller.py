from dataclasses import dataclass
from collections import deque
from typing import Optional, Tuple
import time
import hashlib

import numpy as np
from PySide6.QtCore import QObject, Signal, QThread, QTimer

from .recorder import AudioRecorder, list_input_devices
from .transcriber import WhisperTranscriber


@dataclass(frozen=True)
class TranscriptionTarget:
    row: int
    column: int


class TranscriptionWorker(QThread):
    result_ready = Signal(str, object)
    error = Signal(str)

    def __init__(self, audio: np.ndarray, samplerate: int, model_name: str, target: TranscriptionTarget):
        super().__init__()
        self._audio = audio
        self._samplerate = samplerate
        self._model_name = model_name
        self._target = target

    def run(self):
        try:
            transcriber = WhisperTranscriber(model_name=self._model_name)
            text = transcriber.transcribe(self._audio, self._samplerate)
            self.result_ready.emit(text, self._target)
        except Exception as exc:
            self.error.emit(str(exc))


class VoiceController(QObject):
    recording_started = Signal()
    recording_stopped = Signal()
    transcription_ready = Signal(str, object)
    transcription_error = Signal(str)
    hint_requested = Signal(str)
    level_changed = Signal(float)
    transcription_idle = Signal()

    def __init__(self, max_duration_s: int = 90, model_name: str = "base"):
        super().__init__()
        self._recorder = AudioRecorder()
        self._max_duration_s = max_duration_s
        self._model_name = model_name
        self._selected_device_id: Optional[int] = None
        self._timeout_timer = QTimer(self)
        self._timeout_timer.setSingleShot(True)
        self._timeout_timer.timeout.connect(self._handle_timeout)
        self._active_worker: Optional[TranscriptionWorker] = None
        self._pending_jobs = deque()
        self._pending_signatures = set()
        self._active_signature: Optional[str] = None
        self._recording_target: Optional[TranscriptionTarget] = None
        self._chunk_timer = QTimer(self)
        self._chunk_timer.setInterval(3000)
        self._chunk_timer.timeout.connect(self._flush_current_chunk)
        self._level_timer = QTimer(self)
        self._level_timer.setInterval(80)
        self._level_timer.timeout.connect(self._emit_level)
        self._recorder.set_level_callback(self._handle_level)
        self._latest_level = 0.0
        self._silence_rms_threshold = 0.01
        self._silence_flush_seconds = 0.24
        self._silence_started_at: Optional[float] = None
        self._min_chunk_seconds = 0.8
        self._min_chunk_rms = 0.005

    @property
    def is_recording(self) -> bool:
        return self._recorder.is_recording

    @property
    def is_transcribing(self) -> bool:
        if self._active_worker is None:
            return bool(self._pending_jobs)
        if self._active_worker.isRunning():
            return True

        # Defensive cleanup for stale worker references so UI does not get
        # stuck in a "transcribing" state after merge/runtime edge cases.
        self._active_worker.deleteLater()
        self._active_worker = None
        return False

    @property
    def selected_device_id(self) -> Optional[int]:
        return self._selected_device_id

    def set_selected_device(self, device_id: Optional[int]) -> None:
        self._selected_device_id = device_id

    def list_devices(self):
        return list_input_devices()

    def start_recording(self, target: TranscriptionTarget) -> None:
        if self.is_recording or self.is_transcribing:
            return
        self._recording_target = target
        try:
            self._recorder.start(device_id=self._selected_device_id)
        except Exception as exc:
            self._recording_target = None
            self.transcription_error.emit(str(exc))
            return
        self._level_timer.start()
        self._chunk_timer.start()
        self._silence_started_at = None
        self.recording_started.emit()

    def stop_recording(self, target: TranscriptionTarget, hint: Optional[str] = None) -> None:
        if not self.is_recording:
            return
        self.flush_recording_segment(target)
        self._level_timer.stop()
        self._chunk_timer.stop()
        self._timeout_timer.stop()
        self._recorder.stop()
        self._silence_started_at = None
        self.recording_stopped.emit()
        self._recording_target = None
        if hint:
            self.hint_requested.emit(hint)
        self._start_next_job_if_idle()

    def flush_recording_segment(self, target: Optional[TranscriptionTarget] = None) -> None:
        if not self.is_recording:
            return
        segment_target = target or self._recording_target
        if segment_target is None:
            return
        audio = self._recorder.consume_audio_chunk()
        if audio.size == 0:
            return

        if not self._is_valid_chunk(audio):
            return

        signature = self._audio_signature(audio)
        if signature in self._pending_signatures or signature == self._active_signature:
            return

        self._pending_signatures.add(signature)
        self._pending_jobs.append((signature, audio, segment_target))
        self._start_next_job_if_idle()

    def _flush_current_chunk(self) -> None:
        self.flush_recording_segment(self._recording_target)

    def _start_transcription(self, audio: np.ndarray, target: TranscriptionTarget) -> None:
        if self._active_worker is not None and self._active_worker.isRunning():
            return
        if self._active_worker is not None:
            self._active_worker.deleteLater()
            self._active_worker = None
        worker = TranscriptionWorker(audio, self._recorder.samplerate, self._model_name, target)
        worker.result_ready.connect(self._handle_result)
        worker.error.connect(self._handle_error)
        worker.finished.connect(self._handle_worker_finished)
        self._active_worker = worker
        worker.start()

    def _start_next_job_if_idle(self) -> None:
        if self._active_worker is not None and self._active_worker.isRunning():
            return
        if self._active_worker is not None:
            self._active_worker.deleteLater()
            self._active_worker = None
        if not self._pending_jobs:
            self.transcription_idle.emit()
            return
        signature, audio, target = self._pending_jobs.popleft()
        self._pending_signatures.discard(signature)
        self._active_signature = signature
        self._start_transcription(audio, target)

    def _handle_result(self, text: str, target: TranscriptionTarget) -> None:
        self.transcription_ready.emit(text, target)

    def _handle_error(self, message: str) -> None:
        self.transcription_error.emit(message)

    def _handle_worker_finished(self) -> None:
        self._active_signature = None
        if self._active_worker is not None:
            self._active_worker.deleteLater()
            self._active_worker = None
        self._start_next_job_if_idle()

    def _is_valid_chunk(self, audio: np.ndarray) -> bool:
        if audio.ndim > 1:
            audio = np.mean(audio, axis=1)
        if audio.dtype != np.float32:
            audio = audio.astype(np.float32)

        min_samples = max(int(self._recorder.samplerate * self._min_chunk_seconds), 1)
        if audio.shape[0] < min_samples:
            return False

        rms = float(np.sqrt(np.mean(audio**2)))
        return rms >= self._min_chunk_rms

    def _audio_signature(self, audio: np.ndarray) -> str:
        contiguous = np.ascontiguousarray(audio)
        return hashlib.blake2b(contiguous.tobytes(), digest_size=16).hexdigest()

    def _handle_timeout(self) -> None:
        if not self.is_recording:
            return
        hint = f"Stopped after {self._max_duration_s}s to process audio."
        target = self._recording_target
        if target is None:
            target = TranscriptionTarget(0, 0)
        self.stop_recording(target, hint=hint)

    def set_recording_target(self, target: TranscriptionTarget) -> None:
        self._recording_target = target

    def _handle_level(self, level: float) -> None:
        self._latest_level = level
        if not self.is_recording:
            return
        now = time.monotonic()
        if level < self._silence_rms_threshold:
            if self._silence_started_at is None:
                self._silence_started_at = now
                return
            if now - self._silence_started_at >= self._silence_flush_seconds:
                self.flush_recording_segment(self._recording_target)
                self._silence_started_at = now
        else:
            self._silence_started_at = None

    def _emit_level(self) -> None:
        if not self.is_recording:
            return
        self.level_changed.emit(self._latest_level)
