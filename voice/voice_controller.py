import re

from PySide6.QtCore import QObject, QTimer, Signal

from .audio_stream import AudioStream
from .numeric_recognizer import NumericRecognizer

NON_DIGIT_REGEX = re.compile(r"[^0-9]")


class VoiceController(QObject):
    transcription_ready = Signal(str, tuple)
    recording_state_changed = Signal(bool)
    audio_level_changed = Signal(float)
    error = Signal(str)

    def __init__(self, parent=None, sample_rate=16000, silence_threshold=0.012):
        super().__init__(parent)
        self.sample_rate = sample_rate
        self.silence_threshold = float(silence_threshold)
        self.recognizer = None
        self.audio_stream = None
        self.target_cell = None
        self._recording = False
        self._last_level = 0.0
        self._recognized_fragments = []

        self._pump_timer = QTimer(self)
        self._pump_timer.setInterval(30)
        self._pump_timer.timeout.connect(self._pump_audio)

    @property
    def is_recording(self):
        return self._recording

    def start_recording(self, target_cell):
        if self._recording:
            return True

        try:
            if self.recognizer is None:
                self.recognizer = NumericRecognizer(sample_rate=self.sample_rate)
            else:
                self.recognizer.reset()
            self._recognized_fragments.clear()

            self.audio_stream = AudioStream(
                sample_rate=self.sample_rate,
                channels=1,
                dtype="float32",
                level_callback=self._on_level,
                error_callback=self._on_stream_error,
            )
            self.audio_stream.start()
        except Exception as exc:
            self.error.emit(str(exc))
            self.audio_stream = None
            self._recording = False
            return False

        self.target_cell = tuple(target_cell)
        self._recording = True
        self._pump_timer.start()
        self.recording_state_changed.emit(True)
        return True

    def stop_recording(self):
        if not self._recording:
            return

        self._pump_timer.stop()
        self._pump_audio()

        final_digits = ""
        if self.recognizer is not None:
            print("[Dictate] Calling FinalResult()")
            final_digits = self.recognizer.finalize()

        if self.audio_stream is not None:
            self.audio_stream.stop()
            self.audio_stream = None

        target = self.target_cell
        self.target_cell = None
        self._recording = False
        self.recording_state_changed.emit(False)
        self.audio_level_changed.emit(0.0)
        self._last_level = 0.0

        text = self._digits_only("".join(self._recognized_fragments) + final_digits)
        print(f"[Dictate] transcription_ready target={target} text='{text}'")
        self.transcription_ready.emit(text, target)

        if self.recognizer is not None:
            self.recognizer.reset()
        self._recognized_fragments.clear()

    def finalize_current_cell(self):
        self.stop_recording()

    def restart_for_next_cell(self, new_target):
        self.stop_recording()
        self.start_recording(new_target)

    def _pump_audio(self):
        if not self._recording or self.audio_stream is None or self.recognizer is None:
            return

        for chunk in self.audio_stream.pop_all_chunks():
            accepted, text = self.recognizer.accept_audio(chunk)
            if accepted:
                print(f"[Dictate] AcceptWaveform=True text='{text}'")
                self._recognized_fragments.append(text)

        level = self._last_level if self._last_level >= self.silence_threshold else 0.0
        self.audio_level_changed.emit(level)

    def _on_level(self, rms):
        self._last_level = max(0.0, min(float(rms), 1.0))

    def _on_stream_error(self, message):
        self.error.emit(message)

    @staticmethod
    def _digits_only(text):
        return NON_DIGIT_REGEX.sub("", text or "")
