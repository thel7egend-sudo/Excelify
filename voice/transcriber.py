from threading import Lock
from typing import Dict, Optional

import numpy as np
from faster_whisper import WhisperModel


_MODEL_CACHE: Dict[str, WhisperModel] = {}
_MODEL_LOCK = Lock()
_ACCURACY_MODEL_NAME = "medium"


def _get_model(model_name: str) -> WhisperModel:
    with _MODEL_LOCK:
        model = _MODEL_CACHE.get(model_name)
        if model is None:
            model = WhisperModel(model_name, device="cpu", compute_type="int8")
            _MODEL_CACHE[model_name] = model
        return model


class WhisperTranscriber:
    def __init__(self, model_name: str = "base"):
        # Force the transcription model to medium for better recognition accuracy.
        self.model_name = _ACCURACY_MODEL_NAME

    def transcribe(self, audio: np.ndarray, samplerate: int) -> str:
        if audio.size == 0:
            return ""
        if audio.ndim > 1:
            audio = np.mean(audio, axis=1)
        if audio.dtype != np.float32:
            audio = audio.astype(np.float32)

        model = _get_model(self.model_name)
        segments, _ = model.transcribe(
            audio,
            language="en",
            task="transcribe",
            beam_size=7,
            temperature=0.0,
            condition_on_previous_text=True,
            vad_filter=True,
        )
        parts = [segment.text for segment in segments]
        return "".join(parts).strip()
