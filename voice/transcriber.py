from threading import Lock
from typing import Dict

import numpy as np
from faster_whisper import WhisperModel

_MODEL_CACHE: Dict[str, WhisperModel] = {}
_MODEL_LOCK = Lock()


def _get_model() -> WhisperModel:
    with _MODEL_LOCK:
        model = _MODEL_CACHE.get("base")
        if model is None:
            model = WhisperModel("base", device="cpu", compute_type="float32")
            _MODEL_CACHE["base"] = model
        return model


class WhisperTranscriber:
    def __init__(self, model_name: str = "base"):
        self.model_name = model_name

    def transcribe(self, audio: np.ndarray, samplerate: int) -> str:
        if audio.size == 0:
            return ""

        if audio.ndim > 1:
            audio = np.mean(audio, axis=1)
        if audio.dtype != np.float32:
            audio = audio.astype(np.float32)

        min_samples = max(int(samplerate * 0.8), 1)
        if audio.shape[0] < min_samples:
            return ""

        rms = float(np.sqrt(np.mean(audio**2)))
        if rms < 0.005:
            return ""

        model = _get_model()
        segments, _ = model.transcribe(
            audio,
            language="en",
            task="transcribe",
            beam_size=5,
            temperature=0.0,
            condition_on_previous_text=False,
            vad_filter=False,
        )
        return "".join([segment.text for segment in segments]).strip()
