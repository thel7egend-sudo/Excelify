import json
from threading import Lock
from typing import Dict

import numpy as np
from vosk import KaldiRecognizer, Model


_VOSK_MODEL_PATH = "models/vosk-model-small-en-us-0.15"
_GRAMMAR_TOKENS = [
    "a",
    "b",
    "c",
    "d",
    "e",
    "f",
    "g",
    "h",
    "i",
    "j",
    "k",
    "l",
    "m",
    "n",
    "o",
    "p",
    "q",
    "r",
    "s",
    "t",
    "u",
    "v",
    "w",
    "x",
    "y",
    "z",
    "zero",
    "one",
    "two",
    "three",
    "four",
    "five",
    "six",
    "seven",
    "eight",
    "nine",
]
_WORD_DIGIT_MAP = {
    "zero": "0",
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9",
}

_MODEL_CACHE: Dict[str, Model] = {}
_MODEL_LOCK = Lock()


def _get_model(model_name: str) -> Model:
    with _MODEL_LOCK:
        model = _MODEL_CACHE.get(model_name)
        if model is None:
            model = Model(_VOSK_MODEL_PATH)
            _MODEL_CACHE[model_name] = model
        return model


def _normalize_text(text: str) -> str:
    tokens = text.lower().split()
    normalized = []

    for token in tokens:
        if token in _WORD_DIGIT_MAP:
            normalized.append(_WORD_DIGIT_MAP[token])
        elif len(token) == 1 and token.isalpha():
            normalized.append(token)
        elif len(token) == 1 and token.isdigit():
            normalized.append(token)

    return "".join(normalized)


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

        audio = np.clip(audio, -1.0, 1.0)
        audio_int16 = (audio * 32767).astype(np.int16)

        model = _get_model(self.model_name)
        recognizer = KaldiRecognizer(model, samplerate, json.dumps(_GRAMMAR_TOKENS))
        recognizer.SetWords(False)
        recognizer.AcceptWaveform(audio_int16.tobytes())

        result_json = recognizer.FinalResult()
        result_dict = json.loads(result_json)
        text = result_dict.get("text", "")

        return _normalize_text(text)
