import json
from threading import Lock
from typing import Dict, List

import numpy as np
from vosk import KaldiRecognizer, Model

import os

_BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
_VOSK_MODEL_PATH = os.path.join(_BASE_DIR, "models", "vosk-model-small-en-us-0.15")
_GRAMMAR_TOKENS = [
    "a",
    "b",
    "f",
    "h",
    "i",
    "j",
    "l",
    "n",
    "o",
    "r",
    "s",
    "u",
    "w",
    "x",
    "y",
    "z",
    "cha",
    "delta",
    "echo",
    "gamma",
    "pulse",
    "false",
    "plus",
    "kilo",
    "mike",
    "truck",
    "track",
    "tuck",
    "vector",
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
_TOKEN_NORMALIZATION_MAP = {
    "a": "a",
    "b": "b",
    "f": "f",
    "h": "h",
    "i": "i",
    "j": "j",
    "l": "l",
    "n": "n",
    "o": "o",
    "r": "r",
    "s": "s",
    "u": "u",
    "w": "w",
    "x": "x",
    "y": "y",
    "z": "z",
    "cha": "c",
    "delta": "d",
    "echo": "e",
    "gamma": "g",
    "pulse": "p",
    "false": "p",
    "plus": "p",
    "kilo": "k",
    "k": "j",
    "mike": "m",
    "m": "n",
    "truck": "t",
    "track": "t",
    "tuck": "t",
    "t": "t",
    "vector": "v",
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
        elif token in _TOKEN_NORMALIZATION_MAP:
            normalized.append(_TOKEN_NORMALIZATION_MAP[token])
        elif len(token) == 1 and token.isdigit():
            normalized.append(token)

    return "".join(normalized)


def _extract_text(result_json: str) -> str:
    try:
        result_dict = json.loads(result_json)
    except json.JSONDecodeError:
        return ""
    return result_dict.get("text", "")


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
        step_samples = max(int(samplerate * 0.2), 1)
        tokens: List[str] = []

        for start in range(0, audio_int16.shape[0], step_samples):
            end = start + step_samples
            chunk = audio_int16[start:end]
            if chunk.size == 0:
                continue
            if recognizer.AcceptWaveform(chunk.tobytes()):
                text = _extract_text(recognizer.Result())
                if text:
                    tokens.append(text)

        final_text = _extract_text(recognizer.FinalResult())
        if final_text:
            tokens.append(final_text)

        return _normalize_text(" ".join(tokens))
