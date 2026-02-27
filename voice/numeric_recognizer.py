import json
import os
import re
from pathlib import Path

from vosk import KaldiRecognizer, Model

DIGIT_ONLY_PATTERN = re.compile(r"[^0-9]")
DIGIT_GRAMMAR = json.dumps(["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"])

_DEFAULT_MODEL_PATH = Path(__file__).resolve().parent.parent / "models" / "vosk-model-small-en-us-0.15"
_MODEL_CACHE = {}


class NumericRecognizer:
    def __init__(self, sample_rate=16000, model_path=None):
        self.sample_rate = sample_rate
        self.model_path = str(model_path or _DEFAULT_MODEL_PATH)
        self.model = self._load_model(self.model_path)
        self.recognizer = KaldiRecognizer(self.model, self.sample_rate, DIGIT_GRAMMAR)
        print(f"[Dictate] KaldiRecognizer initialized sr={self.sample_rate} grammar={DIGIT_GRAMMAR}")

    @staticmethod
    def _load_model(model_path):
        model_path = os.path.abspath(model_path)
        cached = _MODEL_CACHE.get(model_path)
        if cached is None:
            if not os.path.isdir(model_path):
                raise FileNotFoundError(f"Vosk model folder not found: {model_path}")
            cached = Model(model_path)
            _MODEL_CACHE[model_path] = cached
        return cached

    def accept_audio(self, pcm_bytes):
        if not pcm_bytes:
            return False, ""

        if self.recognizer.AcceptWaveform(pcm_bytes):
            parsed = json.loads(self.recognizer.Result() or "{}")
            text = self._to_digits(parsed.get("text", ""))
            return True, text
        return False, ""

    def finalize(self):
        parsed = json.loads(self.recognizer.FinalResult() or "{}")
        text = self._to_digits(parsed.get("text", ""))
        print(f"[Dictate] FinalResult text='{text}'")
        return text

    def reset(self):
        self.recognizer = KaldiRecognizer(self.model, self.sample_rate, DIGIT_GRAMMAR)

    @staticmethod
    def _to_digits(text):
        if not text:
            return ""

        pieces = text.strip().split()
        mapped = []
        word_map = {
            "zero": "0",
            "oh": "0",
            "one": "1",
            "two": "2",
            "to": "2",
            "too": "2",
            "three": "3",
            "four": "4",
            "for": "4",
            "five": "5",
            "six": "6",
            "seven": "7",
            "eight": "8",
            "ate": "8",
            "nine": "9",
        }
        for piece in pieces:
            if piece.isdigit() and len(piece) == 1:
                mapped.append(piece)
            else:
                mapped.append(word_map.get(piece.lower(), piece))

        return DIGIT_ONLY_PATTERN.sub("", "".join(mapped))
