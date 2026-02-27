from collections import deque
import threading

import numpy as np
import sounddevice as sd


class AudioStream:
    def __init__(self, sample_rate=16000, channels=1, dtype="float32", blocksize=8000, level_callback=None):
        self.sample_rate = sample_rate
        self.channels = channels
        self.dtype = dtype
        self.blocksize = blocksize
        self.level_callback = level_callback

        self._queue = deque()
        self._lock = threading.Lock()
        self._stream = None

    def start(self):
        if self._stream is not None:
            return

        self._stream = sd.InputStream(
            samplerate=self.sample_rate,
            channels=self.channels,
            dtype=self.dtype,
            blocksize=self.blocksize,
            callback=self._on_audio,
        )
        self._stream.start()

    def stop(self):
        if self._stream is None:
            return
        self._stream.stop()
        self._stream.close()
        self._stream = None

    def clear(self):
        with self._lock:
            self._queue.clear()

    def pop_all_chunks(self):
        with self._lock:
            chunks = list(self._queue)
            self._queue.clear()
        return chunks

    def _on_audio(self, indata, frames, time_info, status):
        if status:
            return

        pcm16 = np.clip(indata, -1.0, 1.0)
        pcm16 = (pcm16 * 32767).astype(np.int16)
        chunk = pcm16.tobytes()

        with self._lock:
            self._queue.append(chunk)

        if self.level_callback is not None:
            rms = float(np.sqrt(np.mean(np.square(indata), axis=0)).mean())
            self.level_callback(rms)
