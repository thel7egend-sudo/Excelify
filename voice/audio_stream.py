from collections import deque
import threading

import numpy as np
import sounddevice as sd


class AudioStream:
    def __init__(
        self,
        sample_rate=16000,
        channels=1,
        dtype="float32",
        blocksize=4000,
        level_callback=None,
        error_callback=None,
        device=None,
    ):
        self.sample_rate = sample_rate
        self.channels = channels
        self.dtype = dtype
        self.blocksize = blocksize
        self.level_callback = level_callback
        self.error_callback = error_callback
        self.device = device

        self._queue = deque()
        self._lock = threading.Lock()
        self._stream = None
        self._chunks_received = 0

    def start(self):
        if self._stream is not None:
            return

        input_device = self.device
        if input_device is None:
            default_device = sd.default.device
            if isinstance(default_device, (tuple, list)) and default_device:
                input_device = default_device[0]

        print(f"[Dictate] Starting InputStream device={input_device} sr={self.sample_rate}")
        self._stream = sd.InputStream(
            samplerate=self.sample_rate,
            channels=self.channels,
            dtype=self.dtype,
            blocksize=self.blocksize,
            device=input_device,
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
        if chunks:
            print(f"[Dictate] Consuming {len(chunks)} queued chunks")
        return chunks

    def _on_audio(self, indata, frames, time_info, status):
        if status:
            msg = f"Audio stream status: {status}"
            print(f"[Dictate] {msg}")
            if self.error_callback is not None:
                self.error_callback(msg)

        if indata is None or frames <= 0:
            return

        self._chunks_received += 1
        if self._chunks_received <= 5 or self._chunks_received % 50 == 0:
            print(f"[Dictate] chunk#{self._chunks_received} shape={indata.shape} dtype={indata.dtype}")

        clipped = np.clip(indata, -1.0, 1.0)
        pcm16 = (clipped * 32767.0).astype(np.int16)
        chunk = pcm16.tobytes()

        with self._lock:
            self._queue.append(chunk)

        if self.level_callback is not None:
            mono = clipped[:, 0] if clipped.ndim > 1 else clipped
            rms = float(np.sqrt(np.mean(np.square(mono))))
            self.level_callback(rms)
