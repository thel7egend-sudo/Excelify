from dataclasses import dataclass
from typing import List, Optional

import numpy as np
import sounddevice as sd


@dataclass(frozen=True)
class InputDevice:
    device_id: int
    name: str
    max_input_channels: int
    default_samplerate: float


def list_input_devices() -> List[InputDevice]:
    devices = sd.query_devices()
    inputs: List[InputDevice] = []
    for device_id, info in enumerate(devices):
        if info.get("max_input_channels", 0) <= 0:
            continue
        inputs.append(
            InputDevice(
                device_id=device_id,
                name=info.get("name", f"Device {device_id}"),
                max_input_channels=info.get("max_input_channels", 0),
                default_samplerate=info.get("default_samplerate", 16000.0),
            )
        )
    return inputs


def default_input_device_id() -> Optional[int]:
    default_devices = sd.default.device
    if not default_devices:
        return None
    device_id = default_devices[0]
    if device_id is None:
        return None
    return int(device_id)


class AudioRecorder:
    def __init__(self, samplerate: int = 16000, channels: int = 1, dtype: str = "float32"):
        self.samplerate = samplerate
        self.channels = channels
        self.dtype = dtype
        self._frames: List[np.ndarray] = []
        self._stream: Optional[sd.InputStream] = None
        self._recording = False

    @property
    def is_recording(self) -> bool:
        return self._recording

    def start(self, device_id: Optional[int] = None) -> None:
        if self._recording:
            return
        self._frames = []
        self._stream = sd.InputStream(
            samplerate=self.samplerate,
            channels=self.channels,
            dtype=self.dtype,
            device=device_id,
            callback=self._callback,
        )
        self._stream.start()
        self._recording = True

    def stop(self) -> np.ndarray:
        if not self._recording:
            return np.array([], dtype=self.dtype)
        try:
            if self._stream is not None:
                self._stream.stop()
                self._stream.close()
        finally:
            self._stream = None
            self._recording = False

        if not self._frames:
            return np.array([], dtype=self.dtype)
        return np.concatenate(self._frames, axis=0)

    def _callback(self, indata, frames, time, status):
        if status:
            pass
        self._frames.append(indata.copy())
