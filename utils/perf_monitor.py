"""Performance monitor: logs operation timing and detects UI freezes."""
import os
import sys
import time
import threading
import traceback
from datetime import datetime

_LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
_LOG_PATH = os.path.join(_LOG_DIR, "performance.log")

_lock = threading.Lock()
_watchdog_last_ping = time.time()
_watchdog_running = False
_watchdog_frozen = False

_MAX_LOG_BYTES = 20 * 1024 * 1024  # 20 MB 轮转上限，防止日志失控
_LOG_BACKUP = _LOG_PATH + ".old"


def _ensure_dir():
    os.makedirs(_LOG_DIR, exist_ok=True)


def _rotate_if_needed():
    """Single-file rotation: roll the log to .old once it exceeds _MAX_LOG_BYTES."""
    try:
        if os.path.exists(_LOG_PATH) and os.path.getsize(_LOG_PATH) > _MAX_LOG_BYTES:
            if os.path.exists(_LOG_BACKUP):
                os.remove(_LOG_BACKUP)
            os.replace(_LOG_PATH, _LOG_BACKUP)
    except OSError:
        pass


def log(msg):
    """Write a timestamped message to the performance log."""
    _ensure_dir()
    ts = datetime.now().strftime("%m-%d %H:%M:%S")
    with _lock:
        _rotate_if_needed()
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")


class Tracker:
    """Context manager to track an operation. Logs start, end, and duration.
    Also reports if the operation exceeded a threshold."""

    def __init__(self, name, threshold_ms=500):
        self.name = name
        self.threshold_ms = threshold_ms
        self.t0 = 0

    def __enter__(self):
        self.t0 = time.time()
        log(f"BEGIN [{self.name}]")
        return self

    def __exit__(self, *args):
        elapsed = (time.time() - self.t0) * 1000
        if elapsed > self.threshold_ms:
            log(f"END   [{self.name}] {elapsed:.0f}ms *** SLOW (> {self.threshold_ms}ms)")
        else:
            log(f"END   [{self.name}] {elapsed:.0f}ms")


def track(func):
    """Decorator: wraps a function with performance tracking."""
    def wrapper(*args, **kwargs):
        with Tracker(func.__qualname__):
            return func(*args, **kwargs)
    return wrapper


def start_watchdog(root_widget, threshold_sec=3.0, check_interval=1.0):
    """Start a background watchdog that detects main-thread freezes.
    Pings the main thread via root.after(). If no pong within threshold_sec,
    logs a freeze warning with the current stack traces."""

    global _watchdog_running, _watchdog_last_ping, _watchdog_frozen
    if _watchdog_running:
        return
    _watchdog_running = True
    _watchdog_last_ping = time.time()
    _watchdog_frozen = False

    def _pong():
        global _watchdog_last_ping
        _watchdog_last_ping = time.time()

    def _ping():
        try:
            root_widget.after(0, _pong)
        except Exception:
            pass

    def _watchdog_loop():
        global _watchdog_frozen
        while _watchdog_running:
            time.sleep(check_interval)
            _ping()
            time.sleep(0.05)  # let after() deliver the pong
            gap = time.time() - _watchdog_last_ping
            if gap > threshold_sec:
                if not _watchdog_frozen:
                    _watchdog_frozen = True
                    log(f"*** FREEZE DETECTED: main thread unresponsive for {gap:.1f}s ***")
                    stacks = []
                    for tid, frame in sys._current_frames().items():
                        lines = traceback.format_stack(frame)
                        if len(lines) > 40:
                            lines = ["  ... (earlier frames omitted)\n"] + lines[-40:]
                        stacks.append(f"  Thread {tid}:\n" + "".join(lines))
                    log("Stack dump:\n" + "\n".join(stacks))
            else:
                _watchdog_frozen = False

    t = threading.Thread(target=_watchdog_loop, daemon=True)
    t.start()
    log(f"Watchdog started (threshold={threshold_sec}s)")


def stop_watchdog():
    global _watchdog_running
    _watchdog_running = False
    log("Watchdog stopped")
