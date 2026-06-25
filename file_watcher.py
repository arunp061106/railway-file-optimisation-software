"""
file_watcher.py
---------------
Monitors a folder for new files using watchdog.
Detects when a file has fully finished downloading (size stabilises)
before triggering the categorisation pipeline.
Fully offline.
"""

import os
import time
import queue
import threading
import logging
from pathlib import Path

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_AVAILABLE = True
except ImportError:
    WATCHDOG_AVAILABLE = False

logger = logging.getLogger("FileWatcher")

# File extensions to ignore (browser temp files during download)
IGNORE_EXTENSIONS = {".crdownload", ".part", ".tmp", ".partial",
                     ".download", ".opdownload", ".!ut"}

# Minimum file age (seconds) before processing — avoids picking up mid-write files
MIN_AGE_SECONDS = 2

# How long to wait for file size to stabilise
STABLE_CHECK_INTERVAL = 1.5   # seconds between size checks
STABLE_REQUIRED_COUNT = 2     # number of identical size readings needed


class _DownloadHandler(FileSystemEventHandler):
    """Internal watchdog event handler. Puts new file paths into a queue."""

    def __init__(self, file_queue: queue.Queue, ignore_exts: set):
        self.file_queue = file_queue
        self.ignore_exts = ignore_exts
        self._seen = set()

    def on_created(self, event):
        if not event.is_directory:
            self._handle_path(event.src_path)

    def on_moved(self, event):
        # Handles files moved/renamed into the watched folder (some browsers do this)
        if not event.is_directory:
            self._handle_path(event.dest_path)

    def _handle_path(self, path: str):
        suffix = Path(path).suffix.lower()
        if suffix in self.ignore_exts:
            return
        if path not in self._seen:
            self._seen.add(path)
            self.file_queue.put(path)
            logger.debug(f"Queued: {path}")


class FileWatcher:
    """
    Watches a folder for new files and calls a callback when a file is ready.

    Usage:
        def on_file(path): ...
        watcher = FileWatcher("C:/Users/.../Downloads", on_file_ready=on_file)
        watcher.start()
        ...
        watcher.stop()
    """

    def __init__(self, watch_folder: str, on_file_ready,
                 ignore_extensions: set = None,
                 stable_wait: float = STABLE_CHECK_INTERVAL,
                 stable_count: int = STABLE_REQUIRED_COUNT):
        """
        Args:
            watch_folder: Folder path to monitor.
            on_file_ready: Callback function(filepath: str) called when file is ready.
            ignore_extensions: Set of extensions to skip (e.g. {'.crdownload'}).
            stable_wait: Seconds between file-size checks.
            stable_count: How many equal-size readings needed to confirm done.
        """
        if not WATCHDOG_AVAILABLE:
            raise ImportError("watchdog library not installed. Run: pip install watchdog")

        self.watch_folder = str(watch_folder)
        self.on_file_ready = on_file_ready
        self.ignore_extensions = ignore_extensions or IGNORE_EXTENSIONS
        self.stable_wait = stable_wait
        self.stable_count = stable_count

        self._file_queue = queue.Queue()
        self._observer = None
        self._processor_thread = None
        self._running = False

    def start(self):
        """Start watching the folder."""
        if not os.path.isdir(self.watch_folder):
            raise FileNotFoundError(f"Watch folder not found: {self.watch_folder}")

        self._running = True

        # Start watchdog observer
        handler = _DownloadHandler(self._file_queue, self.ignore_extensions)
        self._observer = Observer()
        self._observer.schedule(handler, self.watch_folder, recursive=False)
        self._observer.start()
        logger.info(f"Watching: {self.watch_folder}")

        # Start processor thread
        self._processor_thread = threading.Thread(
            target=self._process_queue, daemon=True, name="FileProcessor")
        self._processor_thread.start()

    def stop(self):
        """Stop watching."""
        self._running = False
        if self._observer:
            self._observer.stop()
            self._observer.join(timeout=3)
        logger.info("FileWatcher stopped.")

    def _is_file_stable(self, path: str) -> bool:
        """
        Return True when a file's size hasn't changed for stable_count checks.
        This confirms the download/write has completed.
        """
        if not os.path.exists(path):
            return False
        prev_size = -1
        stable_readings = 0
        while stable_readings < self.stable_count:
            try:
                current_size = os.path.getsize(path)
            except OSError:
                return False
            if current_size == prev_size and current_size > 0:
                stable_readings += 1
            else:
                stable_readings = 0
            prev_size = current_size
            time.sleep(self.stable_wait)
        return True

    def _process_queue(self):
        """Background thread: waits for files to stabilise then fires callback."""
        while self._running:
            try:
                filepath = self._file_queue.get(timeout=1)
            except queue.Empty:
                continue

            # Brief initial delay
            time.sleep(MIN_AGE_SECONDS)

            # Skip temp extensions that appeared while waiting
            suffix = Path(filepath).suffix.lower()
            if suffix in self.ignore_extensions:
                continue

            # Wait for file to finish writing
            if not self._is_file_stable(filepath):
                logger.warning(f"File disappeared or couldn't stabilise: {filepath}")
                continue

            if not os.path.exists(filepath):
                continue

            logger.info(f"File ready: {filepath}")
            try:
                self.on_file_ready(filepath)
            except Exception as e:
                logger.error(f"Error in on_file_ready callback: {e}")


# ---------- Quick test ----------
if __name__ == "__main__":
    import tempfile, time, os

    logging.basicConfig(level=logging.DEBUG)

    with tempfile.TemporaryDirectory() as tmpdir:
        received = []

        def got_file(path):
            print(f"\n✅ File ready: {path}")
            received.append(path)

        watcher = FileWatcher(tmpdir, on_file_ready=got_file,
                              stable_wait=0.5, stable_count=2)
        watcher.start()

        # Simulate a file appearing
        time.sleep(1)
        test_file = os.path.join(tmpdir, "TDR_2024_Test.pdf")
        with open(test_file, "w") as f:
            f.write("Railway tender document content")

        time.sleep(5)
        watcher.stop()
        print(f"\nTotal files received: {len(received)}")
