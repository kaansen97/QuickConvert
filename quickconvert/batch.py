"""Collect a multi-file selection into one process.

Explorer starts one process per selected file. Each process drops its file into a spool folder;
the first one to grab a named semaphore becomes the "worker" and converts everything in the folder,
the others exit straight away.
"""

import ctypes
import os
import tempfile
import time
from pathlib import Path

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
kernel32.CreateSemaphoreW.restype = ctypes.c_void_p
kernel32.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
kernel32.ReleaseSemaphore.argtypes = [ctypes.c_void_p, ctypes.c_long, ctypes.c_void_p]

STALE_SECONDS = 120  # spool entries left behind by a crashed run are ignored
GATHER_SECONDS = 0.6  # Explorer starts the per-file processes over a few hundred ms


class Batch:
    def __init__(self, target: str):
        self.spool = Path(tempfile.gettempdir()) / "QuickConvert" / f"spool-{target}"
        self.spool.mkdir(parents=True, exist_ok=True)
        # A semaphore (not a mutex) so the worker thread may release what the main thread took.
        self.lock = kernel32.CreateSemaphoreW(None, 1, 1, f"Local\\QuickConvert-{target}")
        self.done = 0

    def submit(self, files: list[Path]) -> bool:
        """Queue the files; True if this process should do the converting."""
        for f in files:
            name = self.spool / f"{time.time_ns()}-{os.getpid()}-{f.name}"
            name.with_suffix(".tmp").write_text(str(f.resolve()), encoding="utf-8")
            name.with_suffix(".tmp").replace(name.with_suffix(".job"))
        return self._acquire()

    def _acquire(self) -> bool:
        return kernel32.WaitForSingleObject(self.lock, 0) == 0

    def _release(self) -> None:
        kernel32.ReleaseSemaphore(self.lock, 1, None)

    def _jobs(self) -> list[Path]:
        jobs = []
        for job in sorted(self.spool.glob("*.job")):
            if time.time() - job.stat().st_mtime > STALE_SECONDS:
                job.unlink(missing_ok=True)
            else:
                jobs.append(job)
        return jobs

    def pending(self) -> int:
        return len(self._jobs())

    def wait_for_selection(self) -> None:
        time.sleep(GATHER_SECONDS)

    def next_file(self) -> Path | None:
        """Next queued file, or None once the queue is empty (the lock is then given up)."""
        while True:
            for _ in range(3):  # late arrivals from the same selection
                jobs = self._jobs()
                if jobs:
                    path = Path(jobs[0].read_text(encoding="utf-8"))
                    jobs[0].unlink(missing_ok=True)
                    self.done += 1
                    return path
                time.sleep(0.1)
            self._release()
            # A file queued between the check and the release would be stranded; look once more.
            if not self._jobs() or not self._acquire():
                return None
