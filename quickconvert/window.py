"""Small progress window shown while the right-click conversion runs."""

import os
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk

from .common import set_progress_hook

AUTO_CLOSE_MS = 5000
GREEN, RED = "#1a7f37", "#cf222e"


class ProgressWindow:
    def __init__(self, batch, target: str, convert):
        self.batch, self.target, self.convert = batch, target, convert
        self.events = queue.Queue()  # worker thread -> UI thread
        self.results: list[Path] = []
        self.failures: list[str] = []

        self.root = tk.Tk()
        self.root.title("QuickConvert")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)

        self.title = ttk.Label(frame, font=("Segoe UI Semibold", 11), width=52)
        self.title.pack(anchor="w")
        self.step = ttk.Label(frame, foreground="#59636e", text="Collecting files…")
        self.step.pack(anchor="w", pady=(2, 8))
        self.bar = ttk.Progressbar(frame, length=420, maximum=1.0)
        self.bar.pack(fill="x")
        self.detail = ttk.Label(frame, wraplength=420, justify="left")
        self.detail.pack(anchor="w", pady=(8, 0))
        self.buttons = ttk.Frame(frame)
        self.buttons.pack(anchor="e", pady=(10, 0))

        self.title.config(text=f"Converting to {self.target.upper()}")
        self._center()

    def _center(self) -> None:
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 3
        self.root.geometry(f"+{x}+{y}")

    # ---- worker thread -------------------------------------------------------
    def _work(self) -> None:
        self.batch.wait_for_selection()
        while (source := self.batch.next_file()) is not None:
            index, total = self.batch.done, self.batch.done + self.batch.pending()
            self.events.put(("file", source.name, index, total))
            set_progress_hook(
                lambda step, f, i=index, n=total: self.events.put(("step", step, (i - 1 + f) / n))
            )
            try:
                self.results.append(self.convert(source, self.target))
            except Exception as exc:
                self.failures.append(f"{source.name}: {exc}")
        set_progress_hook(None)
        self.events.put(("done",))

    # ---- UI thread -----------------------------------------------------------
    def _poll(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "file":
                    _, name, index, total = event
                    count = f"  ({index} of {total})" if total > 1 else ""
                    self.title.config(text=f"{name} → {self.target.upper()}{count}")
                    self.bar.config(value=(index - 1) / total)
                elif event[0] == "step":
                    self.step.config(text=event[1])
                    self.bar.config(value=event[2])
                elif event[0] == "done":
                    self._finish()
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _finish(self) -> None:
        self.bar.config(value=1.0)
        converted = len(self.results)
        if converted > 1:
            self.title.config(text=f"{converted} files → {self.target.upper()}")
        if self.failures:
            self.step.config(text=f"Could not convert {len(self.failures)} file(s)", foreground=RED)
            self.detail.config(text="\n".join(self.failures), foreground=RED)
        else:
            self.step.config(text="✓ Done", foreground=GREEN)
            names = ", ".join(p.name for p in self.results[:3])
            more = f" and {converted - 3} more" if converted > 3 else ""
            self.detail.config(text=f"Saved as {names}{more}")
            self.root.after(AUTO_CLOSE_MS, self.root.destroy)
        if converted == 1:
            ttk.Button(self.buttons, text="Open", command=self._open).pack(side="left", padx=4)
        if converted:
            ttk.Button(self.buttons, text="Show in folder", command=self._reveal).pack(side="left")
        ttk.Button(self.buttons, text="Close", command=self.root.destroy).pack(side="left", padx=4)

    def _open(self) -> None:
        os.startfile(self.results[-1])
        self.root.destroy()

    def _reveal(self) -> None:
        subprocess.Popen(["explorer", "/select,", str(self.results[-1])])
        self.root.destroy()

    def run(self) -> int:
        threading.Thread(target=self._work, daemon=True).start()
        self.root.after(100, self._poll)
        self.root.mainloop()
        return 1 if self.failures else 0
