"""
ProPresenter Converter — Menu Bar Agent
Runs as a macOS status bar icon (no dock icon).
Provides quick access to conversion without opening the full window.
"""

import os
import sys
import subprocess
import threading

# When frozen by PyInstaller, resources live in sys._MEIPASS.
# When running from source, fall back to the script directory / bundled .app.
if getattr(sys, 'frozen', False):
    # Running as compiled PyInstaller app
    _here      = os.path.dirname(sys.executable)
    _resources = sys._MEIPASS
    # Add proto_pb2 and project root to path
    sys.path.insert(0, os.path.join(_resources, 'proto_pb2'))
    sys.path.insert(0, _resources)
else:
    # Running from source — set up venv if needed
    _here = os.path.dirname(os.path.abspath(__file__))
    _site = os.path.join(_here, "venv", "lib",
                         f"python{sys.version_info.major}.{sys.version_info.minor}",
                         "site-packages")
    if os.path.isdir(_site):
        sys.path.insert(0, _site)
    sys.path.insert(0, _here)
    _resources = _here if os.path.exists(os.path.join(_here, "menubar_idle.png")) \
                 else os.path.join(_here, "ProPresenter Converter.app", "Contents", "Resources")

import rumps

# Import the converter core (no GUI)
from pro_to_pptx import (
    FolderWatcher, convert_file, _history,
    DEFAULT_WATCH_FOLDER, WATCHDOG_AVAILABLE
)

ICON_IDLE   = os.path.join(_resources, "menubar_idle.png")    # greyed + slash = not watching
ICON_ACTIVE = os.path.join(_resources, "menubar_active.png")  # full colour = watching


class ConverterMenuBarApp(rumps.App):

    def __init__(self):
        super().__init__(
            name="ProPresenter Converter",
            icon=ICON_IDLE if os.path.exists(ICON_IDLE) else None,
            quit_button=None   # we'll add our own at the bottom
        )
        self.title = None   # no text next to icon

        self._watcher: FolderWatcher = None
        self._watch_dir  = DEFAULT_WATCH_FOLDER
        self._output_dir = ""

        self._build_menu()

    # ------------------------------------------------------------------
    # Menu construction
    # ------------------------------------------------------------------

    def _build_menu(self):
        self.menu = [
            rumps.MenuItem("Open Converter Window", callback=self.open_window),
            None,  # separator
            rumps.MenuItem("Watch Folder", callback=self.toggle_watch),
            rumps.MenuItem("Convert Existing Files", callback=self.convert_existing),
            None,
            rumps.MenuItem("Set Watch Folder…", callback=self.set_watch_folder),
            rumps.MenuItem("Set Output Folder…", callback=self.set_output_folder),
            None,
            rumps.MenuItem("Quit", callback=self.quit_app),
        ]
        self._update_watch_labels()

    def _update_watch_labels(self):
        watch_label = os.path.basename(self._watch_dir) or self._watch_dir
        out_label   = os.path.basename(self._output_dir) or "Not set"
        self.menu["Set Watch Folder…"].title  = f"Watch Folder: {watch_label}"
        self.menu["Set Output Folder…"].title = f"Output Folder: {out_label}"

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def open_window(self, _):
        """Launch the full tkinter GUI as a separate process."""
        if getattr(sys, 'frozen', False):
            # Re-launch this same compiled binary with --window flag
            subprocess.Popen([sys.executable, "--window"])
        else:
            # Running from source — launch with venv Python
            python = os.path.join(_here, "venv", "bin", "python3.13")
            script = os.path.join(_here, "pro_to_pptx.py")
            if os.path.exists(python):
                subprocess.Popen([python, script])

    def toggle_watch(self, sender):
        if self._watcher and self._watcher.running:
            self._stop_watch()
        else:
            self._start_watch()

    def _start_watch(self):
        if not WATCHDOG_AVAILABLE:
            rumps.alert("Missing Dependency",
                        "watchdog is not installed.\nRun: pip install watchdog")
            return
        if not self._output_dir:
            rumps.alert("No Output Folder",
                        "Please set an output folder first\n(Output Folder menu item).")
            return
        if not os.path.isdir(self._watch_dir):
            rumps.alert("Watch Folder Not Found",
                        f"Folder does not exist:\n{self._watch_dir}")
            return

        os.makedirs(self._output_dir, exist_ok=True)

        self._watcher = FolderWatcher(
            self._watch_dir,
            self._output_dir,
            callback=self._on_watch_event
        )
        self._watcher.start()
        self.menu["Watch Folder"].title = "⏹ Stop Watching"
        if os.path.exists(ICON_ACTIVE):
            self.icon = ICON_ACTIVE   # swap to full-colour icon

    def _stop_watch(self):
        if self._watcher:
            self._watcher.stop()
            self._watcher = None
        self.menu["Watch Folder"].title = "▶ Start Watching"
        if os.path.exists(ICON_IDLE):
            self.icon = ICON_IDLE   # swap back to greyed icon

    def _on_watch_event(self, pro_path, result):
        name = os.path.basename(pro_path)
        # Show a macOS notification
        rumps.notification(
            title="ProPresenter Converter",
            subtitle=name,
            message=result,
            sound=False
        )

    def convert_existing(self, _):
        if not self._output_dir:
            rumps.alert("No Output Folder",
                        "Please set an output folder first.")
            return
        if not os.path.isdir(self._watch_dir):
            rumps.alert("Watch Folder Not Found",
                        f"Folder does not exist:\n{self._watch_dir}")
            return

        # Collect unconverted files
        pro_files = []
        for dirpath, _, filenames in os.walk(self._watch_dir):
            for fn in filenames:
                if fn.endswith(".pro"):
                    pro_files.append(os.path.join(dirpath, fn))

        new_files = [p for p in pro_files if not _history.is_converted(p)]
        skipped   = len(pro_files) - len(new_files)

        if not new_files:
            rumps.alert("Nothing to Convert",
                        f"All {len(pro_files)} file(s) already converted.\n"
                        f"Use the full window to clear history.")
            return

        msg = f"{len(new_files)} file(s) to convert"
        if skipped:
            msg += f" ({skipped} already done, skipping)"

        response = rumps.alert(
            title="Convert Existing Files",
            message=msg + f"\nWatch: {self._watch_dir}\nOutput: {self._output_dir}",
            ok="Convert", cancel="Cancel"
        )
        if response.clicked != 1:
            return

        threading.Thread(
            target=self._run_batch,
            args=(new_files,),
            daemon=True
        ).start()

    def _run_batch(self, files):
        done, errors = 0, 0
        for path in files:
            try:
                convert_file(path, self._output_dir)
                done += 1
            except Exception:
                errors += 1

        msg = f"{done} file(s) converted"
        if errors:
            msg += f", {errors} failed"
        rumps.notification(
            title="ProPresenter Converter",
            subtitle="Batch Complete",
            message=msg,
            sound=False
        )

    def _pick_folder(self, title: str, default: str) -> str:
        """Show a native macOS folder picker via osascript. Returns path or ''."""
        default = default or os.path.expanduser("~")
        script = (
            f'tell application "Finder"\n'
            f'  activate\n'
            f'end tell\n'
            f'set chosen to choose folder with prompt "{title}" '
            f'default location (POSIX file "{default}")\n'
            f'POSIX path of chosen'
        )
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=60
            )
            path = result.stdout.strip()
            # osascript adds a trailing slash — remove it
            return path.rstrip("/") if path else ""
        except Exception:
            return ""

    def set_watch_folder(self, _):
        path = self._pick_folder(
            "Select your ProPresenter library folder to watch",
            self._watch_dir
        )
        if path:
            self._watch_dir = path
            self._update_watch_labels()

    def set_output_folder(self, _):
        path = self._pick_folder(
            "Select folder to save converted PPTX files",
            self._output_dir or os.path.expanduser("~/Desktop")
        )
        if path:
            self._output_dir = path
            self._update_watch_labels()

    def quit_app(self, _):
        if self._watcher:
            self._watcher.stop()
        rumps.quit_application()


def _acquire_lock():
    """
    Create a lock file so only one menu bar instance runs at a time.
    Returns the lock file handle, or None if another instance is already running.
    """
    import fcntl
    lock_path = os.path.join(
        os.path.expanduser("~/Library/Application Support/ProPresenter Converter"),
        "menubar.lock"
    )
    os.makedirs(os.path.dirname(lock_path), exist_ok=True)
    try:
        fh = open(lock_path, "w")
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        return fh
    except OSError:
        return None  # already locked — another instance is running


if __name__ == "__main__":
    if "--window" in sys.argv:
        # Launch the full tkinter GUI in-process (separate process already spawned)
        sys.path.insert(0, _resources)
        import runpy
        runpy.run_path(os.path.join(_resources, "pro_to_pptx.py"), run_name="__main__")
    else:
        lock = _acquire_lock()
        if lock is None:
            # Another menu bar instance is already running — just bring attention to it
            subprocess.run([
                "osascript", "-e",
                'display notification "ProPresenter Converter is already running in the menu bar." '
                'with title "ProPresenter Converter"'
            ])
            sys.exit(0)
        ConverterMenuBarApp().run()
