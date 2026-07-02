"""
watcher.py — Continuous background service for the County Gap Tracker.

Watches the live spreadsheet for changes and automatically rebuilds
data.json and pushes it to GitHub Pages — with debouncing and a cooldown
so we never hammer git or the dashboard's hosting, even with several
analysts saving the file in quick succession.

Now includes a QUIET WINDOW:
➡️ Idles daily from 5 PM → 6 AM (no processing or pushes)

Run this once and leave it running. Stop it with Ctrl+C.
"""
import sys
sys.pycache_prefix = r"C:\Documents\PythonCache"

import time, threading
from pathlib import Path
from datetime import datetime, time as dtime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import gap_dashboard as gd

# =============================================================================
# TIMING — tuned to be safe with multiple concurrent editors.
# =============================================================================
DEBOUNCE_SECONDS      = 120    # wait after last change before reading
MIN_PUSH_INTERVAL     = 900    # minimum seconds between pushes
SAFETY_RESYNC_SECONDS = 1800   # force resync after this time

# =============================================================================
# QUIET HOURS — system will idle during this window
# =============================================================================
QUIET_START = dtime(17, 0)  # 5:00 PM
QUIET_END   = dtime(6, 0)   # 6:00 AM

def in_quiet_hours():
    now = datetime.now().time()
    # Handles overnight window (crosses midnight)
    return now >= QUIET_START or now <= QUIET_END

# =============================================================================

HEARTBEAT_FILE = gd.WORK_DIR / "GapDashboardSite" / "GapdashBoard" / "watcher_heartbeat.txt"
TARGET_FILE = gd.LIVE_SOURCE.name
WATCH_DIR   = gd.LIVE_SOURCE.parent

_lock = threading.Lock()
_last_event_time = 0.0
_last_push_time  = 0.0
_pending = False

def ts():
    return datetime.now().strftime("%H:%M:%S")

def write_heartbeat(status):
    """Writes a small status file so you can confirm watcher is alive."""
    try:
        gd.WORK_DIR.mkdir(parents=True, exist_ok=True)
        HEARTBEAT_FILE.write_text(
            f"Last check-in: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Status: {status}\n"
        )
    except Exception:
        pass

# =============================================================================

class Handler(FileSystemEventHandler):
    def on_modified(self, event): self._touch(event)
    def on_created(self, event):  self._touch(event)
    def on_moved(self, event):    self._touch(event)

    def _touch(self, event):
        name = Path(getattr(event, "dest_path", "") or event.src_path).name
        if name != TARGET_FILE:
            return
        global _last_event_time, _pending
        with _lock:
            _last_event_time = time.time()
            _pending = True
        print(f"[{ts()}] Change detected on {TARGET_FILE}")

# =============================================================================

def worker_loop():
    global _pending, _last_push_time

    while True:
        time.sleep(2)

        # ------------------------------------------------------------
        # QUIET HOURS CHECK
        # ------------------------------------------------------------
        if in_quiet_hours():
            write_heartbeat("idle (quiet hours)")
            continue
        else:
            write_heartbeat("running")

        with _lock:
            quiet_for  = time.time() - _last_event_time if _pending else None
            since_push = time.time() - _last_push_time

        # Safety resync (disabled during quiet hours)
        force_resync = (since_push > SAFETY_RESYNC_SECONDS) and not in_quiet_hours()

        if not _pending and not force_resync:
            continue

        if _pending and quiet_for is not None and quiet_for < DEBOUNCE_SECONDS:
            continue  # still being edited

        if since_push < MIN_PUSH_INTERVAL:
            continue

        with _lock:
            _pending = False

        reason = "safety resync" if force_resync else "change detected"
        print(f"[{ts()}] Running pipeline... ({reason})")

        try:
            ok = gd.run_pipeline(push=True)
            print(f"[{ts()}] {'Done.' if ok else 'Pipeline reported an issue.'}")
            write_heartbeat("last run OK" if ok else "last run reported an issue")
        except Exception as e:
            print(f"[{ts()}] [ERROR] Pipeline crashed: {e}")
            write_heartbeat(f"last run crashed: {e}")

        with _lock:
            _last_push_time = time.time()

# =============================================================================

def main():
    global _last_push_time

    print("="*64)
    print("  GAP DASHBOARD WATCHER")
    print(f"  Watching      : {gd.LIVE_SOURCE}")
    print(f"  Debounce      : {DEBOUNCE_SECONDS}s")
    print(f"  Min push gap  : {MIN_PUSH_INTERVAL}s")
    print(f"  Safety resync : every {SAFETY_RESYNC_SECONDS//60} min")
    print(f"  Quiet hours   : 5:00 PM → 6:00 AM")
    print(f"  Heartbeat     : {HEARTBEAT_FILE}")
    print("="*64)

    if not WATCH_DIR.exists():
        print(f"\n[ERROR] Watch folder does not exist:\n  {WATCH_DIR}")
        write_heartbeat(f"ERROR: watch folder not found ({WATCH_DIR})")
        return

    print("\nRunning initial sync...")

    try:
        gd.run_pipeline(push=True)
        write_heartbeat("initial sync OK")
    except Exception as e:
        print(f"[WARN] Initial sync failed: {e}")
        write_heartbeat(f"initial sync failed: {e}")

    _last_push_time = time.time()

    observer = Observer()
    observer.schedule(Handler(), str(WATCH_DIR), recursive=False)
    observer.start()

    print(f"\n[{ts()}] Watching for changes. Press Ctrl+C to stop.\n")

    t = threading.Thread(target=worker_loop, daemon=True)
    t.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping watcher...")
        observer.stop()

    observer.join()

# =============================================================================

if __name__ == "__main__":
    main()