"""
vba_modules -> Excel auto sync (단방향)
- vba_modules(.bas/.frm) 변경 시 Excel에 Import
- Excel 저장 시 export는 실행하지 않음 (외부에서 Excel 직접 수정 가능)
Usage: python watch_and_export.py [workbook_path]
"""
import os
import sys
import time
from datetime import datetime
import subprocess
import threading

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    print("ERROR: Install watchdog: pip install watchdog")
    sys.exit(1)


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SYNC_SCRIPT = os.path.join(SCRIPT_DIR, "sync_to_excel.py")
VBA_MODULES_DIR = os.path.join(SCRIPT_DIR, "vba_modules")
DEBOUNCE_BAS_SEC = 1.0  # 실시간에 가깝게 (동시 sync 방지용 짧은 디바운스)


class VbaModulesChangeHandler(FileSystemEventHandler):
    def __init__(self, workbook_path):
        self.workbook_path = os.path.abspath(workbook_path)
        self.timer = None
        self.lock = threading.Lock()

    def _run_sync(self):
        lock_file = os.path.join(SCRIPT_DIR, ".sync_lock")
        fail_file = os.path.join(SCRIPT_DIR, ".sync_last_fail")
        if os.path.isfile(lock_file):
            try:
                if time.time() - os.path.getmtime(lock_file) < 60:
                    print("[Auto-Sync] sync 진행 중 - 건너뜀")
                    return
            except Exception:
                pass
        if os.path.isfile(fail_file):
            try:
                with open(fail_file, "r") as f:
                    t = float(f.read().strip())
                if time.time() - t < 60:
                    print("[Auto-Sync] 이전 실패 - 1분 cooldown")
                    return
            except Exception:
                pass
        cmd = [sys.executable, SYNC_SCRIPT, self.workbook_path]
        print("\n[Auto-Sync] vba_modules changed -> sync_to_excel running...")
        try:
            result = subprocess.run(cmd, cwd=SCRIPT_DIR, capture_output=True, text=True, timeout=300)
            print(result.stdout)
            if result.stderr:
                print(result.stderr)
            done_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if result.returncode == 0:
                print(f"[Auto-Sync] Done (완료: {done_at})\n")
            else:
                print(f"[Auto-Sync] Failed (완료: {done_at})\n")
        except Exception as e:
            print(f"[Auto-Sync] Error: {e}\n")

    def _schedule_sync(self):
        with self.lock:
            if self.timer:
                self.timer.cancel()
            self.timer = threading.Timer(DEBOUNCE_BAS_SEC, self._on_debounce)
            self.timer.daemon = True
            self.timer.start()

    def _on_debounce(self):
        with self.lock:
            self.timer = None
        self._run_sync()

    def _is_vba_file(self, path):
        if not path:
            return False
        name = os.path.basename(path).lower()
        return name.endswith(".bas") or name.endswith(".frm")

    def _maybe_schedule(self, path):
        if path and self._is_vba_file(path):
            print("[Auto-Sync] 변경 감지: " + os.path.basename(path))
            self._schedule_sync()

    def on_modified(self, event):
        if event.is_directory:
            return
        self._maybe_schedule(event.src_path)

    def on_created(self, event):
        if event.is_directory:
            return
        self._maybe_schedule(event.src_path)

    def on_moved(self, event):
        if event.is_directory:
            return
        dest = getattr(event, "dest_path", None) or event.src_path
        self._maybe_schedule(dest)


def main():
    workbook_path = sys.argv[1] if len(sys.argv) > 1 else None
    if not workbook_path or not os.path.isfile(workbook_path):
        workbook_path = os.path.join(SCRIPT_DIR, "이루모스_생산관리.xlsb")
    if not os.path.isfile(workbook_path):
        print("ERROR: Workbook not found.")
        print(f"  Usage: python watch_and_export.py \"C:\\path\\to\\이루모스_생산관리.xlsb\"")
        sys.exit(1)

    workbook_path = os.path.abspath(workbook_path)

    if not os.path.isfile(SYNC_SCRIPT):
        print("ERROR: sync_to_excel.py not found.")
        sys.exit(1)

    print("=" * 50)
    print("VBA Watch - vba_modules -> Excel (Import only)")
    print("=" * 50)
    print(f"Workbook: {workbook_path}")
    print(f"vba_modules change -> Excel Import (debounce {DEBOUNCE_BAS_SEC}s)")
    print("\nPress Ctrl+C to exit")
    print("-" * 50)

    vba_handler = VbaModulesChangeHandler(workbook_path) if os.path.isdir(VBA_MODULES_DIR) else None

    observer = Observer()
    if vba_handler:
        observer.schedule(vba_handler, VBA_MODULES_DIR, recursive=True)
    else:
        print("WARNING: vba_modules folder not found.")
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        observer.join()
        print("\nStopped.")


if __name__ == "__main__":
    main()
