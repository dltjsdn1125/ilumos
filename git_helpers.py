"""
Git 커밋 헬퍼: sync/export 성공 시 자동 커밋
데이터 손실 방지를 위해 수정마다 GitHub에 커밋
"""
import os
import subprocess
from datetime import datetime


def git_commit_after_change(action: str, extra_msg: str = "") -> bool:
    """
    Git 저장소가 있으면 변경사항을 add & commit
    action: "sync" | "export" | "수동"
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not os.path.isdir(os.path.join(script_dir, ".git")):
        return False
    kwargs = {"cwd": script_dir, "capture_output": True}
    if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    try:
        msg = f"VBA {action}: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        if extra_msg:
            msg += f" - {extra_msg}"
        subprocess.run(["git", "add", "-A"], **kwargs, check=True)
        result = subprocess.run(["git", "commit", "-m", msg], **kwargs, text=True)
        if result.returncode == 0:
            print(f"  [Git] 커밋 완료: {msg}")
            return True
        return False
    except Exception as e:
        print(f"  [Git] 커밋 건너뜀: {e}")
        return False
