import io
import shutil
import subprocess
import sys
from pathlib import Path

from celery import shared_task


@shared_task(bind=True, name="Zenbu")
def Zenbu(
    self,
    file: io.BytesIO | str = "Zenbu",
):
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "Zenbu" / "Main.py"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
            ],
            cwd=str(exe_path.parent),
            stdout=f,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
        )
        process.wait()
    paths = [
        exe_path.parent / "Access_token",
        exe_path.parent / "Ankens",
        exe_path.parent / "Logs",
        exe_path.parent / "Access_token_log",
    ]

    for p in paths:
        try:
            if p.is_file():
                p.unlink()
            elif p.is_dir():
                shutil.rmtree(p)
        except Exception:
            pass
