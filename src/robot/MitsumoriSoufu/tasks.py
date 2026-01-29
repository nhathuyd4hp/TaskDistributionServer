import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from celery import shared_task
from celery.app.task import Task


@shared_task(bind=True, name="住建 Mitsumori Soufu")
def main(
    self: Task,
    from_date: datetime | str,
    to_date: datetime | str,
):
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "MitsumoriSoufu" / "New.py"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
                "--from-date",
                from_date,
                "--to-date",
                to_date,
            ],
            cwd=str(exe_path.parent),
            stdout=f,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
        )
        process.wait()

    for path in [
        exe_path / "Access_token",
        exe_path / "Logs",
        exe_path / "Access_token_log",
    ]:
        if path.is_file():
            path.unlink(missing_ok=True)
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
