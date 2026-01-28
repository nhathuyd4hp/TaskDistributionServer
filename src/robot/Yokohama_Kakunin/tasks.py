import subprocess
import sys
from datetime import datetime
from pathlib import Path

from celery import shared_task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(
    bind=True,
    name="Yokohama Kakunin",
)
def YokohamaKakunin(
    self,
    from_date: datetime | str,
    to_date: datetime | str,
):
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "Yokohama_Kakunin" / "Main.py"
    cwd_path = exe_path.parent

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
                "--task-id",
                self.request.id,
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

    result_file = cwd_path / "結果.xlsx"

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"YokohamaKakunin/{self.request.id}/YokohamaKakunin.xlsx",
        file_path=str(result_file),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
