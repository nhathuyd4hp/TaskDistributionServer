import subprocess
from pathlib import Path

from celery import shared_task
from celery.app.task import Task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(bind=True, name="Hinokiya Zumen Soufu")
def main(self: Task):
    id: str = self.request.id
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "HinokiyaZumenSoufu" / "ZumenSoufu_Henshin_V1.9.exe"
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [str(exe_path)], cwd=str(exe_path.parent), stdout=f, stderr=subprocess.STDOUT, text=True
        )
        process.wait()

    result_path = exe_path.parent / "図面送付結果_返信.xlsx"

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"HinokiyaZumenSoufu/{id}/result.xlsx",
        file_path=str(result_path),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
