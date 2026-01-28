import subprocess
from pathlib import Path

from celery import shared_task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(bind=True, name="Toei Noukihikaku")
def KEIAI_ANKENKA(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "ToeiNoukihikaku" / "touei_noukihikaku_V1_4.exe"

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [str(exe_path)], cwd=str(exe_path.parent), stdout=f, stderr=subprocess.STDOUT, text=True
        )
        process.wait()

    result_file = exe_path.parent / "結果.xlsx"

    object_name = f"ToeiNoukihikaku/{self.request.id}.xlsx"

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        file_path=str(result_file),
        object_name=object_name,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
