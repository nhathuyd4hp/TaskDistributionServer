import subprocess
from pathlib import Path

from celery import shared_task

from src.core.config import settings
from src.service import ResultService as minio


@shared_task(bind=True, name="Tama Ankenka")
def Tama_Ankenka(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "TamaAnkenka" / "タマホーム_案件化+資料UP_V1_2.exe"

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [str(exe_path)], cwd=str(exe_path.parent), stdout=f, stderr=subprocess.STDOUT, text=True
        )
        process.wait()

    file_path = exe_path.parent / "結果.xlsx"

    result = minio.fput_object(
        bucket_name=settings.MINIO_BUCKET,
        object_name=f"TamaAnkenka/{self.request.id}/{self.request.id}.xlsx",
        file_path=str(file_path),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    file_path.unlink()

    return result.object_name
