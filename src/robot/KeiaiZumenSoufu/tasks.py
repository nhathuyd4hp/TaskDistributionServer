import os
from src.core.config import settings
import shutil
import subprocess
from src.service import ResultService as minio
from celery import shared_task
from pathlib import Path

@shared_task(bind=True,name="Keiai Zumen Soufu")
def keiai_zumen_soufu(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "KeiaiZumenSoufu" / "Kizuku_V1_1_2.exe"
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()

    access_token_file = cwd_path / "Access_token_log"
    bot_log_file = cwd_path / "bot_log.log"
    
    try:
        access_token_file.unlink()
        bot_log_file.unlink()
    except Exception:
        pass

    logs_folder = cwd_path / "Access_token"

    for path in [logs_folder]:
        shutil.rmtree(path, ignore_errors=True)

    result_file = cwd_path / "Kizuku図面送付.xlsx"
    result = minio.fput_object(
        bucket_name=settings.MINIO_BUCKET,
        object_name=f"KeiaiZumenSoufu/{self.request.id}/{self.request.id}.xlsx",
        file_path=str(result_file),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    return result.object_name