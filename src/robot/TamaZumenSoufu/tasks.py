import subprocess
from pathlib import Path

from celery import shared_task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(bind=True, name="Tama Zumen Soufu")
def Tama_Zumen_Soufu(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "TamaZumenSoufu" / "Tamahome割付図送付_V_3.exe"
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()

    excel_filename = "図面送付結果.xlsx"
    excel_path = cwd_path / excel_filename

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"TamaZumenSoufu/{self.request.id}/TamaZumenSoufu.xlsx",
        file_path=str(excel_path),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
