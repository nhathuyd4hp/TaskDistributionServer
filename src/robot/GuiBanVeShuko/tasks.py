import os
import subprocess
from pathlib import Path

from celery import shared_task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(bind=True, name="Gửi Bản Vẽ Shuko")
def GuiBanVeShuko(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "GuiBanVeShuko" / "GuiBanVeShuko.exe"
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"

    with open(log_file, "a", encoding="utf-8") as f:
        try:
            process = subprocess.Popen(
                [str(exe_path)],
                cwd=str(cwd_path),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                env=env,
                text=True,
                encoding="utf-8",
                bufsize=1,
            )

            for line in process.stdout:
                f.write(line)
                f.flush()

            process.wait()

        except Exception as e:
            f.write(f"\n[CRITICAL ERROR] Python subprocess failed: {str(e)}\n")

    xlsx_files = list(cwd_path.glob("*.xlsx"))
    if xlsx_files:
        latest_file = max(xlsx_files, key=lambda p: p.stat().st_mtime)
        result = minio.fput_object(
            bucket_name=settings.RESULT_BUCKET,
            object_name=f"GuiBanVeShuko/{self.request.id}/GuiBanVeShuko.xlsx",
            file_path=str(latest_file),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return f"{settings.RESULT_BUCKET}/{result.object_name}"
