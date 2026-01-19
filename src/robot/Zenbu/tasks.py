import io
import os
import shutil
import subprocess
import sys
from pathlib import Path

from celery import shared_task

from src.core.config import settings
from src.service import ResultService as minio


@shared_task(bind=True, name="Zenbu")
def Zenbu(
    self,
    file: io.BytesIO | str = "Zenbu",
):
    if file == "null":
        raise ValueError("Chỉ chấp nhận file định dạng .xlsm")
    # ----- Save Path ----- #
    local_file_path = os.path.basename(file) if (file, str) else file.name
    local_file_path = os.path.abspath(local_file_path)
    # ----- Download ----- #
    minio.fget_object(
        bucket_name=settings.TEMP_BUCKET,
        object_name=file,
        file_path=os.path.basename(file),
    )
    # ----- Check FileType (XLSM) ----- #
    if not os.path.basename(file).lower().endswith(".xlsm"):
        if os.path.exists(local_file_path):
            os.remove(local_file_path)
        raise ValueError("Chỉ chấp nhận file định dạng .xlsm")
    # ----- Run ----- #
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "Zenbu" / "Main.py"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
                "--file",
                local_file_path,
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
