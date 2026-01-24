import contextlib
import io
import os
import shutil
import subprocess
import sys
import zipfile
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
    src_path = os.path.basename(file) if (file, str) else file.name
    src_path = os.path.abspath(src_path)
    # ----- Download ----- #
    minio.fget_object(
        bucket_name=settings.TEMP_BUCKET,
        object_name=file,
        file_path=os.path.basename(file),
    )
    # ----- Check FileType (XLSM) ----- #
    if not os.path.basename(file).lower().endswith(".xlsm"):
        if os.path.exists(src_path):
            os.remove(src_path)
        raise ValueError("Chỉ chấp nhận file định dạng .xlsm")
    # ----- Run ----- #
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "Zenbu" / "Main.py"

    dst_path = str(exe_path.parent / os.path.basename(src_path))
    shutil.move(src_path, dst_path)

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
                "--file",
                str(dst_path),
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
        exe_path.parent / "Logs",
        exe_path.parent / "Access_token_log",
        exe_path.parent / os.path.basename(dst_path),
    ]

    for p in paths:
        with contextlib.suppress(Exception):
            if p.is_file():
                p.unlink()
            if p.is_dir():
                shutil.rmtree(p)
    # ----- Zip ----- #
    paths = [
        exe_path.parent / "Ankens",
        exe_path.parent / "ProgressReports",
    ]
    zip_path = exe_path.parent / "Zenbu.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for folder in paths:
            for root, _, files in os.walk(folder):
                for file in files:
                    full_path = Path(root) / file
                    arcname = full_path.relative_to(exe_path.parent)
                    zipf.write(full_path, arcname)
    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"Zenbu/{self.request.id}/Zenbu.zip",
        file_path=str(zip_path),
        content_type="application/zip",
    )
    for p in paths:
        with contextlib.suppress(Exception):
            if p.is_file():
                p.unlink()
            if p.is_dir():
                shutil.rmtree(p)
    with contextlib.suppress(Exception):
        zip_path.unlink()
    return f"{settings.RESULT_BUCKET}/{result.object_name}"
