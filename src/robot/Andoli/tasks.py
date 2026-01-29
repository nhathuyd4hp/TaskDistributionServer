import contextlib
import io
import os
import subprocess
import sys
from pathlib import Path

from celery import shared_task
from celery.app.task import Context, Task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(bind=True, name="Andoli")
def main(
    self: Task,
    file: io.BytesIO | str = "xlsx",
):
    # ----- Metadata -----#
    context: Context = self.request
    id = context.id
    # ----- Download Asset -----#
    file_name = os.path.basename(file) if (file, str) else file.name
    save_path: Path = Path(__file__).resolve().parents[2] / "robot" / "Andoli" / file_name
    minio.fget_object(
        bucket_name=settings.TEMP_BUCKET,
        object_name=file,
        file_path=str(save_path),
    )
    new_path = save_path.with_name("Andoli納期確認送付.xlsx")
    save_path.replace(new_path)

    # ----- Exe Path -----#
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "Andoli" / "AnDoli_v2_7.py"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
            ],
            cwd=str(exe_path.parent),
            stdout=f,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
        )
        process.wait()

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"Andoli/{id}/Andoli.xlsx",
        file_path=str(new_path),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    with contextlib.suppress(Exception):
        new_path.unlink(missing_ok=True)
        Andoli_bot_log = exe_path.parent / "Andoli_bot_log.log"
        Andoli_bot_log.unlink(missing_ok=True)

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
