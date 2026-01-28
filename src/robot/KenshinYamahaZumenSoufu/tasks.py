import io
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from celery import shared_task
from celery.app.task import Context, Task

from src.core.config import settings
from src.service import StorageService as minio


@shared_task(bind=True, name="Kenshin Yamaha Zumen Soufu")
def main(self: Task, file: io.BytesIO | str = "xlsx"):
    # ----- Metadata -----#
    context: Context = self.request
    id = context.id
    # ----- Download Asset -----#
    file_name = os.path.basename(file) if (file, str) else file.name
    save_path: Path = Path(__file__).resolve().parents[2] / "robot" / "KenshinYamahaZumenSoufu" / file_name

    minio.fget_object(
        bucket_name=settings.TEMP_BUCKET,
        object_name=file,
        file_path=str(save_path),
    )
    new_path = save_path.with_name("健新～ヤマダホームズの図面送付.xlsx")
    save_path.replace(new_path)

    # ----- Exe Path -----#
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "KenshinYamahaZumenSoufu" / "Main.py"

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
        object_name=f"KenshinYamahaZumenSoufu/{id}/{id}.xlsx",
        file_path=str(new_path),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    logs_folder = exe_path.parent / "Logs"
    access_token_folder = exe_path.parent / "Access_token"
    download_folder = exe_path.parent / "Downloaded_Zumen"
    Access_token_log = exe_path.parent / "Downloaded_Zumen"
    for path in (logs_folder, access_token_folder, download_folder, Access_token_log):
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        if path.is_file():
            path.unlink()

    for ext in ("*.png", "*.jpg", "*.jpeg"):
        for img_file in exe_path.parent.glob(ext):
            img_file.unlink(missing_ok=True)

    ProgressReports = exe_path.parent / "ProgressReports"
    latest_pdf: Path | None = max(
        ProgressReports.glob("*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        default=None,
    )
    if not latest_pdf:
        raise FileNotFoundError("ProgressReports: FileNotFound")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_zip_path = Path(temp_dir) / "result.zip"
        with zipfile.ZipFile(temp_zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(new_path, arcname=new_path.name)
            archive.write(latest_pdf, arcname=latest_pdf.name)
        object_name = f"KenshinYamahaZumenSoufu/{self.request.id}/report.zip"
        result = minio.fput_object(
            bucket_name=settings.RESULT_BUCKET,
            object_name=object_name,
            file_path=str(temp_zip_path),
            content_type="application/zip",
        )
        return f"{settings.RESULT_BUCKET}/{result.object_name}"
