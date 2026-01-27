import contextlib
import os
import shutil
import subprocess
from pathlib import Path

from celery import shared_task

from src.core.config import settings
from src.service import ResultService as minio


@shared_task(bind=True, name="Osaka Mitsumori Soufu")
def OsakaMitsumoriSoufu(self):
    exe_path = (
        Path(__file__).resolve().parents[2] / "robot" / "OsakaMitsumoriSoufu" / "「大阪・インド」見積書送付_V1.5.exe"
    )
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()
    # Clean
    Access_token_log = cwd_path / "Access_token_log"
    try:
        Access_token_log.unlink()
    except Exception:
        pass
    logs_folder = cwd_path / "Logs"
    access_token_folder = cwd_path / "Access_token"
    ankens_folder = cwd_path / "Ankens"
    for path in (logs_folder, access_token_folder, ankens_folder):
        shutil.rmtree(path, ignore_errors=True)

    reports_folder = cwd_path / "ProgressReports"
    latest_pdf: Path | None = max(
        reports_folder.glob("*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        default=None,
    )
    if not latest_pdf:
        raise FileNotFoundError("ProgressReports: FileNotFound")

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"OsakaMitsumoriSoufu/{self.request.id}/{os.path.basename(latest_pdf)}",
        file_path=str(latest_pdf),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    with contextlib.suppress(Exception):
        latest_pdf.unlink()
        shutil.rmtree(reports_folder)

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
