import shutil
import subprocess
import sys
import typing
from datetime import datetime
from pathlib import Path

import redis
from celery import shared_task

from src.core.config import settings
from src.core.logger import Log
from src.core.redis import REDIS_POOL
from src.service import ResultService as minio


@shared_task(bind=True, name="Furiwake Osaka")
def FuriwakeOsaka(
    self,
    工場: typing.Literal["大阪工場 製造データ", "栃木工場"],
    日付: datetime | str,
):
    logger = Log.get_logger(channel=self.request.id, redis_client=redis.Redis(connection_pool=REDIS_POOL))
    if isinstance(日付, str):
        日付 = datetime.fromisoformat(日付)
    日付: str = f"{日付.month}月{日付.day:02d}日"
    # --- #
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "FuriwakeOsaka" / "Main.py"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
                "--工場",
                "栃木工場" if 工場 == "栃木工場" else "大阪工場　製造データ",
                "--日付",
                日付,
            ],
            cwd=str(exe_path.parent),
            stdout=f,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
        )
        process.wait()

    # Clean
    bom_dir = exe_path.parent / "BOM"
    logs_dir = exe_path.parent / "Logs"
    access_token_dir = exe_path.parent / "Access_token"
    配車表_dir = exe_path.parent / "配車表"
    USB_dir = exe_path.parent / "▽USB"
    shutil.rmtree(bom_dir, ignore_errors=True)
    shutil.rmtree(logs_dir, ignore_errors=True)
    shutil.rmtree(access_token_dir, ignore_errors=True)
    shutil.rmtree(配車表_dir, ignore_errors=True)
    shutil.rmtree(USB_dir, ignore_errors=True)
    # Upload result
    result_dir = exe_path.parent / "Results"
    files = [f for f in result_dir.iterdir() if f.is_file()]
    if not files:
        raise RuntimeError("Results folder is empty")
    latest_file: Path = max(files, key=lambda f: f.stat().st_mtime)
    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"FuriwakeOsaka/{self.request.id}/{self.request.id}.xlsx",
        file_path=str(latest_file),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    try:
        latest_file.unlink(missing_ok=True)
        shutil.rmtree(result_dir, ignore_errors=True)
    except Exception as e:
        logger.error(e)
    return f"{settings.RESULT_BUCKET}/{result.object_name}"
