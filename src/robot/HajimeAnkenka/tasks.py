import shutil
import subprocess
import sys
from pathlib import Path

import redis
from celery import shared_task
from celery.app.task import Context, Task

from src.core.config import settings
from src.core.logger import Log
from src.core.redis import REDIS_POOL
from src.service import ResultService as minio


@shared_task(bind=True, name="Hajime Ankenka")
def HajimeAnkenka(self: Task):
    context: Context = self.request
    id = context.id
    logger = Log.get_logger(channel=id, redis_client=redis.Redis(connection_pool=REDIS_POOL))

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "HajimeAnkenka" / "Hajime_loop_Test_v1_9.py"

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

    result_file = exe_path.parent / "Hajime_案件化.xlsm"
    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        object_name=f"HajimeAnkenka/{id}/Hajime.xlsm",
        file_path=str(result_file),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    paths = [
        exe_path.parent / "Access_token",
        exe_path.parent / "Hajime_shinki_bot_logs",
        exe_path.parent / "Access_token_log.txt",
    ]

    for path in paths:
        try:
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            if path.is_file():
                path.unlink(missing_ok=True)
        except Exception as e:
            logger.error(e)

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
