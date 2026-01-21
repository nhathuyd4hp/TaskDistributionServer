import typing
from datetime import datetime

import redis
from celery import shared_task
from celery.app.task import Task

from src.core.logger import Log
from src.core.redis import REDIS_POOL


@shared_task(
    bind=True,
    name="Chuyển tên File từ Folder",
)
def main(
    self: Task,
    process_date: datetime | str,
    factory: typing.Literal["Shiga", "Toyo", "Chiba"],
):
    task_id = self.request.id
    logger = Log.get_logger(channel=task_id, redis_client=redis.Redis(connection_pool=REDIS_POOL))
    logger.info(f"Input: {process_date} - {factory}")
    if factory == "":
        raise RuntimeError("Invalid factory configuration: value is empty")
    pass
