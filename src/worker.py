from celery import Celery

from src.core.config import settings
from src.robot import *  # noqa
from src.worker_signals import *  # noqa

Worker = Celery(
    "orchestration",
    broker=settings.REDIS_CONNECTION_STRING,
    backend=settings.REDIS_CONNECTION_STRING,
)
