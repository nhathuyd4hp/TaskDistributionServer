import traceback

import redis
from celery import signals
from celery.app.task import Context, Task
from celery.result import AsyncResult
from celery.worker.consumer.consumer import Consumer
from sqlmodel import Session, select

from src.core.config import settings
from src.core.redis import REDIS_POOL
from src.core.type import UserCancelledError
from src.model import Error, Runs
from src.model.runs import Status


@signals.worker_ready.connect
def start_up(sender: Consumer, **kwargs):
    with Session(settings.db_engine) as session:
        statement = select(Runs).where(Runs.status == Status.PENDING)
        records: list[Runs] = session.exec(statement).all()
        for record in records:
            result = AsyncResult(id=record.id, app=sender.app)
            if result.state in ["PENDING", "FAILURE"]:
                record.status = Status.FAILURE
                session.add(record)
        session.commit()


@signals.task_prerun.connect
def task_prerun_handler(sender=None, task_id=None, **kwargs):
    with Session(settings.db_engine) as session:
        statement = select(Runs).where(Runs.id == task_id)
        record: Runs = session.exec(statement).one_or_none()
        if record:
            record.status = Status.PENDING
            session.add(record)
        else:
            record = Runs(
                id=task_id,
                robot=sender.name,
                parameters=kwargs.get("kwargs") if kwargs.get("kwargs") else None,
                status=Status.PENDING,
            )
            session.add(record)
        session.commit()
        message = f"""\n
{record.robot} bắt đầu
----------------------
ID: {record.id}
"""
        redis.Redis(connection_pool=REDIS_POOL).publish("CELERY", message)


@signals.task_success.connect
def task_success_handler(sender=None, result=None, **kwargs):
    with Session(settings.db_engine) as session:
        statement = select(Runs).where(Runs.id == sender.request.id)
        record: Runs = session.exec(statement).one_or_none()
        if record is None:
            return
        record.status = Status.SUCCESS
        record.result = "" if result is None else str(result)
        session.add(record)
        session.commit()
        message = f"""\n
{record.robot} hoàn thành
-------------------------
ID: {record.id}
"""
        redis.Redis(connection_pool=REDIS_POOL).publish("CELERY", message)


@signals.task_failure.connect
def task_failure_handler(sender: Task, exception: Exception, **kwargs):
    context: Context = sender.request
    with Session(settings.db_engine) as session:
        statement = select(Runs).where(Runs.id == context.id)
        record = session.exec(statement).one_or_none()
        if record is None:
            return
        if isinstance(exception, UserCancelledError):
            record.status = Status.CANCEL
            record.result = exception.message
            session.add(record)
        else:
            error = Error(
                run_id=context.id,
                error_type=type(exception).__name__,
                message=str(exception),
                traceback="".join(traceback.format_tb(exception.__traceback__)),
            )
            record.status = Status.FAILURE
            record.result = type(exception).__name__
            session.add(error)
            session.add(record)
        session.commit()
        message = f"""\n
{record.robot} thất bại
-----------------------
ID: {record.id}
"""
        redis.Redis(connection_pool=REDIS_POOL).publish("CELERY", message)
