from celery import shared_task
from celery.app.task import Context, Task

from src.core.inactive_task import InactiveTask


@shared_task(bind=True, name="Kenshin Yamaha Zumen Soufu", base=InactiveTask)
def main(self: Task):
    context: Context = self.request
    id: str = context.id
    return id
