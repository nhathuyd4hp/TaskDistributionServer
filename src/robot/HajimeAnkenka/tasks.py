from celery import shared_task
from celery.app.task import Context, Task

from src.core.inactive_task import InactiveTask


@shared_task(bind=True, name="Hajime Ankenka", base=InactiveTask)
def HajimeAnkenka(self: Task):
    context: Context = self.request
    id = context.id
    return id
