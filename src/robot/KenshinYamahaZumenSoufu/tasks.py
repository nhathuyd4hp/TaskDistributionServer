from celery import shared_task
from celery.app.task import Context, Task


@shared_task(bind=True)
def main(self: Task):
    context: Context = self.request
    id: str = context.id
    return id
