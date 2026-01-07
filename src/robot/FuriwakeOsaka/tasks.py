import shutil
import subprocess
from pathlib import Path
import typing
from datetime import datetime
from celery import shared_task


@shared_task(bind=True, name="Furiwake Osaka")
def FuriwakeOsaka(
    self,
    工場: typing.Literal["大阪工場　製造データ", "栃木工場"],
    日付: datetime,
):
    return