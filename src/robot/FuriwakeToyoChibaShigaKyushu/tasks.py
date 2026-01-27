import os
import re
import tempfile
import typing
from datetime import datetime

import pandas as pd
import redis
from celery import shared_task
from celery.app.task import Task
from playwright.sync_api import sync_playwright

from src.core.config import settings
from src.core.logger import Log
from src.core.redis import REDIS_POOL
from src.robot.FuriwakeToyoChibaShigaKyushu.automation import SharePoint
from src.service import ResultService as minio


def link_data(factory: typing.Literal["Shiga", "Toyo", "Chiba", "Kyushu"]) -> str:
    if factory == "Shiga":
        return "https://nskkogyo.sharepoint.com/sites/shiga/Shared Documents/Forms/AllItems.aspx?id=/sites/shiga/Shared Documents/滋賀工場 製造データ"  # noqa
    if factory == "Toyo":
        return "https://nskkogyo.sharepoint.com/sites/toyohashi/Shared Documents/Forms/AllItems.aspx?id=/sites/toyohashi/Shared Documents/豊橋工場 製造データ"  # noqa
    if factory == "Chiba":
        return "https://nskkogyo.sharepoint.com/sites/nskhome/Shared Documents/Forms/AllItems.aspx?id=/sites/nskhome/Shared Documents/千葉工場 製造データ"  # noqa
    if factory == "Kyushu":
        return "https://nskkogyo.sharepoint.com/sites/kyuusyuukouzyou"  # noqa
    raise RuntimeError(f"Invalid factory: {factory}")


@shared_task(
    bind=True,
    name="Furiwake Toyo-Chiba-Shiga-Kyushu",
)
def main(
    self: Task,
    process_date: datetime | str,
    factory: typing.Literal["Shiga", "Toyo", "Chiba", "Kyushu"],
):
    task_id = self.request.id
    logger = Log.get_logger(channel=task_id, redis_client=redis.Redis(connection_pool=REDIS_POOL))
    if factory not in ["Shiga", "Toyo", "Chiba", "Kyushu"]:
        raise ValueError("Invalid factory configuration: factory must be one of ['Shiga', 'Toyo', 'Chiba','Kyushu']")
    # ----- Convert process_date into datetime ----- #
    if isinstance(process_date, str):
        process_date: datetime = datetime.strptime(process_date, "%Y-%m-%d %H:%M:%S.%f")
    with (
        sync_playwright() as p,
        tempfile.TemporaryDirectory() as temp_dir,
    ):
        try:
            browser = p.chromium.launch(headless=False, args=["--start-maximized"])
            context = browser.new_context(no_viewport=True)
            logger.info("login sharepoint")
            with SharePoint(
                domain=settings.SHAREPOINT_DOMAIN,
                username=settings.SHAREPOINT_EMAIL,
                password=settings.SHAREPOINT_PASSWORD,
                playwright=p,
                browser=browser,
                context=context,
                logger=logger,
            ) as sp:
                if factory == "Kyushu":
                    files = sp.download(
                        url=link_data(factory),
                        file=re.compile(r".*\.(xls|xlsx|xlsm|xlsb)$", re.IGNORECASE),
                        steps=[
                            re.compile("^九州工場 製造データー$"),
                            re.compile(rf"^0?{process_date.month}月0?{process_date.day}日配送分$"),
                            re.compile("^確定データ$"),
                        ],
                        save_to=temp_dir,
                    )
                else:
                    files = sp.download(
                        url=link_data(factory),
                        file=re.compile(r".*\.(xls|xlsx|xlsm|xlsb)$", re.IGNORECASE),
                        steps=[
                            re.compile(rf"^0?{process_date.month}月0?{process_date.day}日配送分$"),
                            re.compile("^確定データ$"),
                        ],
                        save_to=temp_dir,
                    )
                files = [os.path.basename(file) for file in files]
                for i, file in enumerate(files):
                    if match := re.search(r"＿(.*?)＿", file):
                        file = match.group(1)
                    file = re.sub(r"\([^)]*(am|pm)[^)]*\)", "", file, flags=re.IGNORECASE)
                    files[i] = file
                result_path = os.path.join(temp_dir, "filenames.xlsx")
                pd.DataFrame({"filename": files}).to_excel(result_path, index=False)
                result = minio.fput_object(
                    bucket_name=settings.RESULT_BUCKET,
                    object_name=f"ChuyenTenFileTuFolder/{task_id}/{factory}.xlsx",
                    file_path=result_path,
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                return f"{settings.RESULT_BUCKET}/{result.object_name}"
        except Exception as e:
            raise e
