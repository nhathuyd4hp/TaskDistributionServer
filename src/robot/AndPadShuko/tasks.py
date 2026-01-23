import os
import re
import tempfile
from datetime import date, datetime

import numpy as np
import pandas as pd
import redis
from celery import shared_task
from celery.app.task import Task
from playwright.sync_api import sync_playwright

from src.core.config import settings
from src.core.logger import Log
from src.core.redis import REDIS_POOL
from src.core.type import UserCancelledError
from src.robot.AndPadShuko.automation import MailDealer, WebAccess
from src.service import ResultService as minio


@shared_task(
    bind=True,
    name="AndPad Shuko",
)
def main(self: Task):
    # ----- #
    checker = redis.Redis(connection_pool=REDIS_POOL)
    task_id = self.request.id
    if checker.get(task_id) is not None:
        raise UserCancelledError()
    # ----- #
    logger = Log.get_logger(channel=task_id, redis_client=redis.Redis(connection_pool=REDIS_POOL))
    logger.info(date.today())
    # ----- #
    with (
        sync_playwright() as p,
        tempfile.TemporaryDirectory() as temp_dir,
    ):
        ResultFile = f"{date.today().strftime('%d-%m-%Y')}.xlsx"
        browser = p.chromium.launch(
            headless=False,
            args=["--start-maximized"],
        )
        try:
            context = browser.new_context(no_viewport=True)
            context.tracing.start(screenshots=True, snapshots=True, sources=True)
            with (
                MailDealer(
                    domain="https://mds3310.maildealer.jp/",
                    username=settings.MAIL_DEALER_USERNAME,
                    password=settings.MAIL_DEALER_PASSWORD,
                    playwright=p,
                    browser=browser,
                    context=context,
                    logger=logger,
                ) as md,
                WebAccess(
                    domain="https://webaccess.nsk-cad.com/",
                    username=settings.WEBACCESS_USERNAME,
                    password=settings.WEBACCESS_PASSWORD,
                    playwright=p,
                    browser=browser,
                    context=context,
                    logger=logger,
                ) as wa,
            ):
                if checker.get(task_id) is not None:
                    raise UserCancelledError()
                logger.info("Mail box: 専用アドレス・秀光ビルド")
                mails = md.mail_box("専用アドレス・秀光ビルド")
                mails = mails[mails[" 件名 "].str.contains("【ANDPAD】", na=False)]
                mails = mails[(pd.isna(mails[" 担当者 "])) | (mails[" 担当者 "] == "--")]
                for column in mails.columns:
                    if not column:
                        mails.drop(column, axis=1, errors="ignore", inplace=True)
                mails.to_excel(os.path.join(temp_dir, ResultFile), index=False)
                mails = pd.read_excel(os.path.join(temp_dir, ResultFile))
                for index, row in mails.iterrows():
                    if checker.get(task_id) is not None:
                        raise UserCancelledError()
                    id = row[" ID "]
                    subject = row[" 件名 "]
                    # ---- #
                    logger.info(f"{id} - {subject} [Remain: {len(mails) - (index + 1)}]")
                    # ---- #
                    if match := re.search(r"([\w\W]+?新築工事)", subject):
                        案件名_物件名 = match.group(1)
                        案件名_物件名 = (
                            案件名_物件名.replace("新築工事", "").replace("【ANDPAD】", "").replace("Fw:", "").strip()
                        )
                        data = wa.download_data(案件名_物件名)
                        if data.shape[0] != 1:
                            logger.warning("Không tim thấy ở WebAccess")
                            mails.at[index, " Note "] = "Không tim thấy ở WebAccess"
                            continue
                        案件番号 = data.iloc[0]["案件番号"]
                        確定納期 = data.iloc[0]["確定納期"]
                        logger.info(f"案件番号: {案件番号} - 確定納期: {確定納期}")
                        if isinstance(確定納期, float) and (np.isnan(確定納期) or pd.isna(確定納期)):
                            logger.warning("Chưa cập ngày giao hàng")
                            mails.at[index, " Note "] = "Chưa cập ngày giao hàng"
                            continue
                        delivery_date = datetime.strptime(確定納期, "%Y/%m/%d").date()
                        if checker.get(task_id) is not None:
                            raise UserCancelledError()
                        mails.at[index, " Note "] = md.update_mail(
                            mail_id=str(id),
                            label="vietnamrpa",
                            fMatterID=str(案件番号),
                            comment=(
                                "納材済" if delivery_date < datetime.today().date() else f"access納期 {delivery_date}"
                            ),
                        )
                    else:
                        logger.warning("Không tìm thấy 案件名/物件名")
                        mails.at[index, " Note "] = "Không tìm thấy 案件名/物件名"
                mails.to_excel(os.path.join(temp_dir, ResultFile), index=False)
                result = minio.fput_object(
                    bucket_name=settings.RESULT_BUCKET,
                    object_name=f"AndPadShuko/{ResultFile}",
                    file_path=os.path.join(temp_dir, ResultFile),
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                return f"{settings.RESULT_BUCKET}/{result.object_name}"
        except Exception as e:
            trace_file = os.path.join(temp_dir, f"{task_id}.zip")
            context.tracing.stop(path=trace_file)
            minio.fput_object(
                bucket_name=settings.TRACE_BUCKET,
                object_name=os.path.basename(trace_file),
                file_path=trace_file,
                content_type="application/zip",
            )
            raise e
