import redis
import xlwings as xw
from celery import shared_task
from playwright.sync_api import sync_playwright

from src.core.config import settings
from src.core.logger import Log
from src.core.redis import REDIS_POOL
from src.robot.ShukoTaoSo.automation import AndPad, MailDealer, SharePoint
from src.service import ResultService as minio


def get_excel_path():
    return "src/resource/案件化.xlsm"


@shared_task(bind=True, name="Shuko Tạo Số")
def Shuko(self):
    logger = Log.get_logger(channel=self.request.id, redis_client=redis.Redis(connection_pool=REDIS_POOL))
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=[
                "--start-maximized",
            ],
        )
        context = browser.new_context(no_viewport=True)
        context.tracing.start(screenshots=True, snapshots=True, sources=True)
        with (
            MailDealer(
                domain="https://mds3310.maildealer.jp/",
                username=settings.MAIL_DEALER_USERNAME,
                password=settings.MAIL_DEALER_PASSWORD,
                playwright=p,
                logger=logger,
                browser=browser,
                context=context,
            ) as md,
            AndPad(
                domain="https://work.andpad.jp/",
                username=settings.ANDPAD_USERNAME,
                password=settings.ANDPAD_PASSWORD,
                playwright=p,
                browser=browser,
                logger=logger,
                context=context,
            ),
            SharePoint(
                domain="https://nskkogyo.sharepoint.com/",
                email="hanh3@nskkogyo.onmicrosoft.com",
                password="Got21095",
                playwright=p,
                logger=logger,
                browser=browser,
                context=context,
            ) as sp,
        ):
            # ---- Clear Data ---- #
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(get_excel_path())
            wb.macro("ClearSheet1Data")()
            ws = wb.sheets["シート"]
            # ---- Process ---- #
            logger.info("Mailbox: 専用アドレス・秀光ビルド")
            mails = md.mail_lists(mailbox="専用アドレス・秀光ビルド")
            mails = mails[mails[" 件名 "].str.contains("招待", na=False)]
            mails = mails[mails[" 担当者 "] == "--"]
            mails = mails[mails[" ラベル "] == ""]
            i: int = 2
            for _, row in mails.iterrows():
                logger.info(f"{row[" 件名 "]}[{row[" ID "]}]")
                success, fMatterID, orders_name, save_path, address = md.generate(row[" ID "])
                if success:
                    sp.upload_folder(
                        url="https://nskkogyo.sharepoint.com/sites/Shuuko/Shared Documents/Forms/AllItems.aspx",
                        folder_path=save_path,
                    )
                    ws[f"A{i}"].value = row[" ID "]
                    ws[f"B{i}"].value = fMatterID
                    ws[f"C{i}"].value = "秀光ビルド"
                    ws[f"D{i}"].value = orders_name
                    ws[f"K{i}"].value = address
                    logger.info(f"{row[" ID "]} - {fMatterID} - 秀光ビルド - {orders_name} - {address}")
                else:
                    logger.info(f"{row[" ID "]} - {save_path}")
                    ws[f"A{i}"].value = row[" ID "]
                i = i + 1
            wb.save()
            wb.close()
            app.quit()

    object_name = f"ShukoTaoSo/{self.request.id}/{self.request.id}.xlsm"

    result = minio.fput_object(
        bucket_name=settings.RESULT_BUCKET,
        file_path=get_excel_path(),
        object_name=object_name,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    return f"{settings.RESULT_BUCKET}/{result.object_name}"
