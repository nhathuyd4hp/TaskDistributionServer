import io
import re
import tempfile

import pandas as pd
import redis
from celery import shared_task
from playwright.sync_api import sync_playwright

from src.core.config import settings
from src.core.logger import Log
from src.core.redis import REDIS_POOL
from src.robot.CapNhatDienTichWebAccess.automation import OCR, SharePoint, WebAccess
from src.service import ResultService as minio


@shared_task(bind=True, name="Cập nhật diện tích WebAccess")
def update_area_web_access(self):
    logger = Log.get_logger(channel=self.request.id, redis_client=redis.Redis(connection_pool=REDIS_POOL))
    ocr = OCR(logger=logger, poppler_path="src/resource/bin", tesseract_path="src/resource/Tesseract/tesseract.exe")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=[
                "--start-maximized",
            ],
            timeout=10000,
        )
        context = browser.new_context(
            no_viewport=True,
        )
        with WebAccess(
            domain="https://webaccess.nsk-cad.com/",
            username="hanh0704",
            password="159753",
            playwright=p,
            logger=logger,
            browser=browser,
            context=context,
        ) as wa:
            ケイアイ = wa.download_data("ケイアイ")
            秀光ビルド = wa.download_data("秀光ビルド")
            orders = pd.concat([ケイアイ, 秀光ビルド], ignore_index=True)
            orders = orders[["案件番号", "得意先名", "物件名", "延床平米", "資料リンク"]]
            orders = orders[pd.isna(orders["延床平米"])].reset_index(drop=True)
        with (
            SharePoint(
                domain="https://nskkogyo.sharepoint.com/",
                email="hanh3@nskkogyo.onmicrosoft.com",
                password="Got21095",
                playwright=p,
                logger=logger,
                browser=browser,
                context=context,
            ) as sp,
            tempfile.TemporaryDirectory() as temp_dir,
        ):
            total = orders.shape[0]
            for index, row in orders.iterrows():
                url = row["資料リンク"]
                if pd.isna(url):
                    continue
                logger.info(f"{url} [Remaining: {total-(index+1)}]")
                downloads = sp.download_files(
                    url=row["資料リンク"],
                    file=re.compile(r".*\.pdf$", re.IGNORECASE),
                    steps=[
                        re.compile("^割付図・エクセル$"),
                    ],
                    save_to=temp_dir,
                )
                if len(downloads) != 1:
                    continue
                pdf_path = downloads[0]
                area = ocr.get_area(pdf_path)
                if area is None:
                    continue
                orders.at[index, "延床平米 OCR"] = (
                    area
                    if wa.update(
                        case=str(row["案件番号"]),
                        area=str(area),
                    )
                    else f"Update Error ({area})"
                )
            # --- Upload to Minio
            excel_buffer = io.BytesIO()
            orders.to_excel(excel_buffer, index=False, engine="openpyxl")
            excel_buffer.seek(0)
            # --- #
            result = minio.put_object(
                bucket_name=settings.MINIO_BUCKET,
                object_name=f"CapNhatDienTich/{self.request.id}/{self.request.id}.xlsx",
                data=excel_buffer,
                length=excel_buffer.getbuffer().nbytes,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            return result.object_name
