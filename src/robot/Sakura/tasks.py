import logging
import os
import re
import shutil
import tempfile
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd
import redis
import xlwings as xw
from celery import shared_task
from selenium import webdriver
from xlwings.main import Sheet
from xlwings.utils import col_name

from src.core.config import settings
from src.core.logger import Log
from src.core.redis import REDIS_POOL
from src.core.type import UserCancelledError
from src.robot.Sakura.automation.bot import MailDealer, SharePoint, WebAccess
from src.service import ResultService as minio

download_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
# -- Chrome Options
options = webdriver.ChromeOptions()
options.add_argument("--disable-notifications")
options.add_argument("--disable-logging")
options.add_argument("--log-level=3")
options.add_argument("--silent")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
    },
)


def main(
    output: str,
    logger: logging.Logger,
    checker: redis.Redis,
    task_id: str,
):
    # From To
    to_date = datetime.now().replace(day=20)
    if to_date.month == 1:
        from_date = to_date.replace(year=to_date.year - 1, month=12, day=21)
    else:
        from_date = to_date.replace(month=to_date.month - 1, day=21)
    from_date = from_date.strftime("%Y/%m/%d")
    to_date = to_date.strftime("%Y/%m/%d")
    logger.info(f"{from_date} ~ {to_date}")
    # Download
    logger.info("login WebAccess")
    with WebAccess(
        url="https://webaccess.nsk-cad.com/",
        username=settings.WEBACCESS_USERNAME,
        password=settings.WEBACCESS_PASSWORD,
        logger=logger,
        options=options,
    ) as web_access:
        logger.info("download data")
        data = web_access.get_order_list(
            building_name="009300",
            delivery_date=[from_date, to_date],
            fields=[
                "案件番号",
                "得意先名",
                "物件名",
                "確未",
                "確定納期",
                "曜日",
                "追加不足",
                "配送先住所",
                "階",
                "資料リンク",
            ],
        )
    if data.empty:
        return
    # ---- Download files ----
    prices = []
    data = data[data["追加不足"] != "不足"]
    logger.info("login sharepoint")
    with SharePoint(
        url="https://nskkogyo.sharepoint.com/",
        username="vietnamrpa@nskkogyo.onmicrosoft.com",
        password="Robot159753",
        logger=logger,
        options=options,
    ) as share_point:
        for url in data["資料リンク"]:
            logger.info(url)
            downloads = share_point.download(
                site_url=url,
                file_pattern="(見積書|見積もり)/.*.(xlsm|xlsx|xls)$",
            )
            if not downloads:
                raise RuntimeError("Không có file")
            if all(status for _, _, status in downloads):
                prices.append(downloads[0][2])
                continue
            if len(downloads) != 1:
                raise RuntimeError("Có nhiều file")
            for _, file, status in downloads:
                price = None
                found = False
                if status:
                    continue
                for sheet in pd.ExcelFile(file, engine="openpyxl", engine_kwargs={"read_only": True}).sheet_names:
                    sheet: pd.DataFrame = pd.read_excel(file, sheet_name=sheet)
                    for _, row in sheet.iterrows():
                        row: str = " ".join(str(cell) for cell in row)
                        if match := re.search(r"税抜金額[^\d]*([\d,]+(?:\.\d+)?)", row):
                            price = match.group(1)
                            price = Decimal(price).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
                            if price != 0:
                                found = True
                            break
                    if found:
                        break
                prices.append(price)
                logger.info(f"{file}: {price}")
    # Process
    data["金額（税抜）"] = prices
    data["金額（税抜）"] = pd.to_numeric(data["金額（税抜）"], errors="coerce").fillna(0)
    data.drop(columns=["資料リンク"], inplace=True)
    # Append Row
    empty_row = pd.Series({col: pd.NA for col in data.columns})
    append_row = {col: pd.NA for col in data.columns}
    append_row[list(data.columns)[-3]] = "合計"
    append_row[list(data.columns)[-1]] = data["金額（税抜）"].sum()
    data = pd.concat([data, pd.DataFrame([empty_row.to_dict(), append_row])], ignore_index=True)
    # Save
    excel_file = os.path.join(output, f"{datetime.today().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
    logger.info("to_excel")
    data.to_excel(
        os.path.join(output, excel_file),
        index=False,
    )
    app = None
    wb = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(excel_file)
        sheet: Sheet = wb.sheets[0]
        # AutoFitColumn
        sheet.autofit()
        # Header
        sheet.api.PageSetup.LeftHeader = f"さくら建設　鋼製野縁納材報告（{from_date}-{to_date}）　"
        sheet.api.PageSetup.RightHeader = datetime.now().strftime("%Y/%m/%d")
        # Format
        ## - Tô màu Header
        sheet.range(f"A1:{col_name(data.shape[1])}1").color = (166, 166, 166)
        ## - Tô màu ô "合計"
        sheet.range(f"H{data.shape[0] + 1}").color = (166, 166, 166)
        ## - Định dang cột J 金額（税抜）(12345 -> 12,345)
        sheet.range(f"J2:J{data.shape[0] + 1}").number_format = "#,##0"
        # Landspace
        sheet.api.PageSetup.Orientation = 2
        # All Border
        rng = sheet.range(f"A1:J{data.shape[0] + 1}")
        for i in [7, 8, 9, 10, 11, 12]:
            border = rng.api.Borders(i)
            border.LineStyle = 1
            border.Weight = 2
            border.ColorIndex = 0
        wb.save()
        pdfFile = os.path.join(output, f"さくら建設　鋼製野縁納材報告（{from_date} - {to_date}).pdf".replace("/", "-"))
        # ---- Export in one page #
        sheet.api.PageSetup.Zoom = False
        sheet.api.PageSetup.FitToPagesWide = 1
        sheet.api.PageSetup.FitToPagesTall = 1
        logger.info(f"export pdf: {pdfFile}")
        sheet.to_pdf(pdfFile)
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()
        logger.info("login mail dealer")
        with MailDealer(
            url="https://mds3310.maildealer.jp/",
            username=settings.MAIL_DEALER_USERNAME,
            password=settings.MAIL_DEALER_PASSWORD,
            logger=logger,
            options=options,
        ) as mail_dealer:
            logger.info("send_mail")
            mail_dealer.send_mail(
                fr="kantou@nsk-cad.com",
                to="ikeda.k@jkenzai.com",
                subject=f"さくら建設　鋼製野縁納材報告書（{from_date}～{to_date}）",
                content=f"""
                ジャパン建材　池田様

                いつもお世話になっております。

                さくら建設　鋼製野縁納材報告書（{from_date}～{to_date}）
                を送付致しましたので、ご査収の程よろしくお願い致します。

                ----・・・・・----------・・・・・----------・・・・・-----

                　エヌ・エス・ケー工業㈱　横浜営業所
                中山　知凡

                　〒222-0033
                　横浜市港北区新横浜２-４-６　マスニ第一ビル８F-B
                　TEL:(045)595-9165 / FAX:(045)577-0012

                -----・・・・・----------・・・・・----------・・・・・-----
                """,
                attachments=[
                    os.path.abspath(pdfFile),
                ],
            )
    shutil.rmtree(download_path)
    return pdfFile


@shared_task(
    bind=True,
    name="Sakura",
)
def Sakura(self):
    logger = Log.get_logger(channel=self.request.id, redis_client=redis.Redis(connection_pool=REDIS_POOL))
    checker = redis.Redis(connection_pool=REDIS_POOL)
    task_id = self.request.id
    if checker.get(task_id) is not None:
        raise UserCancelledError()
    with tempfile.TemporaryDirectory() as temp_dir:
        pdfFile = main(
            output=temp_dir,
            logger=logger,
            checker=checker,
            task_id=task_id,
        )
        result = minio.fput_object(
            bucket_name=settings.RESULT_BUCKET,
            object_name=f"Sakura/{self.request.id}/Sakura.pdf",
            file_path=pdfFile,
            content_type="application/pdf",
        )
        return f"{settings.RESULT_BUCKET}/{result.object_name}"
