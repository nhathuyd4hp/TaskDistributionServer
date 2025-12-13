import os
import tempfile
from celery import shared_task
from datetime import datetime
from src.core.config import settings
from src.robot.UpDataTochigi.automation import APISharePoint

@shared_task(bind=True)
def up_data_tochigi(
    self,
    process_date: datetime | str
):
    # Static Variable
    macro_file = "src/robot/UpDataTochigi/resource/マクロチェック(240819ver).xlsm"
    if isinstance(process_date, str):
        process_date = datetime.strptime(process_date, "%Y-%m-%d %H:%M:%S.%f").date()

    with tempfile.TemporaryDirectory() as temp_dir:
        DataTochigi = f"DataTochigi{process_date}.xlsx"
        api = APISharePoint(
            TENANT_ID=settings.API_SHAREPOINT_TENANT_ID,
            CLIENT_ID=settings.API_SHAREPOINT_CLIENT_ID,
            CLIENT_SECRET=settings.API_SHAREPOINT_CLIENT_SECRET,
        )
        UP = api.get_site("UP")
        Mouka = api.get_site("mouka")
        DataTochigi_ItemID = None
        DataTochigi_DriveID = None
        DataTochigi_SiteID = None

        Uploaded = api.download_item(
            site_id=UP.get("id"),
            breadcrumb=f"データUP一覧/{os.path.basename(DataTochigi)}",
            save_to=os.path.join(temp_dir,DataTochigi),
        )