from src.core.config import settings
import io
import datetime
from src.service import ResultService as minio
import logging
import re
import pandas as pd
from celery import shared_task
import concurrent.futures
from src.robot.ToeiXacNhanNouki.automation import MailDealer,Touei,WebAccess


TAB_NAME = None # "新着"
MAIL_BOX = '専用アドレス・飯田GH/≪ベトナム納期≫東栄(FAX・メール)'
TASK = "鋼製野縁"
FIELDS = ['確定納期', '案件番号', '物件名','配送先住所']
PROCESS_CONSTRUCTIONS = ["仙台施工","郡山施工","浜松施工","東海施工","関西施工","岡山施工","広島施工","福岡施工","熊本施工","東京施工","神奈川施工"]

def process_schedu_email_content(content:str) -> list[str]:
    buildings = []
    def split_raw_buildings(content) -> list[str]:    
        pattern = r"■--------------------------------------------------------------------\n(.*?)\n--------------------------------------------------------------------■"
        matches = re.findall(pattern, content, re.DOTALL)
        return matches
    def extract_data_building(building: str):   
        sumary_regex = r"(.+?)\s*[\u3000\s](\d+)件"
        building_regex = r"\((\d+)\)" 
        sumary_match = re.search(sumary_regex, building)
        if sumary_match:
            department = sumary_match.group(1)
            project_count = int(sumary_match.group(2))
            department = re.sub(r"（.*?）", "", department)
        else:
            department, project_count = None, 0
        project_ids = re.findall(building_regex, building)
        return {
            "construction": department,
            "quantity": project_count,
            "details": project_ids
        }
        
    matches = split_raw_buildings(content)
    for match in matches:
        building = extract_data_building(match)
        buildings.append(building)
    return buildings

def run(timeout:int=10,headless:bool=False):
    logger = logging.getLogger("Main")
    touei = Touei(
        username="c0032",
        password="nsk159753",
        headless=headless,
        timeout=timeout,
        logger_name="Touei",
    ) 
    web_access = WebAccess(
        username="hanh0704",
        password="159753",
        headless=headless,
        timeout=timeout,
        logger_name="WebAccess",
    )
    mail_dealer = MailDealer(
        username='vietnamrpa',
        password='nsk159753',
        headless=headless,
        timeout=timeout,
        logger_name="MailDealer",
    )
    if not (mail_dealer.authenticated and touei.authenticated and web_access.authenticated):
        logger.error("❌ Kiểm tra thông tin xác thực")
        return
    mailbox: pd.DataFrame = mail_dealer.mailbox(
        mail_box = MAIL_BOX,
        tab_name = TAB_NAME,
    )
    if mailbox is None:
        return
    # Chuyển cột 日付 thành datetime "%y/%m/%d %H:%M"
    mailbox["日付"] = pd.to_datetime(mailbox["日付"], format="%y/%m/%d %H:%M", errors="coerce")
    # Lọc các mail có cột '件名' bắt đầu bằng "【東栄住宅】 工程表更新のお知らせ"
    mailbox = mailbox[mailbox['件名'].str.startswith("【東栄住宅】 工程表更新のお知らせ", na=False)]
    data = []
    # Duyệt từng Mail
    for ID in mailbox['ID'].to_list():
        content = mail_dealer.read_mail(
            mail_box=MAIL_BOX,
            mail_id=ID,
            tab_name=TAB_NAME,
        )
        constructions:list[dict] = process_schedu_email_content(content)
        # Chỉ lấy các element có key construction nằm trong PROCESS_CONSTRUCTIONS cần xử lí
        constructions:list[dict] = [item for item in constructions if any(keyword in item.get("construction", "") for keyword in PROCESS_CONSTRUCTIONS)]
        # Lấy các constructions_id cần xử lí
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
            for construction in constructions:
                for construction_id in construction.get("details"):
                    future_timeline = executor.submit(touei.get_schedule, construction_id=construction_id, task=TASK)
                    future_information = executor.submit(web_access.get_information, construction_id=construction_id, fields=FIELDS)
                    job_timeline = future_timeline.result()
                    web_access_information = future_information.result()
                    web_access_information = web_access_information.sort_values(by="案件番号",ascending=True).reset_index(drop=True)
                    print(f"web_access_information: {web_access_information}")
                    if job_timeline == None:
                        data.append([None,None,construction_id,None,None,None,False,"KHÔNG LẤY ĐƯỢC THÔNG TIN Ở TOEUI"])
                        continue
                    if web_access_information.empty:
                        data.append([None,None,construction_id,None,None,None,False,"KHÔNG LẤY ĐƯỢC THÔNG TIN Ở WEB ACCESS"])
                        continue
                    if construction.get("construction").startswith("東京施工"):
                        # Nếu địa chỉ trong access (cột 配送先住所) không chứa 1 trong những giá trị này ignore_region thì result ghi: vùng không cần làm-> bot không làm các bước tiếp theo
                        ignore_region = ['甲府市、','富士吉田市、','都留市、',"山梨市、",'大月市、',"韮崎市、","南アルプス市、","北杜市、","甲斐市、","笛吹市、","上野原市、","甲州市、","中央市"]
                        配送先住所:list = web_access_information['配送先住所'].to_list()
                        if not any(region in address for region in ignore_region for address in 配送先住所):
                            if web_access_information.empty:
                                data.append([None,None,construction_id,None,None,None,False,"IGNORE"])
                            else:
                                for index, row in web_access_information.iterrows():
                                    touei_endtime:datetime.datetime = job_timeline.get(index+1).get("end")
                                    web_access_endtime = None
                                    try:
                                        web_access_endtime = datetime.datetime.strptime(row['確定納期'],"%Y/%m/%d")
                                    except Exception as e:
                                        logger.error(e)
                                        pass
                                    data.append([row['案件番号'],row['物件名'],construction_id,touei_endtime.strftime("%Y-%m-%d"),web_access_endtime.strftime("%Y-%m-%d"),0,False,"IGNORE"])
                            continue
                    if construction.get("construction").startswith("神奈川施工"):
                        ignore_region = ['静岡県']
                        配送先住所:list = web_access_information['配送先住所'].to_list()
                        if not any(region in address for region in ignore_region for address in 配送先住所):
                            if web_access_information.empty:
                                data.append([None,None,construction_id,None,None,None,False,"IGNORE"])
                            else:
                                for index, row in web_access_information.iterrows():
                                    touei_endtime:datetime.datetime = job_timeline.get(index+1).get("end")
                                    web_access_endtime = None
                                    try:
                                        web_access_endtime = datetime.datetime.strptime(row['確定納期'],"%y/%m/%d")
                                    except Exception as e:
                                        logger.error(e)
                                    data.append(
                                        [
                                            row['案件番号'],
                                            row['物件名'],
                                            construction_id,touei_endtime.strftime("%Y-%m-%d"),
                                            web_access_endtime.strftime("%Y-%m-%d") if web_access_endtime else None,
                                            0,
                                            False,
                                            "IGNORE"
                                        ]
                                    )
                            continue
                    for index,案件番号 in enumerate(web_access_information['案件番号'].to_list()):
                        result_一括操作 = mail_dealer.一括操作(
                            案件ID=案件番号,
                            このメールと同じ親番号のメールをすべて関連付ける=True,
                        )
                        row = web_access_information.loc[index]
                        try:
                            touei_endtime:datetime.datetime = job_timeline.get(index+1).get("end")
                        except Exception as e:
                            logger.error(e)
                            touei_endtime = None
                        web_access_endtime = None
                        try:
                            web_access_endtime = datetime.datetime.strptime(row['確定納期'],"%y/%m/%d")
                        except ValueError:
                            web_access_endtime = datetime.datetime.strptime(row['確定納期'],"%Y/%m/%d")
                        except Exception as e:
                            logger.error(e)
                        data.append([
                            案件番号,
                            row['物件名'],
                            construction_id,
                            touei_endtime.strftime("%Y-%m-%d") if touei_endtime != None else touei_endtime,
                            web_access_endtime.strftime("%Y-%m-%d") if web_access_endtime != None else web_access_endtime,
                            (web_access_endtime-touei_endtime) if web_access_endtime != None and touei_endtime != None else None,
                            result_一括操作[0],
                            result_一括操作[1],
                        ])
    result = pd.DataFrame(
        columns=['案件番号','物件名','CODE','NOUKI TOEUI','NOUKI WEBACCESS','NOUKI DIFF','RESULT',"NOTE"],
        data=data,
        dtype=object,
    )
    result.drop_duplicates(inplace=True)
   
    del touei
    del web_access
    del mail_dealer

    return result

@shared_task(bind=True,name='Toei Xác Nhận Nouki')
def ToeiXacNhanNouki(self):
    result = run(headless=False)
    # --- Upload to Minio
    excel_buffer = io.BytesIO()
    result.to_excel(excel_buffer, index=False, engine="openpyxl")
    excel_buffer.seek(0)
    result = minio.put_object(
        bucket_name=settings.MINIO_BUCKET,
        object_name=f"DrawingClassic/{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
        data=excel_buffer,
        length=excel_buffer.getbuffer().nbytes,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    return result.object_name
