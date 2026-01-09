import subprocess
import sys
import typing
from datetime import datetime
from pathlib import Path
from src.service import ResultService as minio
from celery import shared_task


@shared_task(bind=True, name="Furiwake Osaka")
def FuriwakeOsaka(
    self,
    工場: typing.Literal["大阪工場 製造データ", "栃木工場"],
    日付: datetime | str,
):
    if isinstance(日付, str):
        日付 = datetime.fromisoformat(日付)
    日付: str = f"{日付.month}月{日付.day:02d}日"
    # --- #
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    exe_path = Path(__file__).resolve().parents[2] / "robot" / "FuriwakeOsaka" / "Main.py"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [
                sys.executable,
                str(exe_path),
                "--工場",
                "栃木工場" if 工場 == "栃木工場" else "大阪工場　製造データ",
                "--日付",
                日付,
            ],
            cwd=str(exe_path.parent),
            stdout=f,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
        )
        process.wait()