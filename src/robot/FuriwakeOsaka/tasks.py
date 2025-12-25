import shutil
from pathlib import Path
import subprocess
from celery import shared_task
from celery import shared_task

@shared_task(bind=True,name="Furiwake Osaka")
def FuriwakeOsaka(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "FuriwakeOsaka" / "「大阪・栃木」振り分け_V1.5.exe"
    cwd_path = exe_path.parent
    
    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen(
            [str(exe_path)], 
            cwd=str(cwd_path), 
            stdout=f, 
            stderr=subprocess.STDOUT, 
            text=True
        )
        process.wait()