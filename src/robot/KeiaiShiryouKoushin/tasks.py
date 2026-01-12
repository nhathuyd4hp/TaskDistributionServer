import shutil
import subprocess
from pathlib import Path

from celery import shared_task

from src.core.inactive_task import InactiveTask


@shared_task(bind=True, name="Keiai Shiryou Koushin", base=InactiveTask)
def keiai_shiryou_koushin(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "KeiaiShiryouKoushin" / "KISTAR_資料更新_V3.1.exe"
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()

    logs_folder = cwd_path / "logs"
    access_token_folder = cwd_path / "Ankens"

    for path in (logs_folder, access_token_folder):
        shutil.rmtree(path, ignore_errors=True)
