import shutil
from celery import shared_task
import subprocess
from pathlib import Path

@shared_task(bind=True,name="Yokohama Noukikakunin List")
def YokohamaNoukikakuninList(self):
    exe_path = Path(__file__).resolve().parents[2] / "robot" / "YokohamaNoukikakuninList" / "YokohamaKakunin_V2.0.exe"
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()

    log_source = cwd_path / "Access_token_log"
    logs_folder = cwd_path / "Logs"
    access_token_folder = cwd_path / "Access_token"

    try:
        log_source.unlink()
    except Exception:
        pass

    # xóa folder (nếu có)
    for path in (logs_folder, access_token_folder):
        shutil.rmtree(path, ignore_errors=True)