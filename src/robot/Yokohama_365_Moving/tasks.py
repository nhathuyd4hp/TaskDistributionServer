import subprocess
from pathlib import Path

from celery import shared_task

from src.core.inactive_task import InactiveTask


@shared_task(
    bind=True,
    name="Yokohama 365 Moving",
    base=InactiveTask,
)
def Yokohama(self):
    exe_path = (
        Path(__file__).resolve().parents[2] / "robot" / "Yokohama_365_Moving" / "sharepoint_folder_moving_V1_1.exe"
    )
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()
