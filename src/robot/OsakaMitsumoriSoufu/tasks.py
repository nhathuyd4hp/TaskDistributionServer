import shutil
import subprocess
from pathlib import Path

from celery import shared_task


@shared_task(bind=True, name="Osaka Mitsumori Soufu")
def OsakaMitsumoriSoufu(self):
    exe_path = (
        Path(__file__).resolve().parents[2] / "robot" / "OsakaMitsumoriSoufu" / "「大阪・インド」見積書送付_V1.5.exe"
    )
    cwd_path = exe_path.parent

    log_dir = Path(__file__).resolve().parents[3] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{self.request.id}.log"

    with open(log_file, "w", encoding="utf-8", errors="ignore") as f:
        process = subprocess.Popen([str(exe_path)], cwd=str(cwd_path), stdout=f, stderr=subprocess.STDOUT, text=True)
        process.wait()
    # Clean
    Access_token_log = cwd_path / "Access_token_log"
    try:
        Access_token_log.unlink()
    except Exception:
        pass
    logs_folder = cwd_path / "Logs"
    access_token_folder = cwd_path / "Access_token"
    ankens_folder = cwd_path / "Ankens"
    for path in (logs_folder, access_token_folder, ankens_folder):
        shutil.rmtree(path, ignore_errors=True)
