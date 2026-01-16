# === logging_setup.py ===
import datetime
import logging
import os
import sys


def setup_logging():
    # Create Logs folder if not exists
    logs_folder = os.path.join(os.getcwd(), "Logs")
    os.makedirs(logs_folder, exist_ok=True)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    log_filename = os.path.join(logs_folder, f"Zenbubot_{timestamp}.log")

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)  # ✅ ENABLE DEBUG LOGGING FOR YOUR APP

    # ⛔ Suppress noisy external logs
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger("msal").setLevel(logging.WARNING)
    logging.getLogger("selenium").setLevel(logging.WARNING)
    logging.getLogger("tensorflow").setLevel(logging.WARNING)
    logging.getLogger("asyncio").setLevel(logging.WARNING)
    logging.getLogger("matplotlib").setLevel(logging.WARNING)

    if logger.hasHandlers():
        logger.handlers.clear()

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler(log_filename, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
