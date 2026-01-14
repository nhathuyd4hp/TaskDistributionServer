# === logging_setup.py ===
import sys
import logging
import os
import datetime

def setup_logging():
    # Create Logs folder if not exists
    logs_folder = os.path.join(os.getcwd(), "Logs")
    os.makedirs(logs_folder, exist_ok=True)

    # Generate timestamped log file
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    log_filename = os.path.join(logs_folder, f"「横浜_確認リスト」{timestamp}.log")

    # Setup root logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # --- Clear old handlers (important if re-run multiple times) ---
    if logger.hasHandlers():
        logger.handlers.clear()

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    # File Handler
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.INFO)

    # Formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    # Add handlers
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
