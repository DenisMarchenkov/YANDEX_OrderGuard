import logging
import os
from settings import BASE_DIR

LOG_FILE = os.path.join(BASE_DIR, "main.log")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
