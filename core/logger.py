import logging
import sys
from pathlib import Path

LOG_FILE = Path(__file__).resolve().parent.parent / "gui1c.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)

logger = logging.getLogger("gui1c")
