import logging
from colorama import Fore, Style
from .config import LOG_FILE

class ColourFormatter(logging.Formatter):
    COLOURS = {
        logging.INFO: Fore.GREEN,
        logging.WARNING: Fore.YELLOW,
        logging.ERROR: Fore.RED,
        logging.CRITICAL: Fore.RED + Style.BRIGHT,
        logging.DEBUG: Fore.WHITE,
    }
    RESET = Style.RESET_ALL

    def format(self, record):
        msg = super().format(record)
        colour = self.COLOURS.get(record.levelno, self.RESET)
        return f"{colour}{msg}{self.RESET}"

def setup_logger():
    # File handler (no colour)
    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
    file_handler.setFormatter(logging.Formatter("%(asctime)s | %(message)s", "%Y-%m-%d %H:%M:%S"))

    # Console handler (colour)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(ColourFormatter("%(asctime)s | %(message)s", "%Y-%m-%d %H:%M:%S"))

    logging.basicConfig(level=logging.INFO, handlers=[file_handler, console_handler])

def log(msg: str):
    logging.info(msg)