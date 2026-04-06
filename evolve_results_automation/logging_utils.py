import logging
from .config import current_log_path


class _FlushHandler(logging.FileHandler):
    """FileHandler that flushes after every log entry so nothing is lost on crash."""
    def emit(self, record):
        super().emit(record)
        self.flush()


def setup_logger():
    root = logging.getLogger()
    root.setLevel(logging.INFO)

    # Avoid adding duplicate handlers on repeated calls
    if any(isinstance(h, _FlushHandler) for h in root.handlers):
        return

    log_file = current_log_path()
    file_handler = _FlushHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(logging.Formatter("%(asctime)s | %(message)s", "%Y-%m-%d %H:%M:%S"))
    root.addHandler(file_handler)