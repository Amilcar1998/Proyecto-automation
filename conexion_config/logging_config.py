import logging
from pathlib import Path
from datetime import datetime

# Module-level flag to ensure handlers are configured only once
_configured = False


def get_logger(name=None, level=logging.INFO):
    """Return a named logger that propagates to a single shared file+console handler.

    All modules will write to the same daily file `main_YYYYMMDD.log` inside `logs/`.
    """
    global _configured

    if name is None:
        name = Path(__file__).stem
    else:
        try:
            name = Path(name).stem
        except Exception:
            name = str(name)

    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Configure root handlers once: one FileHandler (daily file) and one StreamHandler
    if not _configured:
        # Custom formatter to match required format: "YYYY-MM-DD HH:MM:SS,mmm - [MODULE] - LEVEL - message"
        class SimpleFormatter(logging.Formatter):
            def __init__(self):
                super().__init__()
                self.formatter_for_exc = None

            def format(self, record):
                # Timestamp with milliseconds
                ct = self.converter(record.created)
                timestamp = datetime.fromtimestamp(record.created).strftime('%Y-%m-%d %H:%M:%S,%f')[:-3]
                # Module name (padded to same length with spaces)
                name = record.name.ljust(25)
                # Level name
                levelname = record.levelname
                # Message
                message = record.getMessage()
                s = f"{timestamp} - {name} - {levelname} - {message}"
                if record.exc_info:
                    if not self.formatter_for_exc:
                        self.formatter_for_exc = logging.Formatter()
                    s = s + "\n" + self.formatter_for_exc.formatException(record.exc_info)
                return s

        formatter = SimpleFormatter()

        try:
            logs_dir = Path.cwd() / 'logs'
            logs_dir.mkdir(parents=True, exist_ok=True)
            date_str = datetime.now().strftime('%Y%m%d')
            file_path = logs_dir / f"main_{date_str}.log"

            root = logging.getLogger()
            root.setLevel(level)

            # Add file handler to root if not present
            has_file = any(isinstance(h, logging.FileHandler) and Path(getattr(h, 'baseFilename', '')).resolve() == file_path.resolve() for h in root.handlers)
            if not has_file:
                fh = logging.FileHandler(str(file_path), encoding='utf-8')
                fh.setFormatter(formatter)
                fh.setLevel(level)
                root.addHandler(fh)

            # Add stream handler to root if not present
            has_stream = any(isinstance(h, logging.StreamHandler) for h in root.handlers)
            if not has_stream:
                sh = logging.StreamHandler()
                sh.setFormatter(formatter)
                sh.setLevel(level)
                root.addHandler(sh)

        except Exception:
            # If any of the above fails, keep going with basic configuration
            logging.basicConfig(level=level)

        _configured = True

    # Ensure individual loggers do not have their own handlers (avoid duplicates)
    if logger.handlers:
        for h in list(logger.handlers):
            logger.removeHandler(h)

    logger.propagate = True
    return logger
