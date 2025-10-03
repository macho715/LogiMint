"""로깅 유틸리티. Logging utilities."""

from __future__ import annotations

import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from config.settings import HVDCSettings, ensure_directory


class JsonFormatter(logging.Formatter):
    """JSON 포맷터. JSON formatter."""

    def format(self, record: logging.LogRecord) -> str:  # noqa: D401 (Docstring inherits)
        payload = {
            "level": record.levelname,
            "name": record.name,
            "message": record.getMessage(),
            "time": self.formatTime(record, self.datefmt),
        }
        if record.exc_info:
            payload["exception"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False)


def configure_logger(name: str, settings: HVDCSettings) -> logging.Logger:
    """로거 구성 헬퍼. Configure logger helper."""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    ensure_directory(settings.log_directory)
    log_file = Path(settings.log_directory) / f"{name}.log"

    handler = RotatingFileHandler(
        log_file,
        maxBytes=settings.log_rotation_mb * 1024 * 1024,
        backupCount=settings.log_retention,
        encoding="utf-8",
    )
    formatter: logging.Formatter
    if settings.log_json:
        formatter = JsonFormatter()
    else:
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    logger.propagate = False
    return logger
