from __future__ import annotations
import shutil
from pathlib import Path
from uuid import uuid4
from fastapi import UploadFile
from ..core.config import get_settings

settings = get_settings()


def save_upload(file: UploadFile, prefix: str) -> Path:
    extension = Path(file.filename or "").suffix or ".bin"
    destination = settings.media_dir / f"{prefix}-{uuid4().hex}{extension}"
    with destination.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return destination
