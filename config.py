from pathlib import Path

UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
REFRESH_INTERVAL = 15 * 60  # 15 minutes
