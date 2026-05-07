import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

MAX_CONTENT_BYTES = 100 * 1024 * 1024  # 100 MB
JOB_TTL_SECONDS = 1800                  # 30분 후 자동 삭제
SECRET_KEY = os.environ.get("FLASK_SECRET", "pi-review-dev-secret")
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", 5001))
