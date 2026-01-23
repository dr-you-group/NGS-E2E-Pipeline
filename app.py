import config
import logging
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from database import init_db
from routers import reports, upload, downloads, static

config.setup_logging()
logger = logging.getLogger("app")

@asynccontextmanager
async def lifespan(app: FastAPI):
    # 디렉토리 확인 및 생성
    if not config.JSON_DIR.exists():
        config.JSON_DIR.mkdir(parents=True, exist_ok=True)

    if not config.TMP_DIR.exists():
        config.TMP_DIR.mkdir(parents=True, exist_ok=True)
        logger.info(f"임시 파일 디렉토리 생성: {config.TMP_DIR}")

    # DB 초기화
    init_db()

    yield

    # 앱 종료 시 실행할 로직이 있다면 여기에 작성

app = FastAPI(lifespan=lifespan)

app.mount("/static", StaticFiles(directory=config.STATIC_DIR), name="static")

# Include Routers
app.include_router(static.router)
app.include_router(reports.router)
app.include_router(upload.router)
app.include_router(downloads.router)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=1234, reload=True, workers=1)
