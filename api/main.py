import logging
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path

# 1. Initialize System Config & Logging FIRST
from core.system_config import sys_config
from core.logger_config import setup_logging
from core.database import db_manager

setup_logging(log_dir=sys_config.run_log_dir)

logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    sys_config.temp_uploads_dir.mkdir(parents=True, exist_ok=True)
    db_manager.init_db()
    yield


# 2. Create FastAPI App
app = FastAPI(title="Giyo Invoice API", lifespan=lifespan)

# Fix Windows MIME type registry issues for Javascript modules
import mimetypes
mimetypes.add_type("application/javascript", ".js")
mimetypes.add_type("text/css", ".css")

# Mount frontend
app.mount("/frontend", StaticFiles(directory=str(sys_config.frontend_dir), html=True), name="frontend")

# 4. Include Modular Routers
from api.routers import blueprint, upload, generate, history, templates, logs, google_sheets

app.include_router(blueprint.router)
app.include_router(upload.router)
app.include_router(generate.router)
app.include_router(history.router)
app.include_router(templates.router)
app.include_router(logs.router)
app.include_router(google_sheets.router)

# 5. Base Routes
@app.get("/")
def redirect_to_frontend():
    return RedirectResponse(url="/frontend/")

@app.get("/api/health")
async def health_check():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=sys_config.api_port)
