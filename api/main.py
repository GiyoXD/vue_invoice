import logging
from fastapi import FastAPI
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path

# 1. Initialize System Config & Logging FIRST
from core.system_config import sys_config
from core.logger_config import setup_logging

setup_logging(log_dir=sys_config.run_log_dir)
logger = logging.getLogger(__name__)

# 2. Initialize Database
from core.database import db_manager
db_manager.init_db()

# 3. Create FastAPI App
app = FastAPI(title="Giyo Invoice API")

# Mount frontend
app.mount("/frontend", StaticFiles(directory=str(sys_config.frontend_dir), html=True), name="frontend")
app.mount("/static", StaticFiles(directory=str(sys_config.frontend_dir), html=True), name="static")

# 4. Include Modular Routers
from api.routers import blueprint, upload, generate, history, templates, logs

app.include_router(blueprint.router)
app.include_router(upload.router)
app.include_router(generate.router)
app.include_router(history.router)
app.include_router(templates.router)
app.include_router(logs.router)

# 5. Base Routes
@app.get("/")
def redirect_to_frontend():
    return RedirectResponse(url="/frontend/")

@app.get("/api/health")
async def health_check():
    return {"status": "ok"}

# 6. Global Setup
sys_config.temp_uploads_dir.mkdir(parents=True, exist_ok=True)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
