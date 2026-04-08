import logging
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from core.system_config import sys_config
from core.logger_config import clear_session_log

router = APIRouter(prefix="/api", tags=["logs"])
logger = logging.getLogger(__name__)

@router.get("/logs/current")
async def get_current_log():
    log_file = sys_config.run_log_dir / "current_session.log"
    if not log_file.exists():
        return {"content": "", "lines": 0}
    try:
        with open(log_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        return {"content": content, "lines": content.count('\n')}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/logs/clear")
async def clear_current_log():
    try:
        clear_session_log()
        return {"status": "ok"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
