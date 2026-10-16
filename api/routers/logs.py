import re
import logging
from typing import Optional
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from core.system_config import sys_config
from core.logger_config import clear_session_log, get_log_level, set_log_level

router = APIRouter(prefix="/api", tags=["logs"])
logger = logging.getLogger(__name__)

LOG_HEADER_REGEX = re.compile(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \|\s*([A-Z]+)\s*\|')

LEVEL_HIERARCHY = {
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}


class LogLevelRequest(BaseModel):
    level: str


@router.get("/logs/level")
async def get_current_log_level():
    return {"level": get_log_level()}


@router.post("/logs/level")
async def update_log_level(payload: LogLevelRequest):
    req_level = payload.level.strip().upper() if payload.level else ""
    if req_level not in ("DEBUG", "INFO", "WARNING", "ERROR"):
        return JSONResponse(
            status_code=400,
            content={"error": f"Invalid log level: {payload.level}", "detail": f"Invalid log level: {payload.level}"}
        )
    try:
        set_log_level(req_level)
        return {"status": "ok", "level": req_level}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@router.get("/logs/current")
async def get_current_log(level: Optional[str] = None):
    log_file = sys_config.run_log_dir / "current_session.log"
    if not log_file.exists():
        return {"content": "", "lines": 0}
    try:
        with open(log_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()

        if not level or not level.strip():
            lines = content.splitlines(keepends=True)
            return {"content": content, "lines": len(lines)}

        req_level = level.strip().upper()
        if req_level == "WARN":
            req_level = "WARNING"

        if req_level not in LEVEL_HIERARCHY:
            return JSONResponse(
                status_code=400,
                content={"error": f"Invalid log level: {level}", "detail": f"Invalid log level: {level}"}
            )

        if req_level == "DEBUG":
            lines = content.splitlines(keepends=True)
            return {"content": content, "lines": len(lines)}

        min_val = LEVEL_HIERARCHY[req_level]
        lines = content.splitlines(keepends=True)
        filtered_lines = []
        include_current = True

        for line in lines:
            match = LOG_HEADER_REGEX.match(line)
            if match:
                line_level_str = match.group(1).upper()
                val = LEVEL_HIERARCHY.get(line_level_str, logging.INFO)
                include_current = (val >= min_val)
            if include_current:
                filtered_lines.append(line)

        return {"content": "".join(filtered_lines), "lines": len(filtered_lines)}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@router.post("/logs/clear")
async def clear_current_log():
    try:
        clear_session_log()
        return {"status": "ok"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
