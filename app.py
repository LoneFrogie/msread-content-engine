"""
MS. READ Content Engine — Web App

Run:
    pip install -r requirements.txt
    cp .env.example .env        # Add your GOOGLE_AI_API_KEY
    python app.py

Opens at http://localhost:8001
"""

import os
import sys
import json
import asyncio
import shutil
import base64
from io import BytesIO
from pathlib import Path
from uuid import uuid4
from threading import Thread
from datetime import datetime
from dataclasses import dataclass, field

from dotenv import load_dotenv
load_dotenv()

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
import uvicorn

from engine import run_pipeline
from sku_engine import run_sku_pipeline

app = FastAPI(title="MS. READ Content Engine")

TEMPLATE_DIR = Path(__file__).parent / "templates"
OUTPUT_BASE = Path("/tmp/msread_content_engine")
OUTPUT_BASE.mkdir(parents=True, exist_ok=True)

GOOGLE_AI_API_KEY = os.getenv("GOOGLE_AI_API_KEY", "")


# ── Session management ──

@dataclass
class Session:
    session_id: str
    creative_brief: str
    mode: str = "calendar"  # "calendar" or "sku"
    product_url: str = ""
    status: str = "pending"
    events: list = field(default_factory=list)
    output_dir: Path = None
    created_at: datetime = field(default_factory=datetime.now)
    error: str = None


sessions: dict[str, Session] = {}


def cleanup_old_sessions():
    """Remove sessions older than 2 hours."""
    now = datetime.now()
    expired = [
        sid for sid, s in sessions.items()
        if (now - s.created_at).total_seconds() > 7200
    ]
    for sid in expired:
        s = sessions.pop(sid, None)
        if s and s.output_dir and s.output_dir.exists():
            shutil.rmtree(s.output_dir, ignore_errors=True)


class GenerateRequest(BaseModel):
    creative_brief: str


class GenerateSkuRequest(BaseModel):
    product_url: str
    creative_brief: str = ""
    avatar_images: list[str] = []  # base64-encoded PNG images


@app.get("/")
async def index():
    return FileResponse(TEMPLATE_DIR / "index.html")


@app.post("/api/generate")
async def start_generation(req: GenerateRequest):
    if not GOOGLE_AI_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="GOOGLE_AI_API_KEY not set. Copy .env.example to .env and add your key."
        )

    if not req.creative_brief.strip():
        raise HTTPException(status_code=400, detail="Creative brief cannot be empty")

    cleanup_old_sessions()

    session_id = uuid4().hex[:12]
    output_dir = OUTPUT_BASE / session_id
    output_dir.mkdir(parents=True, exist_ok=True)

    session = Session(
        session_id=session_id,
        creative_brief=req.creative_brief.strip(),
        mode="calendar",
        status="running",
        output_dir=output_dir,
    )
    sessions[session_id] = session

    def callback(event_type, data):
        event = {"type": event_type, "timestamp": datetime.now().isoformat(), **data}
        session.events.append(event)
        if event_type == "error":
            session.status = "error"
            session.error = data.get("message", "Unknown error")
        elif data.get("phase") == "done":
            session.status = "done"

    thread = Thread(target=_run_calendar_thread, args=(session, callback), daemon=True)
    thread.start()

    return {"session_id": session_id, "mode": "calendar"}


@app.post("/api/generate-sku")
async def start_sku_generation(req: GenerateSkuRequest):
    if not GOOGLE_AI_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="GOOGLE_AI_API_KEY not set. Copy .env.example to .env and add your key."
        )

    if not req.product_url.strip():
        raise HTTPException(status_code=400, detail="Product URL cannot be empty")

    cleanup_old_sessions()

    session_id = uuid4().hex[:12]
    output_dir = OUTPUT_BASE / session_id
    output_dir.mkdir(parents=True, exist_ok=True)

    # Decode avatar images from base64
    avatar_images_bytes = []
    for b64 in req.avatar_images[:3]:  # max 3 avatars
        try:
            # Strip data URI prefix if present
            if "," in b64:
                b64 = b64.split(",", 1)[1]
            avatar_images_bytes.append(base64.b64decode(b64))
        except Exception:
            pass

    session = Session(
        session_id=session_id,
        creative_brief=req.creative_brief.strip(),
        mode="sku",
        product_url=req.product_url.strip(),
        status="running",
        output_dir=output_dir,
    )
    sessions[session_id] = session

    def callback(event_type, data):
        event = {"type": event_type, "timestamp": datetime.now().isoformat(), **data}
        session.events.append(event)
        if event_type == "error":
            session.status = "error"
            session.error = data.get("message", "Unknown error")
        elif data.get("phase") == "done":
            session.status = "done"

    thread = Thread(target=_run_sku_thread, args=(session, callback, avatar_images_bytes), daemon=True)
    thread.start()

    return {"session_id": session_id, "mode": "sku"}


def _run_calendar_thread(session: Session, callback):
    try:
        run_pipeline(
            api_key=GOOGLE_AI_API_KEY,
            creative_brief=session.creative_brief,
            output_dir=session.output_dir,
            callback=callback,
        )
    except Exception as e:
        if session.status != "error":
            session.status = "error"
            session.error = str(e)
            session.events.append({
                "type": "error",
                "message": str(e),
                "timestamp": datetime.now().isoformat(),
            })


def _run_sku_thread(session: Session, callback, avatar_images: list = None):
    try:
        run_sku_pipeline(
            api_key=GOOGLE_AI_API_KEY,
            product_url=session.product_url,
            creative_brief=session.creative_brief,
            output_dir=session.output_dir,
            callback=callback,
            avatar_images=avatar_images,
        )
    except Exception as e:
        if session.status != "error":
            session.status = "error"
            session.error = str(e)
            session.events.append({
                "type": "error",
                "message": str(e),
                "timestamp": datetime.now().isoformat(),
            })


@app.get("/api/progress/{session_id}")
async def progress_stream(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    async def event_generator():
        last_idx = 0
        while True:
            session = sessions.get(session_id)
            if not session:
                break

            new_events = session.events[last_idx:]
            for event in new_events:
                yield f"data: {json.dumps(event)}\n\n"
            last_idx = len(session.events)

            if session.status in ("done", "error"):
                yield f"data: {json.dumps({'type': 'stream_end', 'status': session.status})}\n\n"
                break

            await asyncio.sleep(0.5)

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive", "X-Accel-Buffering": "no"},
    )


@app.get("/api/images/{session_id}/{filename}")
async def serve_image(session_id: str, filename: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]
    for subdir in ["images", "thumbnails"]:
        path = session.output_dir / subdir / filename
        if path.exists():
            return FileResponse(path, media_type="image/png")

    raise HTTPException(status_code=404, detail="Image not found")


@app.get("/api/download/{session_id}")
async def download_zip(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]
    # Find the zip file (could be calendar or sku)
    zip_files = list(session.output_dir.glob("*.zip"))
    if not zip_files:
        raise HTTPException(status_code=404, detail="Package not ready yet")

    zip_path = zip_files[0]
    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=f"{zip_path.stem}_{datetime.now().strftime('%Y%m%d')}.zip",
    )


@app.get("/api/download-excel/{session_id}")
async def download_excel(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]
    # Find the xlsx file (could be calendar or sku)
    xlsx_files = list(session.output_dir.glob("*.xlsx"))
    if not xlsx_files:
        raise HTTPException(status_code=404, detail="Excel not ready yet")

    excel_path = xlsx_files[0]
    return FileResponse(
        excel_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{excel_path.stem}_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


if __name__ == "__main__":
    print()
    print("  +================================================+")
    print("  |   MS. READ Content Engine                       |")
    print("  |   http://localhost:8001                         |")
    print("  +================================================+")
    print()
    uvicorn.run(app, host="0.0.0.0", port=8001)
