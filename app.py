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

    thread = Thread(target=_run_in_thread, args=(session, callback), daemon=True)
    thread.start()

    return {"session_id": session_id}


def _run_in_thread(session: Session, callback):
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

    zip_path = sessions[session_id].output_dir / "MSRead_Content_Engine.zip"
    if not zip_path.exists():
        raise HTTPException(status_code=404, detail="Package not ready yet")

    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=f"MSRead_Content_Engine_{datetime.now().strftime('%Y%m%d')}.zip",
    )


@app.get("/api/download-excel/{session_id}")
async def download_excel(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    excel_path = sessions[session_id].output_dir / "MSRead_Content_Engine.xlsx"
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail="Excel not ready yet")

    return FileResponse(
        excel_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"MSRead_Content_Engine_{datetime.now().strftime('%Y%m%d')}.xlsx",
    )


if __name__ == "__main__":
    print()
    print("  +================================================+")
    print("  |   MS. READ Content Engine                       |")
    print("  |   http://localhost:8001                         |")
    print("  +================================================+")
    print()
    uvicorn.run(app, host="0.0.0.0", port=8001)
