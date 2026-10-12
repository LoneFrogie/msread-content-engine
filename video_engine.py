"""
MS. READ Content Engine — Video Generation via Google Gemini Omni
Generates short product video clips (about 10 sec) from AI-generated images.
Supports image-to-video (product shots → animated clips) for Reels/TikTok.
"""

import time
import base64
import logging
from io import BytesIO
from pathlib import Path
from typing import Callable, Optional

import requests
from PIL import Image as PILImage

logger = logging.getLogger(__name__)

QUOTA_NOTICE = (
    " — video quota exhausted on the Google AI key. Videos resume when the quota "
    "resets (daily, ~3pm MYT) or after the limit is raised in Google AI Studio."
)


def _is_quota_error(e) -> bool:
    s = str(e)
    return "429" in s or "RESOURCE_EXHAUSTED" in s

# Video model. Google shuts down every Veo 3.1 preview model in the Gemini API on
# 22 Oct 2026 and names gemini-omni-1.1-flash as the replacement. Omni is served
# through the Interactions API, not the models.generate_videos call Veo used, so
# render_clip() calls the REST endpoint directly: the SDK's interactions surface
# is marked experimental and changed its response shape in May 2026.
VIDEO_MODEL = "gemini-omni-1.1-flash"
INTERACTIONS_URL = "https://generativelanguage.googleapis.com/v1beta/interactions"
API_REVISION = "2026-05-20"  # pins the response shape: steps[] -> model_output -> video


def _json_or_raise(r):
    """Return the JSON body, or raise with the HTTP and Google status in the message."""
    if r.status_code == 200:
        return r.json()
    try:
        err = r.json().get("error", {})
    except ValueError:
        err = {}
    raise RuntimeError(f"{r.status_code} {err.get('status', '')}: "
                       f"{err.get('message', r.text[:200])}")


def render_clip(api_key: str, image_bytes: bytes, prompt: str, out_path: Path,
                aspect_ratio: str = "9:16", timeout_s: int = 300,
                on_wait: Optional[Callable[[int], None]] = None) -> bool:
    """
    Animate one still into a clip and write it to out_path.

    Starts a background interaction, polls it every 10 s, and saves the video
    when it completes. Returns True when a file was saved, False on timeout.
    Raises on an API error; the message carries the HTTP and Google status, so
    a quota error reads "429 RESOURCE_EXHAUSTED: ...". on_wait(seconds) is
    called every 30 s while the clip renders.
    """
    if not api_key:
        raise RuntimeError("No Google AI API key for video generation")
    headers = {"x-goog-api-key": api_key, "Content-Type": "application/json",
               "Api-Revision": API_REVISION}
    body = {
        "model": VIDEO_MODEL,
        "background": True,
        "input": [
            {"type": "image", "mime_type": "image/png",
             "data": base64.b64encode(image_bytes).decode()},
            {"type": "text", "text": prompt},
        ],
        "response_format": {"type": "video", "aspect_ratio": aspect_ratio},
        "generation_config": {"video_config": {"task": "image_to_video"}},
    }
    job = _json_or_raise(requests.post(INTERACTIONS_URL, headers=headers,
                                       json=body, timeout=120))
    started, polls = time.time(), 0
    while job.get("status") == "in_progress":
        if time.time() - started > timeout_s:
            return False
        time.sleep(10)
        polls += 1
        if on_wait and polls % 3 == 0:
            on_wait(int(time.time() - started))
        job = _json_or_raise(requests.get(f"{INTERACTIONS_URL}/{job['id']}",
                                          headers=headers, timeout=120))
    if job.get("status") != "completed":
        raise RuntimeError(f"video job {job.get('status')}: "
                           f"{str(job.get('error') or '')[:200]}")
    video = next((c for step in job.get("steps", []) if step.get("type") == "model_output"
                  for c in step.get("content", []) if c.get("type") == "video"), None)
    if not video:
        raise RuntimeError("video job completed but returned no video")
    if video.get("data"):
        out_path.write_bytes(base64.b64decode(video["data"]))
    else:
        d = requests.get(video["uri"], headers={"x-goog-api-key": api_key}, timeout=300)
        d.raise_for_status()
        out_path.write_bytes(d.content)
    return True

# Video prompt templates for different scene types
VIDEO_SCENE_PROMPTS = {
    "product_showcase": (
        "Slow cinematic camera orbit around the garment on display. "
        "Fabric gently swaying. Warm golden studio lighting. "
        "Smooth, premium feel. No text overlays."
    ),
    "lifestyle": (
        "Confident plus-size woman walking naturally through a bright, modern space. "
        "Fabric flows with movement. Warm natural lighting. "
        "Slow motion details of fabric texture. Malaysian urban setting."
    ),
    "detail_closeup": (
        "Extreme close-up showing fabric texture and construction details. "
        "Camera slowly pans across the garment's key design elements. "
        "Soft, diffused lighting revealing material quality."
    ),
    "street_style": (
        "Fashion editorial movement shot. Model walking towards camera on a city street. "
        "Outfit visible in full. Confident stride, natural movement. "
        "Golden hour lighting. Slow motion fabric movement."
    ),
}

# Map image prompt themes to video scene types
def _classify_scene(theme: str) -> str:
    """Map an image theme/scene label to a video scene type."""
    theme_lower = theme.lower()
    if any(kw in theme_lower for kw in ["detail", "close", "flat lay", "texture"]):
        return "detail_closeup"
    if any(kw in theme_lower for kw in ["street", "urban", "city", "outdoor"]):
        return "street_style"
    if any(kw in theme_lower for kw in ["lifestyle", "casual", "daily", "brunch"]):
        return "lifestyle"
    return "product_showcase"


def generate_videos(client, image_dir: Path, output_dir: Path,
                    product_title: str, creative_brief: str,
                    callback: Callable, max_videos: int = 4,
                    api_key: str = None) -> list:
    """
    Generate short video clips from AI-generated product images.

    Takes the best images from the image generation phase and creates
    clips of about 10 seconds, suitable for Reels/TikTok/Stories.

    Args:
        client: Gemini client
        image_dir: Directory containing generated PNG images
        output_dir: Output directory for video files
        product_title: Product name for context
        creative_brief: User's creative direction
        callback: Progress callback
        max_videos: Maximum number of videos to generate (default 4)

    Returns:
        List of dicts with video metadata
    """
    video_dir = output_dir / "videos"
    video_dir.mkdir(parents=True, exist_ok=True)

    # Find available generated images
    images = sorted(image_dir.glob("*.png"))
    if not images:
        callback("status", {
            "phase": "generating_videos",
            "message": "No images found for video generation, skipping."
        })
        return []

    # Select up to max_videos images (spread evenly across available images)
    if len(images) > max_videos:
        step = len(images) / max_videos
        selected = [images[int(i * step)] for i in range(max_videos)]
    else:
        selected = images[:max_videos]

    total = len(selected)
    callback("status", {
        "phase": "generating_videos",
        "message": f"Generating {total} product video clips (about 10 sec each)...",
        "total": total,
        "current": 0,
    })

    generated_videos = []
    quota_exhausted = False

    for i, img_path in enumerate(selected):
        scene_name = img_path.stem.replace("day_", "").replace("sku_", "")
        if quota_exhausted:
            callback("video_done", {
                "index": i, "total": total, "scene": scene_name,
                "filename": None, "success": False,
                "message": f"{scene_name} — skipped (video quota exhausted)"
            })
            continue
        scene_type = _classify_scene(scene_name)
        scene_prompt = VIDEO_SCENE_PROMPTS.get(scene_type, VIDEO_SCENE_PROMPTS["product_showcase"])

        # Build the video prompt
        prompt = (
            f"MS. READ fashion brand. {product_title}. "
            f"In a single continuous shot, no scene cuts. "
            f"Animate this exact image; keep the garment exactly as shown. "
            f"{scene_prompt} "
            f"Premium plus-size fashion brand aesthetic."
        )
        if creative_brief:
            prompt += f" Creative direction: {creative_brief[:100]}."

        filename = f"video_{scene_name}.mp4"
        filepath = video_dir / filename

        callback("video_start", {
            "index": i,
            "total": total,
            "scene": scene_name,
            "message": f"Generating video {i + 1}/{total}: {scene_name}..."
        })

        try:
            # Load the source image
            img = PILImage.open(img_path)
            # Resize if needed (the video model works best with reasonable sizes)
            if max(img.size) > 1024:
                ratio = 1024 / max(img.size)
                img = img.resize(
                    (int(img.width * ratio), int(img.height * ratio)),
                    PILImage.LANCZOS
                )
            buf = BytesIO()
            img.save(buf, format="PNG")
            image_bytes = buf.getvalue()

            # Render the clip (image-to-video). A 429 can be a transient
            # per-minute limit, so back off and retry before giving up.
            def _tick(secs, n=i + 1):
                callback("status", {
                    "phase": "generating_videos",
                    "message": f"Video {n}/{total} rendering... ({secs}s)",
                })

            saved = False
            for attempt in range(3):
                try:
                    saved = render_clip(api_key, image_bytes, prompt, filepath,
                                        on_wait=_tick)
                    break
                except Exception as e:
                    if _is_quota_error(e) and attempt < 2:
                        wait = (15, 45)[attempt]
                        callback("status", {
                            "phase": "generating_videos",
                            "message": f"Video model rate-limited — retrying in {wait}s...",
                        })
                        time.sleep(wait)
                        continue
                    raise

            if not saved:
                callback("video_done", {
                    "index": i, "total": total, "scene": scene_name,
                    "filename": None, "success": False,
                    "message": f"{scene_name} — Timed out after 5 minutes"
                })
                continue

            size_mb = filepath.stat().st_size / (1024 * 1024)
            generated_videos.append({
                "scene": scene_name,
                "filename": filename,
                "size_mb": round(size_mb, 1),
            })
            callback("video_done", {
                "index": i, "total": total, "scene": scene_name,
                "filename": filename, "success": True,
                "message": f"{scene_name} done ({size_mb:.1f} MB)"
            })

        except Exception as e:
            logger.warning(f"Video generation failed for {scene_name}: {e}")
            if _is_quota_error(e):
                quota_exhausted = True
                callback("video_done", {
                    "index": i, "total": total, "scene": scene_name,
                    "filename": None, "success": False,
                    "message": f"{scene_name} — video quota exhausted (skipping the rest)"
                })
                continue
            callback("video_done", {
                "index": i, "total": total, "scene": scene_name,
                "filename": None, "success": False,
                "message": f"{scene_name} — Failed: {str(e)[:120]}"
            })

        # Rate limiting between video generations
        if i < total - 1:
            time.sleep(5)

    done_msg = f"{len(generated_videos)}/{total} videos generated"
    if quota_exhausted:
        done_msg += QUOTA_NOTICE
    callback("status", {
        "phase": "videos_done",
        "message": done_msg,
        "total_videos": len(generated_videos),
        "quota_exhausted": quota_exhausted,
    })

    return generated_videos
