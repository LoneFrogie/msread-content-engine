"""
MS. READ Content Engine — Higgsfield Generation Adapter (CONTENT +)

CONTENT + uses Higgsfield as the *premium* viral-video generator: Higgsfield's
signature is one-click cinematic "camera-control" motion presets (crash zoom,
dolly-in, 360 orbit, bullet-time, FPV drone, robo-arm, etc.) applied to a still
image — the kind of scroll-stopping motion plain text-to-video can't match.

It is wired as an ADAPTER so the tab works TODAY and upgrades cleanly:

  * If a Higgsfield / fal.ai key is configured (HIGGSFIELD_API_KEY or FAL_KEY),
    clips render through Higgsfield motion-preset models.
  * Otherwise it falls back to the proven Google Veo 3.1 image-to-video call
    (identical to the SKU/Calendar tabs), so nothing is blocked on a key.

Integrates the OFFICIAL Higgsfield REST API (platform.higgsfield.ai, verified
against docs.higgsfield.ai): upload the still -> POST the DoP image-to-video
endpoint (camera motion steered by free-text in the prompt — the documented REST
mechanism) -> poll /requests/{id}/status until "completed" -> download video.url.
Higgsfield is NOT hosted on fal.ai, so this talks to the official host directly.

Credentials (an id + a secret) and endpoint are env-overridable so the route can
be corrected WITHOUT a redeploy:

    HF_KEY               "id:secret" (checked first), OR
    HF_API_KEY + HF_API_SECRET     (also HIGGSFIELD_API_KEY + HIGGSFIELD_API_SECRET)
    HIGGSFIELD_API_BASE  Default "https://platform.higgsfield.ai"
    HIGGSFIELD_DOP_ENDPOINT  Default "/higgsfield-ai/dop/standard"
"""

import os
import time
import logging
from io import BytesIO
from pathlib import Path
from typing import Callable, Optional

import requests
from google.genai import types
from PIL import Image as PILImage

logger = logging.getLogger(__name__)

QUOTA_NOTICE = (
    " — Veo quota exhausted on the Google AI key. Clips resume when the quota "
    "resets (daily, ~3pm MYT) or after the limit is raised in Google AI Studio."
)


def is_quota_error(e) -> bool:
    s = str(e)
    return "429" in s or "RESOURCE_EXHAUSTED" in s

# ── Config (official Higgsfield REST API; all overridable via env) ──
VEO_MODEL = "veo-3.1-fast-generate-preview"  # proven fallback (matches video_engine)
HIGGSFIELD_API_BASE = os.getenv("HIGGSFIELD_API_BASE", "https://platform.higgsfield.ai").rstrip("/")
HIGGSFIELD_DOP_ENDPOINT = os.getenv("HIGGSFIELD_DOP_ENDPOINT", "/higgsfield-ai/dop/standard")


def _higgsfield_auth() -> str:
    """
    Build the 'Authorization: Key {id}:{secret}' header value from env, or "".
    Accepts HF_KEY="id:secret" (checked first), or HF_API_KEY + HF_API_SECRET,
    or HIGGSFIELD_API_KEY + HIGGSFIELD_API_SECRET.
    """
    combined = os.getenv("HF_KEY")
    if combined and ":" in combined:
        return f"Key {combined}"
    key_id = os.getenv("HF_API_KEY") or os.getenv("HIGGSFIELD_API_KEY")
    secret = os.getenv("HF_API_SECRET") or os.getenv("HIGGSFIELD_API_SECRET")
    if key_id and secret:
        return f"Key {key_id}:{secret}"
    return ""


def higgsfield_enabled() -> bool:
    """True when valid Higgsfield credentials (id + secret) are configured."""
    return bool(_higgsfield_auth())


# ── Viral camera-motion presets (Higgsfield's signature) ──
# Each preset carries a `motion` phrase appended to the generation prompt so the
# Veo fallback produces the same *kind* of move Higgsfield's named preset does.
PRESETS = {
    "crash_zoom": {
        "label": "Crash Zoom",
        "motion": "sudden fast crash-zoom push into the subject, punchy and energetic, snappy timing",
        "best_for": "hook openers, reveals, 'wait for it' moments",
    },
    "dolly_in": {
        "label": "Cinematic Dolly-In",
        "motion": "slow smooth cinematic dolly-in toward the subject, shallow depth of field, premium film look",
        "best_for": "elegant reveals, occasion/hero pieces",
    },
    "orbit_360": {
        "label": "360° Orbit",
        "motion": "camera slowly orbits 360 degrees around the subject, full outfit visible, floating gimbal motion",
        "best_for": "full-look showcases, new arrivals",
    },
    "bullet_time": {
        "label": "Bullet Time",
        "motion": "bullet-time frozen-moment effect, camera arcs around the subject while motion holds, dramatic",
        "best_for": "transformation / transition moments",
    },
    "fpv_drone": {
        "label": "FPV Drone",
        "motion": "fast FPV drone fly-through that swoops in and around the subject, dynamic and immersive",
        "best_for": "store/lifestyle scenes, energetic launches",
    },
    "robo_arm": {
        "label": "Robo Arm",
        "motion": "precise robotic-arm camera move, fast mechanical sweep past the subject then settling, ad-grade",
        "best_for": "product-detail glam, premium ads",
    },
    "float": {
        "label": "Float / Levitate",
        "motion": "subject and fabric gently float and drift, dreamy weightless slow motion, soft light",
        "best_for": "fabric/texture beauty shots, mood pieces",
    },
    "handheld_vlog": {
        "label": "Handheld Vlog",
        "motion": "natural handheld vlog-style movement, authentic UGC feel, subtle sway, relatable",
        "best_for": "GRWM, testimonials, day-in-the-life",
    },
}
DEFAULT_PRESET = "dolly_in"


def preset_motion(preset_id: str) -> str:
    return PRESETS.get(preset_id, PRESETS[DEFAULT_PRESET])["motion"]


def _prep_image_bytes(img_path: Path, max_side: int = 1024) -> bytes:
    """Load + downscale a PNG for generation (Veo/Higgsfield prefer <=1024)."""
    img = PILImage.open(img_path)
    if max(img.size) > max_side:
        ratio = max_side / max(img.size)
        img = img.resize((int(img.width * ratio), int(img.height * ratio)), PILImage.LANCZOS)
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ── Premium path: Higgsfield (via fal.ai host or custom endpoint) ──

def _higgsfield_clip(image_bytes: bytes, prompt: str, out_path: Path,
                     timeout: int = 240) -> Optional[Path]:
    """
    Render one motion clip through the OFFICIAL Higgsfield DoP API. Returns the
    saved path, or None on any failure so the caller falls back to Veo.

    Flow (docs.higgsfield.ai): the API takes a public image URL, not raw bytes,
    so first upload the still; then enqueue the DoP image-to-video job (camera
    motion is carried by the free-text prompt — the documented REST mechanism);
    then poll /requests/{id}/status until "completed" and download video.url.
    """
    auth = _higgsfield_auth()
    if not auth:
        return None
    headers = {"Authorization": auth, "Content-Type": "application/json", "Accept": "application/json"}
    try:
        # 1. Get an upload URL, then PUT the bytes (endpoint wants a public URL)
        up = requests.post(f"{HIGGSFIELD_API_BASE}/files/generate-upload-url",
                           json={"content_type": "image/png"}, headers=headers, timeout=60)
        up.raise_for_status()
        up_data = up.json()
        public_url = up_data.get("public_url") or up_data.get("publicUrl")
        upload_url = up_data.get("upload_url") or up_data.get("uploadUrl")
        if not public_url or not upload_url:
            logger.warning("Higgsfield upload-url response missing fields: %s", str(up_data)[:160])
            return None
        put = requests.put(upload_url, data=image_bytes,
                           headers={"Content-Type": "image/png"}, timeout=timeout)
        put.raise_for_status()

        # 2. Enqueue the DoP image-to-video job (motion via free-text prompt)
        body = {"image_url": public_url, "prompt": prompt, "duration": 5, "aspect_ratio": "9:16"}
        sub = requests.post(f"{HIGGSFIELD_API_BASE}{HIGGSFIELD_DOP_ENDPOINT}",
                            json=body, headers=headers, timeout=60)
        sub.raise_for_status()
        job = sub.json()
        request_id = job.get("request_id") or job.get("id")
        status_url = job.get("status_url") or (
            f"{HIGGSFIELD_API_BASE}/requests/{request_id}/status" if request_id else None)
        if not status_url:
            logger.warning("Higgsfield submit missing status_url/request_id: %s", str(job)[:160])
            return None

        # 3. Poll until a terminal state
        waited = 0
        while waited < timeout:
            time.sleep(6)
            waited += 6
            st = requests.get(status_url, headers=headers, timeout=30)
            st.raise_for_status()
            data = st.json()
            status = (data.get("status") or "").lower()
            if status == "completed":
                video = data.get("video") or {}
                video_url = video.get("url") if isinstance(video, dict) else None
                if not video_url:
                    logger.warning("Higgsfield completed but no video.url: %s", str(data)[:160])
                    return None
                vid = requests.get(video_url, timeout=timeout)
                vid.raise_for_status()
                out_path.write_bytes(vid.content)
                return out_path
            if status in ("failed", "nsfw", "canceled", "cancelled", "error"):
                logger.warning("Higgsfield job %s -> %s", request_id, status)
                return None
        logger.warning("Higgsfield job %s timed out after %ss", request_id, timeout)
        return None
    except Exception as e:
        logger.warning("Higgsfield path failed (%s) — falling back to Veo.", str(e)[:160])
        return None


# ── Fallback path: proven Google Veo 3.1 image-to-video ──

def _veo_clip(client, image_bytes: bytes, prompt: str, out_path: Path,
              api_key: str = None, max_polls: int = 30) -> Optional[Path]:
    """The exact working Veo 3.1 image-to-video call used across the app."""
    operation = client.models.generate_videos(
        model=VEO_MODEL,
        source=types.GenerateVideosSource(
            prompt=prompt,
            image=types.Image(image_bytes=image_bytes, mime_type="image/png"),
        ),
        config=types.GenerateVideosConfig(
            number_of_videos=1,
            aspect_ratio="9:16",
        ),
    )
    polls = 0
    while not operation.done and polls < max_polls:
        time.sleep(10)
        operation = client.operations.get(operation)
        polls += 1
    if not operation.done:
        return None
    if not (operation.result and operation.result.generated_videos):
        return None
    video = operation.result.generated_videos[0].video
    if video.video_bytes:
        out_path.write_bytes(video.video_bytes)
    elif video.uri:
        dl = video.uri + ((f"&key={api_key}" if "?" in video.uri else f"?key={api_key}") if api_key else "")
        r = requests.get(dl, timeout=120)
        r.raise_for_status()
        out_path.write_bytes(r.content)
    else:
        return None
    return out_path


def generate_clip(client, image_path: Path, out_path: Path, *,
                  scene_prompt: str, preset_id: str, product_title: str,
                  creative_brief: str, callback: Callable, api_key: str = None) -> dict:
    """
    Render ONE viral motion clip from a still image.

    Tries Higgsfield first (if a key is set), then falls back to Veo. Returns
    {"filename", "size_mb", "engine", "preset"} on success, {"success": False} else.
    """
    preset = PRESETS.get(preset_id, PRESETS[DEFAULT_PRESET])
    prompt = (
        f"MS. READ — Malaysian women's fashion. {product_title}. "
        f"{scene_prompt} Camera motion: {preset['motion']}. "
        f"Vertical 9:16 social video, premium editorial aesthetic, no text overlays."
    )
    if creative_brief:
        prompt += f" Creative direction: {creative_brief[:120]}."

    image_bytes = _prep_image_bytes(image_path)
    engine = None

    # Premium path
    if higgsfield_enabled():
        callback("status", {"phase": "generating_videos",
                            "message": f"Higgsfield — {preset['label']} motion..."})
        if _higgsfield_clip(image_bytes, prompt, out_path):
            engine = "higgsfield"

    # Fallback path (a 429 can be a transient per-minute limit — back off first)
    if engine is None:
        for attempt in range(3):
            try:
                if _veo_clip(client, image_bytes, prompt, out_path, api_key=api_key):
                    engine = "veo"
                break
            except Exception as e:
                if is_quota_error(e):
                    if attempt < 2:
                        wait = (15, 45)[attempt]
                        callback("status", {"phase": "generating_videos",
                                            "message": f"Veo rate-limited — retrying in {wait}s..."})
                        time.sleep(wait)
                        continue
                    logger.warning("Veo quota exhausted for %s", out_path.stem)
                    return {"success": False, "preset": preset_id, "quota_exhausted": True}
                logger.warning("Veo clip failed for %s: %s", out_path.stem, str(e)[:160])
                break

    if engine and out_path.exists():
        size_mb = out_path.stat().st_size / (1024 * 1024)
        return {"filename": out_path.name, "size_mb": round(size_mb, 1),
                "engine": engine, "preset": preset_id}
    return {"success": False, "preset": preset_id}
