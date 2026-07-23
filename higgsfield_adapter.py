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

Everything about the remote call is env-overridable (base URL, model slug),
so the exact Higgsfield endpoint can be corrected WITHOUT a redeploy:

    HIGGSFIELD_API_KEY   Higgsfield/fal key. Presence flips on the premium path.
    FAL_KEY              Alternative key name (fal.ai hosts Higgsfield models).
    HIGGSFIELD_API_BASE  Default "https://fal.run"
    HIGGSFIELD_MODEL     Default "fal-ai/higgsfield/dop-i2v"  (image->video, motion presets)
"""

import os
import time
import base64
import logging
from io import BytesIO
from pathlib import Path
from typing import Callable, Optional

import requests
from google.genai import types
from PIL import Image as PILImage

logger = logging.getLogger(__name__)

# ── Config (all overridable via env — no redeploy needed to correct) ──
VEO_MODEL = "veo-3.1-fast-generate-preview"  # proven fallback (matches video_engine)
HIGGSFIELD_API_BASE = os.getenv("HIGGSFIELD_API_BASE", "https://fal.run").rstrip("/")
HIGGSFIELD_MODEL = os.getenv("HIGGSFIELD_MODEL", "fal-ai/higgsfield/dop-i2v")


def higgsfield_key() -> str:
    """Return the configured Higgsfield/fal key, or empty string."""
    return os.getenv("HIGGSFIELD_API_KEY") or os.getenv("FAL_KEY") or ""


def higgsfield_enabled() -> bool:
    """True when a premium generation key is configured."""
    return bool(higgsfield_key())


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

def _higgsfield_clip(image_bytes: bytes, prompt: str, preset_id: str,
                     out_path: Path, timeout: int = 240) -> Optional[Path]:
    """
    Render one motion clip through Higgsfield. Returns the saved path or None.

    Any failure returns None so the caller falls back to Veo — the feature is
    never blocked by a misconfigured premium endpoint.
    """
    key = higgsfield_key()
    if not key:
        return None
    try:
        data_uri = "data:image/png;base64," + base64.b64encode(image_bytes).decode()
        url = f"{HIGGSFIELD_API_BASE}/{HIGGSFIELD_MODEL}"
        payload = {
            "prompt": prompt,
            "image_url": data_uri,
            "motion": preset_id,           # Higgsfield preset id
            "aspect_ratio": "9:16",
            "duration": 5,
        }
        headers = {
            "Authorization": f"Key {key}",   # fal.ai auth scheme
            "Content-Type": "application/json",
        }
        resp = requests.post(url, json=payload, headers=headers, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()

        # Extract a video URL from common response shapes
        video_url = (
            (data.get("video") or {}).get("url")
            if isinstance(data.get("video"), dict) else data.get("video")
        ) or (
            (data.get("videos") or [{}])[0].get("url")
            if isinstance(data.get("videos"), list) and data["videos"] else None
        ) or data.get("url")

        if not video_url:
            logger.warning("Higgsfield returned no video url: %s", str(data)[:200])
            return None

        vid = requests.get(video_url, timeout=timeout)
        vid.raise_for_status()
        out_path.write_bytes(vid.content)
        return out_path
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
        if _higgsfield_clip(image_bytes, prompt, preset_id, out_path):
            engine = "higgsfield"

    # Fallback path
    if engine is None:
        try:
            if _veo_clip(client, image_bytes, prompt, out_path, api_key=api_key):
                engine = "veo"
        except Exception as e:
            logger.warning("Veo clip failed for %s: %s", out_path.stem, str(e)[:160])

    if engine and out_path.exists():
        size_mb = out_path.stat().st_size / (1024 * 1024)
        return {"filename": out_path.name, "size_mb": round(size_mb, 1),
                "engine": engine, "preset": preset_id}
    return {"success": False, "preset": preset_id}
