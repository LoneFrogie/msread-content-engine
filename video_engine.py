"""
MS. READ Content Engine — Video Generation via Google Veo
Generates short product video clips (5-8 sec) from AI-generated images.
Supports image-to-video (product shots → animated clips) for Reels/TikTok.
"""

import time
import logging
from io import BytesIO
from pathlib import Path
from typing import Callable, Optional

import requests
from google.genai import types
from PIL import Image as PILImage

logger = logging.getLogger(__name__)

# Veo model for image-to-video. Google retired veo-2.0-generate-001 from the
# Gemini API (only Veo 3.1 preview models remain: standard / fast / lite).
# "fast" balances speed + cost for short social clips. If this key loses access
# to it, list models at /v1beta/models and pick an available veo-* model.
VEO_MODEL = "veo-3.1-fast-generate-preview"

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
    5-8 second video clips suitable for Reels/TikTok/Stories.

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
        "message": f"Generating {total} product video clips (5-8 sec each)...",
        "total": total,
        "current": 0,
    })

    generated_videos = []

    for i, img_path in enumerate(selected):
        scene_name = img_path.stem.replace("day_", "").replace("sku_", "")
        scene_type = _classify_scene(scene_name)
        scene_prompt = VIDEO_SCENE_PROMPTS.get(scene_type, VIDEO_SCENE_PROMPTS["product_showcase"])

        # Build the video prompt
        prompt = (
            f"MS. READ fashion brand. {product_title}. "
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
            # Resize if needed (Veo works best with reasonable sizes)
            if max(img.size) > 1024:
                ratio = 1024 / max(img.size)
                img = img.resize(
                    (int(img.width * ratio), int(img.height * ratio)),
                    PILImage.LANCZOS
                )
            buf = BytesIO()
            img.save(buf, format="PNG")
            image_bytes = buf.getvalue()

            # Submit video generation job (image-to-video)
            operation = client.models.generate_videos(
                model=VEO_MODEL,
                source=types.GenerateVideosSource(
                    prompt=prompt,
                    image=types.Image(
                        image_bytes=image_bytes,
                        mime_type="image/png",
                    ),
                ),
                config=types.GenerateVideosConfig(
                    number_of_videos=1,
                    aspect_ratio="9:16",  # Vertical for Reels/TikTok
                ),
            )

            # Poll for completion (videos take 30-120 sec)
            poll_count = 0
            max_polls = 30  # 5 minutes max
            while not operation.done and poll_count < max_polls:
                time.sleep(10)
                operation = client.operations.get(operation)
                poll_count += 1

                if poll_count % 3 == 0:
                    callback("status", {
                        "phase": "generating_videos",
                        "message": f"Video {i + 1}/{total} rendering... ({poll_count * 10}s)",
                    })

            if not operation.done:
                callback("video_done", {
                    "index": i, "total": total, "scene": scene_name,
                    "filename": None, "success": False,
                    "message": f"{scene_name} — Timed out after 5 minutes"
                })
                continue

            # Save the generated video
            if (operation.result and
                    operation.result.generated_videos and
                    len(operation.result.generated_videos) > 0):
                video = operation.result.generated_videos[0].video

                # Handle both inline bytes and remote URI
                if video.video_bytes:
                    with open(filepath, "wb") as f:
                        f.write(video.video_bytes)
                elif video.uri:
                    download_url = video.uri
                    if api_key and "?" in download_url:
                        download_url += f"&key={api_key}"
                    elif api_key:
                        download_url += f"?key={api_key}"
                    resp = requests.get(download_url, timeout=120)
                    resp.raise_for_status()
                    with open(filepath, "wb") as f:
                        f.write(resp.content)
                else:
                    raise ValueError("Video has no bytes or URI")

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
            else:
                callback("video_done", {
                    "index": i, "total": total, "scene": scene_name,
                    "filename": None, "success": False,
                    "message": f"{scene_name} — No video returned"
                })

        except Exception as e:
            logger.warning(f"Video generation failed for {scene_name}: {e}")
            callback("video_done", {
                "index": i, "total": total, "scene": scene_name,
                "filename": None, "success": False,
                "message": f"{scene_name} — Failed: {str(e)[:120]}"
            })

        # Rate limiting between video generations
        if i < total - 1:
            time.sleep(5)

    callback("status", {
        "phase": "videos_done",
        "message": f"{len(generated_videos)}/{total} videos generated",
        "total_videos": len(generated_videos),
    })

    return generated_videos
