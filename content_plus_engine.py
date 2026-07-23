"""
MS. READ Content Engine — CONTENT + (Trend-to-Reel Viral Studio)

The premium tab. Turns a *current* trend (Malaysia-local or global) + one product
into a launch-ready viral asset:

  1. Fetch the product from msreadshop.com               (reuse: sku_engine.fetch_product)
  2. Write a seasoned-marketer STRATEGY BRIEF            (Claude Opus, Gemini fallback)
     — hook, shotlist, on-screen text, MY-localised caption + hashtags, CTA,
       best post time, WHY it beats competitors, KPI to watch
  3. Generate trend-styled, product-faithful hero images (reuse: sku_engine.generate_sku_images)
  4. Animate each shot into a viral motion clip          (higgsfield_adapter — Higgsfield → Veo)
  5. Package brief + images + clips, and serve the brief JSON to the UI

Works today on the existing Google key; upgrades to Higgsfield motion + Claude
copy the moment those keys are added. Nothing here hard-fails on a missing key.
"""

import os
import json
import time
import zipfile
import logging
from pathlib import Path
from typing import Callable

from google import genai

from sku_engine import fetch_product, generate_sku_images
from higgsfield_adapter import generate_clip, PRESETS, DEFAULT_PRESET, higgsfield_enabled
from trend_catalog import get_trend

logger = logging.getLogger(__name__)

BRAND = "MS. READ"
# Skill default (claude-api). Env-overridable; Gemini fallback means a wrong/
# unavailable id degrades gracefully instead of hard-failing.
CLAUDE_MODEL = os.getenv("CONTENT_PLUS_CLAUDE_MODEL", "claude-opus-4-6")
GEMINI_TEXT_MODEL = "gemini-2.5-flash"

_PRESET_MENU = "; ".join(f"{k} = {v['label']} ({v['best_for']})" for k, v in PRESETS.items())


# ─────────────────────────────────────────────────────────────
# Strategy brief (the marketing brain)
# ─────────────────────────────────────────────────────────────

STRATEGY_PROMPT = """You are a seasoned digital-marketing and social-content strategist for {brand}, an elegant, size-inclusive Malaysian women's fashion label (modest workwear, casual & occasion wear; boutiques in shopping malls across Malaysia + online at msreadshop.com). Core customer: Malaysian women ~25–45, style-aware but value-conscious, many wear modestwear/hijab.

Your job: design ONE scroll-stopping short-video concept that rides the trend below and makes {brand} look more desirable than local competitors (Poplook, dUCk, Nafeesa, Zalia, Jovian) and fast fashion (Shein, Lovito, Uniqlo).

TREND TO RIDE:
- Name: {trend_name}
- Scope/moment: {trend_scope} · {trend_moment}
- Format: {trend_format}
- Platform: {trend_platform}
- Edge: {trend_why}

PRODUCT:
- Title: {product_title}
- Type: {product_type}
- Price: {product_price}
- Notes: {product_desc}

EXTRA CREATIVE DIRECTION (optional): {creative_brief}

AVAILABLE CAMERA-MOTION PRESETS (pick the best fit per shot):
{preset_menu}

Write the concept as a STRICT JSON object with EXACTLY these keys:
{{
  "concept_name": "punchy internal name for this idea",
  "big_idea": "1–2 sentences: the creative hook and why it stops the scroll",
  "hook_line": "the on-screen text/spoken line in the first 2 seconds",
  "shotlist": [
     {{
       "scene": "short label, e.g. 'hero reveal'",
       "image_prompt": "a vivid, specific prompt to generate this shot as a still — describe setting, framing, mood, lighting, styling. The product garment is fixed (it will be composited faithfully); describe everything AROUND it.",
       "preset": "one preset id from the menu above",
       "on_screen_text": "the caption text burned onto this shot"
     }}
  ],
  "caption": "the post caption in MS. READ's warm, aspirational voice — natural Malaysian English with a light BM touch where it feels authentic (not forced). 2–4 short lines + a question to drive comments.",
  "hashtags": ["8–12 tags mixing MY-local, fashion-niche and trend tags, no '#' symbol"],
  "cta": "one clear call to action (shop link / visit store / comment)",
  "best_post_time_myt": "best day+time to post in Malaysia time, with a one-line reason",
  "audio_suggestion": "what kind of trending sound/music to use",
  "why_it_wins": ["2–4 bullets: specifically how this beats what competitors are posting"],
  "kpi_to_watch": "the single metric that proves this worked"
}}

Rules:
- 3 or 4 shots in the shotlist.
- Every "preset" must be one of these ids: {preset_ids}.
- Keep it genuinely on-trend and platform-native, never corporate or stiff.
- Output ONLY the JSON object. No markdown fences, no commentary."""


def _parse_json_robust(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1] if "\n" in text else text
    if text.endswith("```"):
        text = text.rsplit("```", 1)[0]
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start != -1 and end != -1 and end > start:
            return json.loads(text[start:end + 1])
        raise


def _build_prompt(product: dict, trend: dict, custom_topic: str, creative_brief: str) -> str:
    if trend:
        t = trend
    else:  # free-text trend the user typed
        t = {
            "name": custom_topic or "Trending now",
            "scope": "Custom", "moment": "Now",
            "format": "Short-form vertical video", "platform": "TikTok + Instagram Reels",
            "why": "A timely angle the audience is already engaging with.",
        }
    return STRATEGY_PROMPT.format(
        brand=BRAND,
        trend_name=t.get("name", ""), trend_scope=t.get("scope", ""),
        trend_moment=t.get("moment", ""), trend_format=t.get("format", ""),
        trend_platform=t.get("platform", ""), trend_why=t.get("why", ""),
        product_title=product.get("title", ""), product_type=product.get("product_type", ""),
        product_price=product.get("price", "") or "—",
        product_desc=(product.get("description_text", "") or "")[:400],
        creative_brief=creative_brief.strip() or "(none)",
        preset_menu=_PRESET_MENU, preset_ids=", ".join(PRESETS.keys()),
    )


def build_strategy_brief(client, product: dict, trend: dict, custom_topic: str,
                         creative_brief: str, callback: Callable) -> dict:
    """Generate the strategy brief via Claude (if ANTHROPIC_API_KEY) else Gemini."""
    prompt = _build_prompt(product, trend, custom_topic, creative_brief)
    callback("status", {"phase": "strategy", "message": "Writing the viral content strategy..."})

    if os.getenv("ANTHROPIC_API_KEY"):
        try:
            brief = _brief_via_claude(prompt, callback)
            callback("status", {"phase": "strategy_done",
                                "message": f"Strategy locked (Claude {CLAUDE_MODEL})"})
            return _sanitize_brief(brief)
        except Exception as e:
            callback("status", {"phase": "strategy",
                                "message": f"Claude unavailable ({str(e)[:60]}) — using Gemini..."})

    brief = _brief_via_gemini(client, prompt, callback)
    callback("status", {"phase": "strategy_done", "message": "Strategy locked"})
    return _sanitize_brief(brief)


def _brief_via_claude(prompt: str, callback: Callable) -> dict:
    import anthropic
    aclient = anthropic.Anthropic()
    last_error = None
    for attempt, (temp, wait) in enumerate([(0.9, 5), (0.6, 15), (0.4, 30)]):
        try:
            resp = aclient.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=4000,
                temperature=temp,
                system="You are a precise JSON generator. Output ONLY one valid JSON object — no fences, no commentary.",
                messages=[{"role": "user", "content": prompt}],
            )
            text = "".join(b.text for b in resp.content if getattr(b, "type", None) == "text")
            return _parse_json_robust(text)
        except (json.JSONDecodeError, ValueError) as e:
            last_error = e
            if attempt < 2:
                time.sleep(wait)
        except Exception as e:
            es = str(e)
            transient = any(c in es for c in ("429", "500", "502", "503", "529")) or "overloaded" in es.lower()
            if transient and attempt < 2:
                time.sleep(wait)
            else:
                raise
    raise last_error


def _brief_via_gemini(client, prompt: str, callback: Callable) -> dict:
    last_error = None
    for attempt, (temp, wait) in enumerate([(0.85, 5), (0.6, 15), (0.4, 30)]):
        try:
            resp = client.models.generate_content(
                model=GEMINI_TEXT_MODEL,
                contents=prompt + "\n\nRemember: output ONLY the JSON object.",
            )
            return _parse_json_robust(resp.text)
        except (json.JSONDecodeError, ValueError) as e:
            last_error = e
            if attempt < 2:
                callback("status", {"phase": "strategy", "message": "Refining the brief..."})
                time.sleep(wait)
        except Exception as e:
            es = str(e)
            if ("503" in es or "UNAVAILABLE" in es or "overloaded" in es.lower()) and attempt < 2:
                time.sleep(wait)
            else:
                raise
    raise last_error


def _sanitize_brief(brief: dict) -> dict:
    """Coerce to the expected shape and clamp the shotlist to valid presets."""
    if not isinstance(brief, dict):
        brief = {}
    shots = brief.get("shotlist") or []
    clean_shots = []
    for s in shots[:4]:
        if not isinstance(s, dict):
            continue
        preset = s.get("preset", DEFAULT_PRESET)
        if preset not in PRESETS:
            preset = DEFAULT_PRESET
        clean_shots.append({
            "scene": str(s.get("scene", "shot"))[:40],
            "image_prompt": str(s.get("image_prompt", "")),
            "preset": preset,
            "on_screen_text": str(s.get("on_screen_text", "")),
        })
    if not clean_shots:  # never leave the pipeline without something to shoot
        clean_shots = [{
            "scene": "hero reveal",
            "image_prompt": "Cinematic editorial reveal of the product on a confident model, soft premium lighting, modern Malaysian setting.",
            "preset": DEFAULT_PRESET, "on_screen_text": "",
        }]
    brief["shotlist"] = clean_shots
    for key in ("hashtags", "why_it_wins"):
        if not isinstance(brief.get(key), list):
            brief[key] = [str(brief.get(key))] if brief.get(key) else []
    return brief


# ─────────────────────────────────────────────────────────────
# Asset generation
# ─────────────────────────────────────────────────────────────

def _generate_clips(client, brief: dict, product: dict, image_files: list,
                    creative_brief: str, output_dir: Path, api_key: str,
                    callback: Callable) -> list:
    """Animate each generated still into a viral motion clip using its shot preset."""
    video_dir = output_dir / "videos"
    video_dir.mkdir(parents=True, exist_ok=True)
    image_dir = output_dir / "images"
    shots = brief.get("shotlist", [])
    total = len(image_files)

    engine_label = "Higgsfield" if higgsfield_enabled() else "Veo 3.1"
    callback("status", {"phase": "generating_videos",
                        "message": f"Animating {total} clip(s) via {engine_label}...",
                        "total": total, "current": 0})

    videos = []
    for i, img in enumerate(image_files):
        shot = shots[i] if i < len(shots) else shots[-1] if shots else {}
        preset = shot.get("preset", DEFAULT_PRESET)
        scene = img.get("scene", f"shot_{i+1}")
        img_path = image_dir / img["filename"]
        if not img_path.exists():
            continue
        out_path = video_dir / f"video_{i+1}_{preset}.mp4"

        callback("video_start", {
            "index": i, "total": total, "scene": scene,
            "message": f"Clip {i+1}/{total}: {PRESETS.get(preset, {}).get('label', preset)}...",
        })
        result = generate_clip(
            client, img_path, out_path,
            scene_prompt=shot.get("image_prompt", "")[:200],
            preset_id=preset, product_title=product.get("title", ""),
            creative_brief=creative_brief, callback=callback, api_key=api_key,
        )
        if result.get("filename"):
            result["scene"] = scene
            videos.append(result)
            callback("video_done", {
                "index": i, "total": total, "scene": scene, "success": True,
                "filename": result["filename"],
                "message": f"{scene} done ({result['size_mb']} MB · {result['engine']})",
            })
        else:
            callback("video_done", {
                "index": i, "total": total, "scene": scene, "success": False,
                "filename": None, "message": f"{scene} — clip failed",
            })
        if i < total - 1:
            time.sleep(4)

    callback("status", {"phase": "videos_done",
                        "message": f"{len(videos)}/{total} clips generated",
                        "total_videos": len(videos)})
    return videos


def _write_brief_files(brief: dict, product: dict, trend: dict, output_dir: Path) -> None:
    """Human-readable brief for the download pack."""
    lines = [
        f"MS. READ — CONTENT + Viral Brief",
        f"Product: {product.get('title','')}",
        f"Trend: {trend.get('name','') if trend else 'Custom'}",
        "=" * 48, "",
        f"CONCEPT: {brief.get('concept_name','')}",
        f"BIG IDEA: {brief.get('big_idea','')}",
        f"HOOK: {brief.get('hook_line','')}", "",
        "SHOTLIST:",
    ]
    for i, s in enumerate(brief.get("shotlist", []), 1):
        lines.append(f"  {i}. [{s.get('preset')}] {s.get('scene')} — {s.get('on_screen_text','')}")
        lines.append(f"     {s.get('image_prompt','')}")
    lines += [
        "", f"CAPTION:\n{brief.get('caption','')}", "",
        "HASHTAGS: " + " ".join(f"#{h}" for h in brief.get("hashtags", [])),
        f"CTA: {brief.get('cta','')}",
        f"BEST TIME (MYT): {brief.get('best_post_time_myt','')}",
        f"AUDIO: {brief.get('audio_suggestion','')}", "",
        "WHY IT WINS:",
    ]
    lines += [f"  • {w}" for w in brief.get("why_it_wins", [])]
    lines.append(f"\nKPI TO WATCH: {brief.get('kpi_to_watch','')}")
    (output_dir / "viral_brief.txt").write_text("\n".join(lines), encoding="utf-8")


def _package(output_dir: Path, product_title: str, callback: Callable) -> Path:
    callback("status", {"phase": "packaging", "message": "Building download pack..."})
    safe = "".join(c if c.isalnum() else "_" for c in product_title)[:40] or "content_plus"
    zip_path = output_dir / f"content_plus_{safe}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for sub in ("images", "videos"):
            d = output_dir / sub
            if d.exists():
                for f in sorted(d.glob("*")):
                    zf.write(f, f"{sub}/{f.name}")
        brief_txt = output_dir / "viral_brief.txt"
        if brief_txt.exists():
            zf.write(brief_txt, "viral_brief.txt")
    return zip_path


# ─────────────────────────────────────────────────────────────
# Pipeline
# ─────────────────────────────────────────────────────────────

def run_content_plus_pipeline(api_key: str, product_url: str, trend_id: str,
                              custom_topic: str, creative_brief: str, num_clips: int,
                              output_dir: Path, callback: Callable) -> None:
    """Full CONTENT + run. Emits SSE callback events; writes content_plus.json."""
    client = genai.Client(api_key=api_key)
    trend = get_trend(trend_id) if trend_id else {}
    num_clips = max(1, min(4, num_clips or 3))

    # 1. Product
    callback("status", {"phase": "fetching_product", "message": "Fetching product from msreadshop.com..."})
    product = fetch_product(product_url, callback)
    callback("status", {"phase": "product_ready",
                        "message": f"Product: {product.get('title','')}",
                        "product": {"title": product.get("title", ""),
                                    "image": product.get("image_url")}})

    # 2. Strategy brief
    brief = build_strategy_brief(client, product, trend, custom_topic, creative_brief, callback)
    # Trim the shotlist to the requested clip count
    brief["shotlist"] = brief["shotlist"][:num_clips]
    callback("status", {"phase": "strategy_ready", "message": "Concept ready", "brief": brief})

    # 3. Hero images (product-faithful, trend-styled) — reuse the SKU image engine
    content = {"image_prompts": [{"scene": s["scene"], "prompt": s["image_prompt"]}
                                 for s in brief["shotlist"]]}
    image_files = generate_sku_images(client, content, product, creative_brief,
                                      output_dir, callback, avatar_images=[])

    # 4. Motion clips (Higgsfield → Veo)
    videos = _generate_clips(client, brief, product, image_files, creative_brief,
                             output_dir, api_key, callback)

    # 5. Persist + package
    _write_brief_files(brief, product, trend, output_dir)
    result = {
        "product": {"title": product.get("title", ""),
                    "url": product_url,
                    "image": product.get("image_url")},
        "trend": trend or {"name": custom_topic or "Custom", "scope": "Custom"},
        "brief": brief,
        "images": image_files,
        "videos": videos,
        "engine": "higgsfield" if higgsfield_enabled() else "veo",
    }
    (output_dir / "content_plus.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    _package(output_dir, product.get("title", "content_plus"), callback)

    callback("status", {"phase": "done",
                        "message": f"CONTENT + ready — {len(image_files)} images, {len(videos)} clips"})
