"""
MS. READ Content Engine — Product Description Page (PDP) Optimizer
Fetches product data + images from msreadshop.com, then generates
a premium structured product description via Gemini.

Output format:
  1. Quick Specs & Fit
  2. The Style Story
  3. Why You'll Love It
  4. How To Style
  5. Sizing Recommendation
"""

import json
import re
import time
from io import BytesIO
from pathlib import Path
from typing import Callable

import requests
from google import genai
from google.genai import types
from PIL import Image as PILImage

from sku_engine import (
    extract_handle_from_url,
    fetch_product,
    _parse_json_robust,
    BRAND_NAME,
)
from humanizer import humanize_content

# ═══════════════════════════════════════════════════════════════
# Gemini prompt for structured product description
# ═══════════════════════════════════════════════════════════════

PDP_PROMPT = """You are a senior product copywriter for MS. READ, Malaysia's leading plus-size fashion brand (UK sizes 10-24+, founded 1997).

BRAND VOICE:
- Warm, empowering, inclusive, confident
- NEVER use diet culture language or apologetic tone about size
- Malaysian English with premium, editorial tone
- Celebrate the female form — never hide it

You are given PRODUCT DATA and PRODUCT IMAGES. Analyze the images carefully to identify:
- Exact garment type, silhouette, and construction details
- Fabric texture, pattern, print details
- Neckline, sleeve style, hem type, closures
- Fit characteristics visible on the model

PRODUCT DATA:
- Name: {title}
- Type: {product_type}
- SKU: {sku}
- Price: RM{price}{price_note}
- Sizes Available: {sizes}
- Colors Available: {colors}
- Fabric (from listing): {fabric}
- Care (from listing): {care}
- Current Description: {description}
- URL: {url}

{creative_brief_section}

Generate a JSON object with these EXACT keys:

1. "quick_specs": Object with:
   - "sku": The product SKU code
   - "fabric": Detailed fabric composition (e.g., "78% Nylon, 22% Spandex (High-Stretch Ribbed Knit)") — infer from images + listing data
   - "fit_type": Silhouette/fit description (e.g., "A-Line Silhouette / Body-Contouring Fit")
   - "key_features": Comma-separated key construction features (e.g., "Sleeveless, Round Neck, Straight Hem, Unlined for Lightweight Comfort")
   - "care": Care instructions (e.g., "Machine Wash Cold | Dry Flat | Wash Inside Out | Dark Colors Separately")

2. "style_story": A 100-120 word narrative paragraph. Write in second person ("you"). Describe the garment's purpose, when/where to wear it, and the feeling it evokes. Reference specific design details visible in the images. End with a signature phrase that captures the essence (e.g., "Simply Polished", "Effortlessly Elegant").

3. "why_youll_love_it": Array of 4 benefit objects, each with:
   - "title": Bold benefit headline (e.g., "Superior Sculpting Fabric")
   - "description": 30-50 word benefit explanation. Be specific about fabric, construction, or climate suitability. Reference Malaysian weather/lifestyle where relevant.

4. "how_to_style": Array of 3 styling suggestion objects, each with:
   - "look_name": Name of the look (e.g., "The Corporate Look", "The Weekend Edit", "Layering Tip")
   - "description": 20-35 word specific styling suggestion. Mention actual garment types to pair with (not generic). Reference MS. READ complementary pieces where possible.

5. "sizing_recommendation": A 40-60 word sizing paragraph. Reference the model's size if visible. Mention fabric stretch properties. Give specific advice for true-to-size vs sizing up/down based on the garment's construction and fit.

6. "seo_title": SEO-optimized product title (60-70 chars)
7. "meta_description": SEO meta description (150-155 chars)

IMPORTANT: Analyze the product IMAGES carefully. The description must match what is VISUALLY shown — the actual garment design, not a generic description.

Respond with ONLY valid JSON, no markdown code fences."""


def _download_product_images(product: dict, max_images: int = 3) -> list:
    """Download product images for Gemini visual analysis."""
    downloaded = []
    image_urls = product.get("image_urls", [])
    if not image_urls and product.get("image_url"):
        image_urls = [product["image_url"]]

    for url in image_urls[:max_images]:
        try:
            resp = requests.get(url, timeout=15, headers={
                "User-Agent": "Mozilla/5.0 (compatible; MSReadContentEngine/1.0)"
            })
            if resp.status_code == 200:
                img = PILImage.open(BytesIO(resp.content))
                if max(img.size) > 1024:
                    ratio = 1024 / max(img.size)
                    img = img.resize((int(img.width * ratio), int(img.height * ratio)), PILImage.LANCZOS)
                buf = BytesIO()
                img.save(buf, format="PNG")
                downloaded.append(buf.getvalue())
        except Exception:
            continue

    return downloaded


def generate_pdp_content(client, product: dict, creative_brief: str, callback: Callable) -> dict:
    """Generate optimized product description using Gemini with product images."""
    callback("status", {"phase": "generating_content", "message": "Analyzing product images & generating description..."})

    # Download product images for visual analysis
    ref_images = _download_product_images(product)
    if ref_images:
        callback("status", {"phase": "generating_content", "message": f"Analyzing {len(ref_images)} product image(s) with AI..."})

    price_note = ""
    if product.get("compare_price"):
        price_note = f" (was RM{product['compare_price']})"

    creative_brief_section = ""
    if creative_brief and creative_brief.strip():
        creative_brief_section = f"ADDITIONAL DIRECTION:\n{creative_brief}\n\nIncorporate this direction into the product description tone and focus."

    prompt_text = PDP_PROMPT.format(
        title=product["title"],
        product_type=product["product_type"],
        sku=product.get("sku", "N/A"),
        price=product["price"],
        price_note=price_note,
        sizes=", ".join(product["sizes"]) if product["sizes"] else "UK 10-24+",
        colors=", ".join(product["colors"]) if product["colors"] else "Multiple colors",
        fabric=product.get("fabric", "See images"),
        care=product.get("care", "See label"),
        description=product.get("description_text", "")[:500],
        url=product["url"],
        creative_brief_section=creative_brief_section,
    )

    # Build multimodal content: images + text
    content_parts = []
    for img_bytes in ref_images:
        content_parts.append(types.Part.from_bytes(data=img_bytes, mime_type="image/png"))
    content_parts.append(types.Part.from_text(text=prompt_text))

    # Try up to 2 attempts
    last_error = None
    for attempt, temp in enumerate([0.6, 0.3]):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=content_parts,
                config=types.GenerateContentConfig(
                    temperature=temp,
                    max_output_tokens=8000,
                    response_mime_type="application/json",
                ),
            )

            text = ""
            for part in response.candidates[0].content.parts:
                if part.text:
                    text += part.text

            text = text.strip()
            if text.startswith("```"):
                text = text.split("\n", 1)[1]
            if text.endswith("```"):
                text = text.rsplit("```", 1)[0]
            text = text.strip()

            content = _parse_json_robust(text)
            callback("status", {"phase": "content_generated", "message": "Product description optimized successfully"})
            return content

        except (json.JSONDecodeError, ValueError) as e:
            last_error = e
            if attempt == 0:
                callback("status", {"phase": "generating_content", "message": "Retrying with adjusted parameters..."})
                time.sleep(2)

    raise last_error


# ═══════════════════════════════════════════════════════════════
# ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════

def run_pdp_pipeline(api_key: str, product_url: str, creative_brief: str,
                     output_dir: Path, callback: Callable):
    """Run the product description optimization pipeline."""
    try:
        client = genai.Client(api_key=api_key)

        # Phase 1: Fetch product
        product = fetch_product(product_url, callback)

        # Phase 2: Generate optimized description
        content = generate_pdp_content(client, product, creative_brief, callback)

        # Phase 2.5: Humanize — remove AI writing patterns
        content = humanize_content(client, content, callback, fields_to_humanize=[
            "style_story", "why_youll_love_it", "how_to_style", "sizing_recommendation",
        ])

        # Phase 3: Save as JSON for the frontend to render
        output_path = output_dir / "pdp_content.json"
        with open(output_path, "w") as f:
            json.dump({
                "product": {
                    "title": product["title"],
                    "price": product["price"],
                    "compare_price": product.get("compare_price"),
                    "image_url": product.get("image_url"),
                    "image_urls": product.get("image_urls", [])[:5],
                    "url": product["url"],
                    "sku": product.get("sku", ""),
                },
                "description": content,
            }, f, indent=2)

        callback("status", {
            "phase": "done",
            "message": "Product description ready!",
            "pdp_content": content,
        })

    except Exception as e:
        callback("error", {"message": f"Pipeline failed: {str(e)}"})
        raise
