"""
MS. READ Content Engine — Humanizer Post-Processor
Removes AI writing patterns from generated content using Gemini.
Based on the humanizer kit (28 AI writing patterns).
"""

import json
import time
import logging
from typing import Callable, Optional

from google.genai import types

logger = logging.getLogger(__name__)

# Condensed humanizer prompt — key patterns that matter for marketing copy
HUMANIZER_PROMPT = """You are an editor that removes AI-generated writing patterns from marketing copy.

RULES — fix these patterns:
1. Remove inflated significance: "stands as", "testament to", "pivotal", "setting the stage"
2. Cut trailing -ing phrases: "highlighting...", "ensuring...", "showcasing...", "reflecting..."
3. Replace promotional filler: "nestled", "vibrant", "breathtaking", "renowned", "boasts" → use "is", "has"
4. Kill AI vocabulary: Additionally, delve, tapestry, landscape (abstract), fostering, garner, underscore, interplay, intricate, crucial, showcase, enduring, enhance
5. Use "is"/"has" instead of "serves as"/"features"/"boasts"
6. Remove "It's not just X, it's Y" and "Not only...but..."
7. Don't force ideas into groups of three
8. Replace em dashes with commas, colons, or periods
9. Remove filler: "In order to" → "To", "Due to the fact that" → "Because"
10. Cut excessive hedging and stacked qualifiers
11. Remove "exciting times ahead" / generic positive conclusions
12. Cut "Maybe both.", "And honestly?", "Maybe that's the point."
13. Remove parenthetical personality: "(and honestly?)", "(not that I'm complaining)"
14. Use straight quotes, not curly quotes

PRESERVE:
- The brand voice (warm, empowering, inclusive, confident)
- All factual product details, prices, sizes, SKUs
- Malaysian English nuances
- Emojis in social media captions (they belong there)
- JSON structure — return ONLY valid JSON
- All hashtags exactly as-is

You will receive a JSON object. Humanize ONLY the text content within the values.
Do NOT change keys, structure, arrays, numbers, URLs, SKUs, hashtags, or emoji usage.

Return the humanized JSON object. Respond with ONLY valid JSON, no markdown."""


def humanize_content(client, content: dict, callback: Callable,
                     fields_to_humanize: Optional[list] = None) -> dict:
    """
    Post-process generated content through the humanizer.

    Args:
        client: Gemini client
        content: The generated content dict
        callback: Progress callback
        fields_to_humanize: Optional list of top-level keys to humanize.
                           If None, humanizes the entire content.
    """
    callback("status", {
        "phase": "humanizing",
        "message": "Polishing copy — removing AI writing patterns..."
    })

    # If specific fields requested, extract only those
    if fields_to_humanize:
        subset = {}
        for key in fields_to_humanize:
            if key in content:
                subset[key] = content[key]
        if not subset:
            return content
        to_humanize = subset
    else:
        to_humanize = content

    input_json = json.dumps(to_humanize, indent=2, ensure_ascii=False)

    # If content is very large, split into chunks
    if len(input_json) > 15000:
        return _humanize_chunked(client, content, fields_to_humanize or list(content.keys()), callback)

    prompt = f"{HUMANIZER_PROMPT}\n\nJSON to humanize:\n{input_json}"

    last_error = None
    attempts = [(0.4, 2), (0.3, 10), (0.3, 20)]
    for attempt, (temp, wait) in enumerate(attempts):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=temp,
                    max_output_tokens=32000,
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

            humanized = json.loads(text)

            # Merge humanized fields back into original content
            if fields_to_humanize:
                for key in fields_to_humanize:
                    if key in humanized:
                        content[key] = humanized[key]
                return content
            else:
                return humanized

        except (json.JSONDecodeError, ValueError) as e:
            last_error = e
            if attempt < len(attempts) - 1:
                logger.warning(f"Humanizer JSON parse failed, retrying: {e}")
                time.sleep(wait)
        except Exception as e:
            last_error = e
            err_str = str(e)
            if ("503" in err_str or "UNAVAILABLE" in err_str or "overloaded" in err_str.lower()) and attempt < len(attempts) - 1:
                logger.warning(f"Humanizer API overloaded, waiting {wait}s")
                time.sleep(wait)
            else:
                break

    # If humanizer fails, return original content (non-fatal)
    logger.warning(f"Humanizer could not process content, returning original: {last_error}")
    callback("status", {
        "phase": "humanizing",
        "message": "Humanizer skipped (content preserved as-is)"
    })
    return content


def _humanize_chunked(client, content: dict, keys: list, callback: Callable) -> dict:
    """Humanize large content by processing each top-level key separately."""
    for i, key in enumerate(keys):
        if key not in content:
            continue

        val = content[key]
        # Skip non-text structures (numbers, booleans, None)
        if not isinstance(val, (dict, list, str)):
            continue
        # Skip keys that are purely structural/numeric
        if key in ("session_id", "mode", "product_url"):
            continue

        chunk_json = json.dumps({key: val}, indent=2, ensure_ascii=False)
        if len(chunk_json) < 100:
            continue

        prompt = f"{HUMANIZER_PROMPT}\n\nJSON to humanize:\n{chunk_json}"

        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.4,
                    max_output_tokens=32000,
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

            result = json.loads(text.strip())
            if key in result:
                content[key] = result[key]

            callback("status", {
                "phase": "humanizing",
                "message": f"Polishing copy ({i + 1}/{len(keys)})..."
            })
        except Exception as e:
            logger.warning(f"Humanizer chunk '{key}' failed, keeping original: {e}")
            continue

    return content
