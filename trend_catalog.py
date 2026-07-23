"""
MS. READ Content Engine — CONTENT + Trend Catalog

Single source of truth for the "Trend Radar" the CONTENT + tab shows. Both the
backend engine and the frontend (via /api/content-plus/trends) read this, so a
trend can be refreshed in ONE place without touching the UI.

Each trend maps a *currently-hot format or cultural moment* (Malaysia-local or
global) to a Higgsfield motion preset and a ready-to-shoot angle for MS. READ —
elegant/modest Malaysian women's fashion. `preset` values must match keys in
higgsfield_adapter.PRESETS.

`scope`  : "MY" (Malaysia-specific) | "Global" (worldwide format)
`moment` : a dated cultural window, or "Evergreen" for always-on formats
"""

# NOTE: dates below target the Aug 2026 -> mid-2027 MY retail calendar.
TREND_CATALOG = [
    # ─────────────── Malaysia cultural moments ───────────────
    {
        "id": "merdeka_pride",
        "name": "Merdeka / Malaysia Day Pride",
        "scope": "MY",
        "moment": "Merdeka 31 Aug · Malaysia Day 16 Sep",
        "format": "Patriotic styling reel — modern takes on heritage tones",
        "preset": "dolly_in",
        "platform": "Instagram Reels + TikTok",
        "hook_idea": "'Merdeka but make it fashion' — office-to-flag-day looks in batik-inspired prints",
        "why": "Local brands post flat graphics; a cinematic styling reel owns the moment emotionally.",
    },
    {
        "id": "raya_countdown_2027",
        "name": "Raya 2027 Early Countdown",
        "scope": "MY",
        "moment": "Ramadan ~Feb 2027 · Raya ~20 Mar 2027 (plan from Dec)",
        "format": "Baju kurung / modest occasion reveal + family-moment storytelling",
        "preset": "orbit_360",
        "platform": "TikTok + Instagram",
        "hook_idea": "'Raya fitting starts now' — 360° reveal of a modest occasion look, jewel tones",
        "why": "Raya is the #1 MY fashion moment; brands that seed content early win the pre-order race.",
    },
    {
        "id": "modest_workwear_grwm",
        "name": "Modest Workwear GRWM",
        "scope": "MY",
        "moment": "Evergreen (Mon–Fri office cycle)",
        "format": "Get-ready-with-me, desk-to-dinner styling",
        "preset": "handheld_vlog",
        "platform": "TikTok",
        "hook_idea": "'Corporate but make it soft' — one blouse styled 3 ways for the KL working woman",
        "why": "Speaks to MS. READ's core 25–45 working-woman customer that fast-fashion ignores.",
    },
    {
        "id": "mall_popup_fomo",
        "name": "In-Mall Drop / Store FOMO",
        "scope": "MY",
        "moment": "Evergreen (weekend footfall)",
        "format": "FPV walk-through of the in-store new drop",
        "preset": "fpv_drone",
        "platform": "TikTok + Instagram",
        "hook_idea": "'New in at [mall] now' — fly-through of the rack, ending on the hero piece",
        "why": "Turns MS. READ's omni-channel mall presence into a content advantage pure-online rivals lack.",
    },
    # ─────────────── Global viral formats ───────────────
    {
        "id": "outfit_morph",
        "name": "Outfit-Morph Transition",
        "scope": "Global",
        "moment": "Evergreen",
        "format": "Seamless outfit-change / product-morph transition",
        "preset": "crash_zoom",
        "platform": "TikTok + Reels",
        "hook_idea": "One snap → the plain outfit morphs into the styled MS. READ look",
        "why": "Transitions are the most-shared fashion format globally; instant scroll-stopper.",
    },
    {
        "id": "cinematic_reveal",
        "name": "Cinematic Hero Reveal",
        "scope": "Global",
        "moment": "Evergreen (new arrivals)",
        "format": "Slow film-grade dolly reveal of a single hero piece",
        "preset": "dolly_in",
        "platform": "Instagram Reels",
        "hook_idea": "Editorial dolly-in on the drape and detail of a new-arrival dress",
        "why": "Elevates perceived value — makes MS. READ read as premium, not mass.",
    },
    {
        "id": "texture_asmr",
        "name": "Fabric & Detail ASMR",
        "scope": "Global",
        "moment": "Evergreen",
        "format": "Macro 'textures & details' beauty shots, float motion",
        "preset": "float",
        "platform": "TikTok + Reels",
        "hook_idea": "Weightless close-ups of lace, pleats and embroidery drifting in soft light",
        "why": "Sensory content converts on quality perception — ideal for occasion/premium lines.",
    },
    {
        "id": "ai_ugc_testimonial",
        "name": "AI-UGC Style Testimonial",
        "scope": "Global",
        "moment": "Evergreen (always-on ads)",
        "format": "Authentic talking-style UGC recommendation",
        "preset": "handheld_vlog",
        "platform": "TikTok + Meta ads",
        "hook_idea": "'The one blouse I keep re-buying' — relatable first-person style story",
        "why": "UGC-style creative outperforms polished ads on cost-per-result across Meta/TikTok.",
    },
]


def get_trend(trend_id: str) -> dict:
    for t in TREND_CATALOG:
        if t["id"] == trend_id:
            return t
    return {}
