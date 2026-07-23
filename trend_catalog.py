"""
MS. READ Content Engine — CONTENT + Trend Catalog

Single source of truth for the "Trend Radar" the CONTENT + tab shows. Both the
backend engine and the frontend (via /api/content-plus/trends) read this, so a
trend can be refreshed in ONE place without touching the UI.

Each trend maps a *currently-hot format or cultural moment* (Malaysia-local or
global) to a Higgsfield motion preset and a ready-to-shoot angle for MS. READ —
elegant, size-inclusive Malaysian women's fashion. `preset` values must match
keys in higgsfield_adapter.PRESETS.

Grounded in mid-2026 research (MY retail calendar Aug 2026 -> mid-2027 + current
global short-video formats). The strategic thread across the MY entries is the
white space competitors under-serve: **modest WORKWEAR for real working women**,
plus hijab/fit education and plus-size/older representation — versus the festive/
aspirational content everyone else over-indexes on.

`scope`  : "MY" (Malaysia-specific) | "Global" (worldwide format)
`moment` : a dated cultural window, or "Evergreen" for always-on formats
"""

TREND_CATALOG = [
    # ─────────── Malaysia: the workwear/education white space ───────────
    {
        "id": "modest_workwear_grwm",
        "name": "Monday Modest GRWM  #KerjaFits",
        "scope": "MY",
        "moment": "Evergreen · Mon–Fri office cycle",
        "format": "Narrative GRWM — get ready for a work moment",
        "preset": "handheld_vlog",
        "platform": "TikTok",
        "hook_idea": "'GRWM for a client pitch — modest, wrinkle-free, under RM200'",
        "why": "Everyone does festive GRWM; almost no one owns office/workwear GRWM — MS. READ's core 25–45 working woman.",
    },
    {
        "id": "tudung_9hour",
        "name": "No-Slip Tudung, 9-Hour Day",
        "scope": "MY",
        "moment": "Evergreen · educational series",
        "format": "Edu-entertain hijab tutorial (teach, don't sell)",
        "preset": "handheld_vlog",
        "platform": "TikTok (BM-led)",
        "hook_idea": "'3 tudung wraps that stay neat on camera through a full workday'",
        "why": "Hijab-for-the-office education is the exact content Poplook, dUCk & Nafeesa under-produce.",
    },
    {
        "id": "one_piece_five_days",
        "name": "One Piece, Five Days",
        "scope": "MY",
        "moment": "Evergreen · value/versatility",
        "format": "Before/after hero-piece restyling",
        "preset": "crash_zoom",
        "platform": "Instagram Reels + TikTok",
        "hook_idea": "One midi-kurung → office / client-dinner / casual-Friday / weekend / open-house",
        "why": "Restyle reels drive 3–5× more link requests than flat-lays and answer the 'is it worth it' shopper.",
    },
    {
        "id": "raya_to_kerja_2027",
        "name": "Raya → Kerja Outfit-Repeat",
        "scope": "MY",
        "moment": "Post-Raya tail · after 10 Mar 2027",
        "format": "Sustainability-led restyle of a baju raya",
        "preset": "orbit_360",
        "platform": "TikTok + Reels",
        "hook_idea": "'Don't retire your baju raya — 3 Monday-office ways to wear it again'",
        "why": "Rides the weeks-long post-Raya open-house tail AND the 'is festive fashion wasteful' sentiment.",
    },
    {
        "id": "merdeka_workwear",
        "name": "Jalur Gemilang at Work",
        "scope": "MY",
        "moment": "Merdeka 31 Aug · Malaysia Day 16 Sep 2026",
        "format": "Patriotic colour-story workwear capsule",
        "preset": "dolly_in",
        "platform": "TikTok + Instagram",
        "hook_idea": "Merdeka office looks in heritage tones on a Malay / Chinese / Indian model trio",
        "why": "Malls run 4-week Merdeka spend campaigns to piggyback; rivals only post flat flag graphics.",
    },
    {
        "id": "raya_countdown_2027",
        "name": "Raya 2027 Countdown Reveal",
        "scope": "MY",
        "moment": "Ramadan ~Feb · Raya 10 Mar 2027 (seed from Jan)",
        "format": "Modest occasion reveal + family-moment story",
        "preset": "orbit_360",
        "platform": "TikTok + Instagram",
        "hook_idea": "360° reveal of a Raya occasion look in this year's colour story",
        "why": "Raya is the #1 MY fashion moment; a Top-20 TikTok leaderboard forms by February — seed content early.",
    },
    {
        "id": "live_rack_raid",
        "name": "Live Rack Raid — Drop Tease",
        "scope": "MY",
        "moment": "Evergreen · TikTok Shop live cadence",
        "format": "Product-morph transition teasing a live-sell",
        "preset": "fpv_drone",
        "platform": "TikTok Shop",
        "hook_idea": "New-arrival rack morphs SKU-to-SKU → 'full reveal tonight, TikTok Live 9pm'",
        "why": "Turns MS. READ's mall-boutique omnichannel edge into content; live converts 5–20% in-session.",
    },
    {
        "id": "cny_2027",
        "name": "CNY 2027 — Elegant Red",
        "scope": "MY",
        "moment": "CNY 6 Feb 2027 (seed from Jan)",
        "format": "Festive occasion styling, modest cheongsam-cut",
        "preset": "dolly_in",
        "platform": "TikTok + Xiaohongshu",
        "hook_idea": "Modern modest takes on lucky red for the office CNY open house",
        "why": "Multicultural festive reach the modestwear rivals skip; Xiaohongshu + Mandarin subtitles widen it.",
    },
    {
        "id": "yearend_glam",
        "name": "Year-End Party Glam",
        "scope": "MY",
        "moment": "11.11 · 12.12 · Christmas (Nov–Dec 2026)",
        "format": "Occasion/partywear reveal tied to mega-sales",
        "preset": "robo_arm",
        "platform": "Instagram Reels + TikTok",
        "hook_idea": "'Office-party-approved, still modest' — evening looks under the 12.12 deal",
        "why": "Pins content to the 11.11/12.12 purchase intent window instead of generic sale graphics.",
    },
    # ─────────────── Global viral formats ───────────────
    {
        "id": "modest_boss_pov",
        "name": "POV: The Modest Boss",
        "scope": "Global",
        "moment": "Evergreen · brand-persona anchor",
        "format": "POV / first-person aspirational identity",
        "preset": "bullet_time",
        "platform": "TikTok + Reels",
        "hook_idea": "'POV: you walk into the boardroom, overdressed by MS. READ'",
        "why": "First-person immersion builds a repeatable brand muse rivals' flat-lays can't create.",
    },
    {
        "id": "fabric_talk_asmr",
        "name": "Fabric Talk — Desk-Proof",
        "scope": "Global",
        "moment": "Evergreen · trust-building",
        "format": "Textures & details macro + ASMR",
        "preset": "float",
        "platform": "TikTok + Xiaohongshu",
        "hook_idea": "Slow macro on workwear fabric — 'survives KL heat + a 9-hour day, no iron'",
        "why": "Edu-entertain sensory proof converts on quality perception — the opposite of price-led fast fashion.",
    },
    {
        "id": "cinematic_reveal",
        "name": "Cinematic Hero Reveal",
        "scope": "Global",
        "moment": "Evergreen · new arrivals",
        "format": "Slow film-grade dolly reveal of one hero piece",
        "preset": "dolly_in",
        "platform": "Instagram Reels",
        "hook_idea": "Editorial dolly-in on the drape and detail of a new-arrival piece",
        "why": "Premium camera language makes MS. READ read as elevated, not mass-market.",
    },
]


def get_trend(trend_id: str) -> dict:
    for t in TREND_CATALOG:
        if t["id"] == trend_id:
            return t
    return {}
