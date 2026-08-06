"""
MS. READ Content Engine — CONTENT + Content Calendar (the scheduler)

Turns the strategy into an operating plan: a rolling, week-by-week posting
calendar that maps the Malaysian retail calendar (Merdeka -> Malaysia Day ->
11.11/12.12 -> Christmas -> CNY -> Ramadan -> Raya -> post-Raya) onto the right
trend + best post time, and raises "seed-ahead" alerts for the big moments
(Raya/CNY/Merdeka) so content goes out early enough to win the leaderboard race.

Stateless by design (the app has no DB): the plan is COMPUTED from today's date
on each request, so it's always current. Weeks run Monday–Sunday (MY business
week). Times are MYT (server runs UTC; we offset +8h).
"""

from datetime import datetime, timedelta, date

from trend_catalog import get_trend

MYT_OFFSET = timedelta(hours=8)

# ── Dated MY moments (Aug 2026 -> mid-2027). `seed_weeks` = how many weeks
#    BEFORE the date content should already be running. `big` moments get alerts.
MY_MOMENTS = [
    {"name": "Merdeka / National Day", "date": "2026-08-31", "seed_weeks": 3,
     "trend_id": "merdeka_workwear", "post": ("Sat", "10:00 AM"), "big": True,
     "note": "Heritage-tone workwear + multicultural model mix."},
    {"name": "Malaysia Day", "date": "2026-09-16", "seed_weeks": 2,
     "trend_id": "merdeka_workwear", "post": ("Tue", "7:00 AM"), "big": False,
     "note": "Multicultural angle; ride the malls' Merdeka spend campaigns."},
    {"name": "Deepavali", "date": "2026-11-08", "seed_weeks": 4,
     "trend_id": "cinematic_reveal", "post": ("Sat", "10:00 AM"), "big": False,
     "note": "Festive occasion styling, jewel tones; fusion looks."},
    {"name": "11.11 mega-sale", "date": "2026-11-11", "seed_weeks": 3,
     "trend_id": "yearend_glam", "post": ("Tue", "8:00 PM"), "big": True,
     "note": "Tie content to the deal window; pin purchase intent."},
    {"name": "12.12 mega-sale", "date": "2026-12-12", "seed_weeks": 3,
     "trend_id": "yearend_glam", "post": ("Fri", "8:00 PM"), "big": True,
     "note": "Party/occasion looks + the 12.12 offer."},
    {"name": "Christmas / year-end", "date": "2026-12-25", "seed_weeks": 3,
     "trend_id": "yearend_glam", "post": ("Sat", "10:00 AM"), "big": False,
     "note": "Office-party-approved but modest partywear."},
    {"name": "Chinese New Year", "date": "2027-02-06", "seed_weeks": 4,
     "trend_id": "cny_2027", "post": ("Sat", "10:00 AM"), "big": True,
     "note": "Modern modest red; add Mandarin subs + Xiaohongshu."},
    {"name": "Ramadan begins", "date": "2027-02-08", "seed_weeks": 6,
     "trend_id": "raya_countdown_2027", "post": ("Sun", "8:30 PM"), "big": True,
     "note": "Ramadan on-ramp — accelerate the Raya countdown."},
    {"name": "Hari Raya Aidilfitri", "date": "2027-03-10", "seed_weeks": 9,
     "trend_id": "raya_countdown_2027", "post": ("Sun", "8:30 PM"), "big": True,
     "note": "#1 MY fashion moment. Seed from Jan — the TikTok leaderboard forms by Feb."},
    {"name": "Post-Raya open house", "date": "2027-03-24", "seed_weeks": 1,
     "trend_id": "raya_to_kerja_2027", "post": ("Tue", "7:00 AM"), "big": False,
     "note": "Outfit-repeat / Raya->Kerja restyle; rides the weeks-long tail."},
]

# Always-on pillars for weeks without an active moment (rotate weekly).
EVERGREEN = [
    {"trend_id": "modest_workwear_grwm", "post": ("Mon", "6:30 AM")},
    {"trend_id": "tudung_9hour", "post": ("Tue", "7:00 AM")},
    {"trend_id": "one_piece_five_days", "post": ("Sun", "8:30 PM")},
    {"trend_id": "fabric_talk_asmr", "post": ("Tue", "7:30 AM")},
]


def _today_myt() -> date:
    return (datetime.utcnow() + MYT_OFFSET).date()


def _monday_of(d: date) -> date:
    return d - timedelta(days=d.weekday())


def _parse(d: str) -> date:
    return datetime.strptime(d, "%Y-%m-%d").date()


def _trend_name(trend_id: str) -> str:
    t = get_trend(trend_id)
    return t.get("name", trend_id)


def build_plan(weeks: int = 12) -> dict:
    """Compute a rolling week-by-week posting plan starting from this week."""
    weeks = max(4, min(26, weeks))
    today = _today_myt()
    start_monday = _monday_of(today)

    moments = [{**m, "d": _parse(m["date"])} for m in MY_MOMENTS]

    plan = []
    ev_i = 0
    for w in range(weeks):
        wk_start = start_monday + timedelta(days=7 * w)
        wk_end = wk_start + timedelta(days=6)

        # A moment "claims" a week if the week overlaps [date - seed_weeks, date].
        active = None
        for m in moments:
            seed_start = m["d"] - timedelta(days=7 * m["seed_weeks"])
            if seed_start <= wk_end and m["d"] >= wk_start:
                if active is None or m["d"] < active["d"]:
                    active = m

        if active:
            post_day, post_time = active["post"]
            entry = {
                "week_start": wk_start.isoformat(),
                "week_label": wk_start.strftime("%d %b %Y"),
                "type": "moment",
                "moment": active["name"],
                "moment_date": active["date"],
                "trend_id": active["trend_id"],
                "trend_name": _trend_name(active["trend_id"]),
                "post_day": post_day,
                "post_time": post_time,
                "note": active["note"],
                "is_launch_week": wk_start <= active["d"] <= wk_end,
            }
        else:
            pillar = EVERGREEN[ev_i % len(EVERGREEN)]
            ev_i += 1
            post_day, post_time = pillar["post"]
            entry = {
                "week_start": wk_start.isoformat(),
                "week_label": wk_start.strftime("%d %b %Y"),
                "type": "evergreen",
                "moment": None,
                "trend_id": pillar["trend_id"],
                "trend_name": _trend_name(pillar["trend_id"]),
                "post_day": post_day,
                "post_time": post_time,
                "note": "Always-on pillar — build the franchise between moments.",
                "is_launch_week": False,
            }
        plan.append(entry)

    # Seed-ahead alerts for the big moments.
    alerts = []
    for m in moments:
        days_out = (m["d"] - today).days
        if days_out < 0 or not m["big"]:
            continue
        weeks_out = days_out // 7
        seed = m["seed_weeks"]
        if weeks_out <= seed:
            alerts.append({
                "level": "seed_now", "moment": m["name"], "weeks_out": weeks_out,
                "trend_id": m["trend_id"], "trend_name": _trend_name(m["trend_id"]),
                "message": f"SEED NOW — {m['name']} is {weeks_out} week(s) away. Start posting {_trend_name(m['trend_id'])}.",
            })
        elif weeks_out <= seed + 4:
            alerts.append({
                "level": "prep", "moment": m["name"], "weeks_out": weeks_out,
                "trend_id": m["trend_id"], "trend_name": _trend_name(m["trend_id"]),
                "message": f"Coming up — {m['name']} in {weeks_out} weeks. Plan {_trend_name(m['trend_id'])} soon.",
            })
    alerts.sort(key=lambda a: a["weeks_out"])

    return {"generated": today.isoformat(), "weeks": weeks,
            "alerts": alerts[:3], "plan": plan}
