# MS. READ Content Engine

AI-powered 30-day social media content calendar generator for [MS. READ](https://www.msreadshop.com), Malaysia's leading plus-size fashion brand.

**What it does:** You type a creative direction, and it generates a complete content pack:
- 6-sheet Excel (30-day calendar, SEO blog, FB captions, image prompts, hashtags, schedule)
- 13 AI-generated editorial images (via Nano Banana / Google Gemini)
- Downloadable as a single zip file

## Quick Start

```bash
# 1. Clone
git clone https://github.com/LoneFrogie/msread-content-engine.git
cd msread-content-engine

# 2. Install
pip install -r requirements.txt

# 3. Add your API key
cp .env.example .env
# Edit .env and add your GOOGLE_AI_API_KEY
# Get a free key at: https://aistudio.google.com/apikey

# 4. Run
python app.py
```

Open **http://localhost:8001** in your browser.

## How It Works

1. **Enter your creative direction** — e.g., "Raya collection launch, focus on modest fashion in emerald and pastel pink"
2. **AI adapts the content** — Gemini rewrites all 30 days, blog, captions, and image prompts to match your brief (~60s)
3. **Images generate live** — 13 editorial photos created via Nano Banana, visible in real-time (~2.5 min)
4. **Download** — Full zip (Excel + images) or Excel only

## Screenshots

### Input — Creative Direction
Enter your campaign focus, target products, themes, or goals. Example briefs are provided as quick-start chips.

### Progress — Live Generation
Real-time progress bar with phase indicators. Image thumbnails appear as each one completes.

### Results — Download
Full image gallery with lightbox. Download as zip or Excel only.

## Tech Stack

- **Backend:** FastAPI + Python
- **Frontend:** React 18 (CDN) — single HTML file, no build step
- **AI — Text:** Google Gemini 2.5 Flash (content adaptation)
- **AI — Images:** Nano Banana / Gemini 2.5 Flash Image (editorial photography)
- **Excel:** openpyxl (6 branded sheets with styling)
- **Progress:** Server-Sent Events (SSE)

## API Key

This app uses Google's Gemini API for both text and image generation.

1. Go to [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. Click "Create API Key"
3. Copy the key into your `.env` file

The free tier supports image generation with rate limiting (10 requests/minute).

## Project Structure

```
msread-content-engine/
├── app.py              # FastAPI web server
├── engine.py           # Content generation pipeline
├── templates/
│   └── index.html      # React SPA (MS. READ branded UI)
├── requirements.txt
├── .env.example
└── README.md
```

## Team Access

Anyone on your network can access the app:

```bash
# Start the server
python app.py

# Share with colleagues
# They open: http://<your-ip>:8001
```

Only the host machine needs the API key. Colleagues just open the URL.

## License

Internal use — MS. READ / BettieAI.
