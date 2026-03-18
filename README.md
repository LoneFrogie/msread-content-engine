# MS. READ Content Engine

AI-powered 30-day social media content calendar generator for [MS. READ](https://www.msreadshop.com), Malaysia's leading plus-size fashion brand.

**What it does:** You type a creative direction, and it generates a complete content pack:
- 6-sheet Excel (30-day calendar, SEO blog, FB captions, image prompts, hashtags, schedule)
- 13 AI-generated editorial images (via Nano Banana / Google Gemini)
- Downloadable as a single zip file

## Quick Start (Local)

```bash
# 1. Clone
git clone https://github.com/LoneFrogie/msread-content-engine.git
cd msread-content-engine

# 2. Install
pip install -r requirements.txt

# 3. Run (API key is pre-configured)
python app.py
```

Open **http://localhost:8001** — no API key setup needed.

## Deploy to Google Cloud Run

One command to deploy to the cloud. Your colleagues get a public URL.

```bash
# Prerequisites (one-time):
# 1. Install gcloud CLI: https://cloud.google.com/sdk/docs/install
# 2. gcloud auth login
# 3. gcloud config set project YOUR_PROJECT_ID

# Deploy:
./deploy.sh
```

This will:
- Build a Docker container
- Deploy to Cloud Run (Singapore region — closest to Malaysia)
- Give you a public URL like `https://msread-content-engine-xxxxx-as.a.run.app`
- Share that URL with your team — no setup required on their end

## How It Works

1. **Enter your creative direction** — e.g., "Raya collection launch, focus on modest fashion in emerald and pastel pink"
2. **AI adapts the content** — Gemini rewrites all 30 days, blog, captions, and image prompts to match your brief (~60s)
3. **Images generate live** — 13 editorial photos created via Nano Banana, visible in real-time (~2.5 min)
4. **Download** — Full zip (Excel + images) or Excel only

## Tech Stack

- **Backend:** FastAPI + Python
- **Frontend:** React 18 (CDN) — single HTML file, no build step
- **AI — Text:** Google Gemini 2.5 Flash (content adaptation)
- **AI — Images:** Nano Banana / Gemini 2.5 Flash Image (editorial photography)
- **Excel:** openpyxl (6 branded sheets with styling)
- **Progress:** Server-Sent Events (SSE)
- **Hosting:** Google Cloud Run (serverless, auto-scaling)

## Project Structure

```
msread-content-engine/
├── app.py              # FastAPI web server
├── engine.py           # Content generation pipeline
├── templates/
│   └── index.html      # React SPA (MS. READ branded UI)
├── Dockerfile          # Cloud Run container
├── deploy.sh           # One-command deploy script
├── requirements.txt
├── .env.example
└── README.md
```

## License

Internal use — MS. READ / BettieAI.
