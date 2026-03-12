# Campaign Analyst

A chatbot that answers natural-language questions about media campaign performance data for Q1 2026. It uses GPT-4o to analyze a 1,350-row dataset spanning three brands and five advertising channels, and can generate downloadable Excel reports on demand.

## Prerequisites

- Python 3.9+
- Node 18+
- An OpenAI API key with access to the `gpt-4o` model

## Project Structure

```
├── data/
│   ├── generate_data.py      # Script to regenerate the dataset
│   └── campaign_data.csv     # 3-month campaign performance data
├── backend/
│   ├── main.py               # FastAPI application
│   ├── services/
│   │   ├── data_service.py   # CSV loading and caching
│   │   ├── llm_service.py    # OpenAI integration
│   │   └── report_service.py # Excel report generation
│   ├── reports/              # Generated Excel files
│   ├── requirements.txt
│   └── .env.example
├── frontend/
│   └── src/
│       ├── components/       # React UI components
│       ├── hooks/useChat.js  # Chat state management
│       └── App.jsx
└── README.md
```

## Backend Setup

```bash
cd backend
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Add your OpenAI API key
cp .env.example .env
# Edit .env and set OPENAI_API_KEY=sk-...

uvicorn main:app --reload --port 8000
```

The API will be available at `http://localhost:8000`. Verify with `curl http://localhost:8000/api/health`.

## Frontend Setup

```bash
cd frontend
npm install
npm run dev
```

The app will open at `http://localhost:5173`.

## Deployment (Vercel + Render)

This repo is a monorepo:

- `frontend/`: Vite + React static app (deploy to Vercel)
- `backend/`: FastAPI API (deploy to Render)

### 1) Deploy the backend to Render

- Create a new **Web Service** on Render from this GitHub repo (you can also use the included `render.yaml` blueprint).
- Set:
  - **Build Command**: `pip install -r backend/requirements.txt`
  - **Start Command**: `cd backend && uvicorn main:app --host 0.0.0.0 --port $PORT`
- Add environment variables:
  - `OPENAI_API_KEY`: your OpenAI key
  - `ALLOWED_ORIGINS`: your Vercel domain, e.g. `https://your-app.vercel.app`

After deploy, copy your backend URL (it will look like `https://<service>.onrender.com`).

### 2) Deploy the frontend to Vercel

- Import this GitHub repo in Vercel.
- Set **Root Directory** to `frontend/`.
- Add an environment variable:
  - `VITE_API_BASE_URL`: your Render backend URL, e.g. `https://<service>.onrender.com`

Redeploy, then your frontend should be able to call the backend in production.

## Sample Questions

- **Comparison:** "Compare Apex Athletics performance: Week 1 vs Week 4"
- **Ranking:** "Which channel had the best ROAS across all brands?"
- **Anomaly:** "Analyze TechNova's February performance dip"
- **Report:** "Generate a full Q1 performance report for all brands"
- **Trend:** "Show me the weekly spend trend for Luminara Beauty on Meta"
- **Summary:** "What was the total revenue and CPA for each brand in March?"
