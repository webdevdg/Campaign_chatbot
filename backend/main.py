import os
from typing import Optional, List

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from services.data_service import load_data
from services.llm_service import chat as llm_chat
from services.report_service import generate_report, REPORTS_DIR

app = FastAPI(title="Campaign Analyst API")

allowed_origins_env = os.getenv("ALLOWED_ORIGINS", "http://localhost:5173")
allowed_origins = [
    origin.strip()
    for origin in allowed_origins_env.split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def startup():
    load_data()


# --- Request / Response models ---

class ChatMessage(BaseModel):
    role: str
    content: str


class ChatRequest(BaseModel):
    message: str
    conversation_history: List[ChatMessage] = []


class ChatResponse(BaseModel):
    response: str
    report_url: Optional[str] = None


class ReportRequest(BaseModel):
    report_type: str
    parameters: dict = {}


class ReportResponse(BaseModel):
    download_url: str


def _summarize_report_params(params: dict) -> str:
    """Build a readable summary of report params so the LLM can reference them in follow-ups."""
    parts = [f"**Report type:** {params.get('report_type', 'general')}"]

    if params.get("title"):
        parts.append(f"**Title:** {params['title']}")
    if params.get("brand"):
        b = params["brand"]
        parts.append(f"**Brand filter:** {b if isinstance(b, str) else ', '.join(b)}")
    if params.get("channel"):
        c = params["channel"]
        parts.append(f"**Channel filter:** {c if isinstance(c, str) else ', '.join(c)}")
    if params.get("start_date") or params.get("end_date"):
        parts.append(f"**Date range:** {params.get('start_date', 'start')} to {params.get('end_date', 'end')}")
    if params.get("group_by"):
        parts.append(f"**Grouped by:** {params['group_by']}")
    if params.get("metrics"):
        parts.append(f"**Metrics:** {', '.join(params['metrics'])}")
    if params.get("sort_by"):
        parts.append(f"**Sorted by:** {params['sort_by']} ({params.get('sort_order', 'desc')})")
    if params.get("top_n"):
        parts.append(f"**Top N:** {params['top_n']}")
    if params.get("comparison"):
        comp = params["comparison"]
        parts.append(f"**Comparison:** {comp.get('period_1_label', 'Period 1')} vs {comp.get('period_2_label', 'Period 2')}")
    if params.get("include_sheets"):
        parts.append(f"**Sheets included:** {', '.join(params['include_sheets'])}")
    if params.get("no_charts"):
        parts.append("**Charts:** excluded (tables only)")

    return "[Report Config: " + " | ".join(parts) + "]"


# --- Endpoints ---

@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/chat", response_model=ChatResponse)
def chat_endpoint(req: ChatRequest):
    try:
        history = [{"role": m.role, "content": m.content} for m in req.conversation_history]
        result = llm_chat(req.message, history)

        if result["type"] == "report":
            params = result["parameters"]
            filename = generate_report(
                report_type=params.get("report_type", "general"),
                parameters=params,
            )
            download_url = f"/api/reports/{filename}"
            params_summary = _summarize_report_params(params)
            return ChatResponse(
                response=f"Your report has been generated and is ready for download.\n\n{params_summary}",
                report_url=download_url,
            )

        return ChatResponse(response=result["content"])

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/report", response_model=ReportResponse)
def report_endpoint(req: ReportRequest):
    try:
        filename = generate_report(req.report_type, req.parameters)
        return ReportResponse(download_url=f"/api/reports/{filename}")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/reports/{filename}")
def download_report(filename: str):
    filepath = os.path.join(REPORTS_DIR, filename)
    if not os.path.isfile(filepath):
        raise HTTPException(status_code=404, detail="Report not found")
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )
