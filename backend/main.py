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

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
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
            return ChatResponse(
                response="Your report has been generated and is ready for download.",
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
