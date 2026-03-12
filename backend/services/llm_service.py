import json
import os
import re

from openai import OpenAI
from dotenv import load_dotenv

from services.data_service import get_data_summary, get_csv_text

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

SYSTEM_PROMPT_TEMPLATE = """You are a senior media campaign performance analyst. You have access to a complete Q1 2026 (January–March) media campaign dataset for three brands: Apex Athletics, Luminara Beauty, and TechNova, across five channels: Google Search, Meta (Facebook/Instagram), YouTube Video, Programmatic Display, and TikTok.

{summary}

=== Full Raw Data (CSV) ===
{csv_data}

RESPONSE FORMAT RULES (follow strictly):
Your responses are rendered as Markdown in a chat UI. You MUST use rich Markdown formatting in every answer:

1. **Structure**: Start with a brief 1-2 sentence summary, then present data. Use ## and ### headings to organize sections.
2. **Tables are mandatory for comparisons**: Whenever you compare brands, channels, time periods, or list metrics side-by-side, you MUST use a Markdown table. Never list metrics as plain text when a table would be clearer.
   Example table format:
   | Brand | Spend | Revenue | ROAS |
   |-------|-------|---------|------|
   | Apex  | $X    | $Y      | Z    |
3. **Bold key numbers**: Always **bold** the most important metric values and highlight the best/worst performers.
4. **Number formatting**: Use $ for currency, commas for thousands, % for percentages. Round currency to 2 decimals, percentages to 2 decimals.
5. **Comparisons**: When comparing periods, always show a table with both absolute values and percentage change (use a Δ Change column).
6. **Insights**: End with a brief "Key Takeaway" or "Insight" section summarizing the main finding in 1-2 sentences.
7. **Lists**: Use bullet points for qualitative insights or recommendations, numbered lists for rankings.

ANALYSIS RULES:
- Answer precisely using the provided data. Always cite specific numbers.
- Be concise but thorough. If a question is ambiguous, state your assumptions.
- For trend analysis, break data into meaningful time periods (weekly or monthly).

REPORT GENERATION:
If the user asks to generate/create/download a report or export data to Excel, respond with EXACTLY this JSON and nothing else:
{{"action": "generate_report", "parameters": {{"report_type": "<type>", "brand": "<brand or null>", "channel": "<channel or null>", "start_date": "<YYYY-MM-DD or null>", "end_date": "<YYYY-MM-DD or null>"}}}}
Fill in the parameters based on the user's request. Use null for any filter not specified."""


def _build_system_prompt() -> str:
    return SYSTEM_PROMPT_TEMPLATE.format(
        summary=get_data_summary(),
        csv_data=get_csv_text(),
    )


def chat(user_message: str, conversation_history: "list[dict]") -> dict:
    system_prompt = _build_system_prompt()

    messages = [{"role": "system", "content": system_prompt}]
    for msg in conversation_history:
        messages.append({"role": msg["role"], "content": msg["content"]})
    messages.append({"role": "user", "content": user_message})

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        temperature=0.2,
        max_tokens=4096,
    )

    content = response.choices[0].message.content.strip()

    # Check if the response is a report generation action
    report_action = _parse_report_action(content)
    if report_action:
        return {"type": "report", "parameters": report_action["parameters"]}

    return {"type": "text", "content": content}


def _parse_report_action(content: str):
    # Try to extract JSON from the response (may be wrapped in markdown code fences)
    cleaned = content.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)

    try:
        data = json.loads(cleaned)
        if isinstance(data, dict) and data.get("action") == "generate_report":
            return data
    except (json.JSONDecodeError, TypeError):
        pass
    return None
