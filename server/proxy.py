"""
MC & S CoWorker — Backend Classification Proxy
================================================
Lightweight FastAPI server that proxies email classification requests
to the Anthropic API. Runs on DigitalOcean so individual client
installs don't need their own API keys.
"""

import json
import os
import re

import anthropic
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

app = FastAPI(title="MC & S CoWorker Proxy", version="1.0.0")

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")


class RuleItem(BaseModel):
    category: str
    keywords: str


class ClassifyRequest(BaseModel):
    email_subject: str
    email_body: str
    rules: list[RuleItem]


class ClassifyResponse(BaseModel):
    category: str
    confidence: float
    reasoning: str
    sender_name: str


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/classify", response_model=ClassifyResponse)
def classify(req: ClassifyRequest):
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured on server.")

    categories_desc = "\n".join(
        f"- {r.category}: Keywords include: {r.keywords}"
        for r in req.rules
    )

    prompt = f"""You are an email classifier for an accounting firm.

Classify the email below into one of these categories, or OTHER if none fit:

{categories_desc}
- OTHER: Anything not listed above (meeting requests, complaints, ATO notices, etc.)

Also extract the sender's first name from any sign-off in the body. If not found, return "there".

Subject: {req.email_subject}
Body: {req.email_body[:1500]}

Respond ONLY with valid JSON:
{{
  "category": "EXAMPLE_CATEGORY",
  "confidence": 0.85,
  "reasoning": "Brief explanation.",
  "sender_name": "John"
}}"""

    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )

        text = response.content[0].text.strip()
        text = re.sub(r"```json\s*|```", "", text).strip()
        parsed = json.loads(text)

        # Normalise confidence to float
        confidence = parsed.get("confidence", 0.5)
        if isinstance(confidence, str):
            confidence_map = {"high": 0.9, "medium": 0.7, "low": 0.4}
            confidence = confidence_map.get(confidence.lower(), 0.5)

        return ClassifyResponse(
            category=parsed.get("category", "OTHER"),
            confidence=float(confidence),
            reasoning=parsed.get("reasoning", ""),
            sender_name=parsed.get("sender_name", "there"),
        )

    except json.JSONDecodeError as e:
        raise HTTPException(status_code=500, detail=f"Failed to parse classification response: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Classification failed: {e}")
