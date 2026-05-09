import httpx
import json
import re

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "phi4-mini:latest"

CLASSIFICATION_PROMPT = """You are an email classification assistant. Classify the following email into exactly one of these categories:

- spam: unsolicited bulk email, phishing, scams
- newsletter: marketing emails, subscriptions, promotional content
- notification: automated system notifications, receipts, confirmations
- action_required: emails that need a reply or decision from the user
- fyi: emails that are informational but need no response
- calendar: meeting invites, scheduling requests

Respond with JSON only. No explanation. Format:
{{"classification": "<category>", "confidence": <0.0 to 1.0>, "reason": "<one sentence>"}}

Email:
Subject: {subject}
From: {sender}
Body: {body}"""

async def classify_email(subject: str, sender: str, body: str) -> dict:
    prompt = CLASSIFICATION_PROMPT.format(
        subject=subject,
        sender=sender,
        body=body[:2000]
    )
    try:
        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.post(OLLAMA_URL, json={
                "model": OLLAMA_MODEL,
                "prompt": prompt,
                "stream": False
            })
            response.raise_for_status()
            raw = response.json().get("response", "")
            match = re.search(r'\{.*\}', raw, re.DOTALL)
            if match:
                result = json.loads(match.group())
                result["confidence"] = float(result.get("confidence", 0.0))
                return result
            return {"classification": "unknown", "confidence": 0.0, "reason": "Could not parse model response"}
    except Exception as e:
        return {"classification": "error", "confidence": 0.0, "reason": str(e)}
