import os
import httpx
import logging
from dotenv import load_dotenv
from agent.drafter import ANTHROPIC_URL, CLAUDE_MODEL

load_dotenv("config/.env")

logger = logging.getLogger(__name__)

_AUTHOR_NAME = os.getenv("AUTHOR_NAME", "the adviser")

MIN_BODY_CHARS = 300
MAX_EXAMPLES_FOR_SYNTHESIS = 50
MAX_CHARS_PER_EXAMPLE = 400

SYNTHESIS_PROMPT = """You are analysing {n} emails written by {author_name}, a financial planner and tax adviser in Queensland, Australia.

Based on these emails, produce a detailed voice and style profile. Cover:
- Tone and register — how formal, warm, direct he is
- Sentence structure — length, rhythm, use of prose vs lists
- How he opens and closes (beyond "Kind Regards")
- How he handles uncertainty, hedging, or sensitive topics
- Vocabulary and phrases he favours or avoids
- How he explains complex financial or tax concepts
- His relationship style with clients in writing

Write the profile as specific, actionable writing instructions (250–400 words). Focus only on patterns you observe consistently across multiple emails — ignore one-offs.

Emails:
{email_texts}"""


def strip_signature(body: str) -> str:
    lower = body.lower()
    for marker in ["kind regards", "regards,", "regards\n"]:
        idx = lower.rfind(marker)
        if idx != -1:
            return body[:idx].strip()
    return body.strip()


async def _synthesize(account: str, examples: list) -> str:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise Exception("ANTHROPIC_API_KEY not set")

    sample = examples[:MAX_EXAMPLES_FOR_SYNTHESIS]
    email_texts = "\n\n---\n\n".join(
        f"Subject: {ex['subject']}\n\n{ex['body'][:MAX_CHARS_PER_EXAMPLE]}"
        for ex in sample
    )

    prompt = SYNTHESIS_PROMPT.format(n=len(sample), email_texts=email_texts, author_name=_AUTHOR_NAME)

    async with httpx.AsyncClient(timeout=60.0) as client:
        response = await client.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": CLAUDE_MODEL,
                "max_tokens": 1024,
                "messages": [{"role": "user", "content": prompt}]
            }
        )
        response.raise_for_status()
        return response.json()["content"][0]["text"].strip()


async def build_voice_profiles() -> dict:
    """
    Full pipeline: fetch sent emails → filter → store → synthesize → save voice profile.
    Returns counts per account.
    """
    from connectors.graph import get_sent_emails as graph_sent
    from connectors.gmail import get_sent_emails as gmail_sent
    from db.database import save_sent_example, get_all_sent_examples, save_voice_profile

    results = {}

    for account, fetch_fn in [("financial", lambda: graph_sent("financial", days=90)),
                               ("personal", lambda: graph_sent("personal", days=90)),
                               ("gmail", lambda: gmail_sent(days=90))]:
        logger.info(f"[{account}] Fetching sent emails...")
        try:
            emails = await fetch_fn()
            new = 0
            for email in emails:
                body = strip_signature(email["body"])
                if len(body) >= MIN_BODY_CHARS:
                    saved = await save_sent_example(
                        account=account,
                        message_id=email["id"],
                        subject=email["subject"],
                        body=body,
                        sent_at=email.get("sent_at"),
                    )
                    if saved:
                        new += 1
            logger.info(f"[{account}] {new} new examples stored")

            all_examples = await get_all_sent_examples(account)
            if not all_examples:
                logger.warning(f"[{account}] No examples available — skipping synthesis")
                results[account] = {"examples": 0, "profile": False}
                continue

            logger.info(f"[{account}] Synthesising voice profile from {len(all_examples)} examples...")
            profile = await _synthesize(account, all_examples)
            await save_voice_profile(account, profile, len(all_examples))
            logger.info(f"[{account}] Voice profile saved")
            results[account] = {"examples": len(all_examples), "profile": True}

        except Exception as e:
            logger.error(f"[{account}] Voice profile build failed: {e}")
            results[account] = {"examples": 0, "profile": False, "error": str(e)}

    return results
