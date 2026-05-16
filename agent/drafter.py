import os
import re
import json
import httpx
import logging
from dotenv import load_dotenv
from db.database import get_setting

_INVISIBLE_RE = re.compile(
    "[­​‌‍‎‏⁠﻿᠎]+"
)


def _sanitize_body(text: str) -> str:
    text = _INVISIBLE_RE.sub("", text)
    text = re.sub(r"[ \t]{3,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

load_dotenv("config/.env")

logger = logging.getLogger(__name__)

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
CLAUDE_MODEL = "claude-sonnet-4-6"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODEL = "gpt-4o"

_AUTHOR_NAME    = os.getenv("AUTHOR_NAME", "the adviser")
_AUTHOR_FIRST   = os.getenv("AUTHOR_FIRST_NAME", "the adviser")
_FINANCIAL_BIZ  = os.getenv("FINANCIAL_BUSINESS_NAME", "the financial planning business")
_TAX_BIZ        = os.getenv("TAX_BUSINESS_NAME", "the tax business")
_AR             = os.getenv("AR_NUMBER", "")
_TAX_TRN        = os.getenv("TAX_TRN_NUMBER", "")
_PERSONAL_EMAIL = os.getenv("PERSONAL_EMAIL", "")

DEFAULT_PROMPT_FINANCIAL = f"""You are drafting an email reply on behalf of {_AUTHOR_NAME}, a financial planner based in Queensland, Australia.

{_AUTHOR_FIRST} runs {_FINANCIAL_BIZ}, specialising in superannuation, investments, insurance, cashflow management, and retirement planning. He is an Authorised Representative (AR {_AR}) operating under an AFSL.

{_AUTHOR_FIRST}'s voice and style — follow this precisely:

Tone: Wise, calm, and direct. Genuinely on the client's side. Educational and open — he shares knowledge freely without using complexity as a moat. Conversational but professional. Not corporate. Not falsely warm.

Sentences: Short and precise. One idea per sentence. Do not run ideas together with "and" or "but" when a full stop would be cleaner. Each paragraph addresses one distinct topic only.

Structure: Prose by default. Use numbered steps or a short list only when there are genuinely parallel items that are clearer in list form — for example, three things the client needs to action, or four distinct risks to weigh up. Never use structure for decoration or to appear thorough.

Content: Lead with the human moment when one exists — acknowledgement before business. Take a clear position and state the reasoning plainly. Cite ATO references, ASIC rules, or legislation when relevant and useful. Flag what is not yet known and what needs to happen before advice can be formalised. Where factual context is relevant (market data, contribution caps, insurance structures), include it briefly and plainly.

Sign-off: Always "Kind Regards". Never "Warm regards", "Thanks", "Cheers", or "Best".

Compliance rules — non-negotiable:
- This is a regulated financial services environment. Never make specific recommendations in an email.
- Frame observations as "general considerations", "factual and general in nature", or "areas worth exploring".
- If a formal Statement of Advice (SOA) is required before recommending, say so.
- Never commit to specific outcomes, returns, or strategies without appropriate caveats.
- Do not include a signature block — it is added automatically.
- Do not include a subject line."""

DEFAULT_PROMPT_GMAIL = f"""You are drafting an email reply on behalf of {_AUTHOR_NAME}, a financial (tax) adviser based in Queensland, Australia.

{_AUTHOR_FIRST} runs {_TAX_BIZ}, specialising in income tax returns, tax planning, BAS lodgements, SMSF tax compliance, and business structuring. He is a registered Tax Agent (TRN {_TAX_TRN}) and Authorised Representative (AR {_AR}).

{_AUTHOR_FIRST}'s voice and style — follow this precisely:

Tone: Practical and clear. Focused on outcomes the client can act on. Educational without being condescending — he explains the why behind the numbers. Approachable and human, not stiff or overly formal.

Sentences: Short and precise. One idea per sentence. Do not run ideas together with "and" or "but" when a full stop would be cleaner. Each paragraph addresses one distinct topic only.

Structure: Prose by default. Use numbered steps or a short list only when there are genuinely parallel items that are clearer in list form — for example, documents needed, or sequential lodgement steps. Never use structure for decoration or to appear thorough.

Content: Lead with the human moment when one exists — acknowledgement before business. Take a clear position and state the reasoning plainly. Cite ATO references, tax rulings, or legislation when relevant — not to impress, because it helps. Flag what information is still required before lodging or advising. Where factual context is relevant (tax rates, deduction rules, ATO deadlines), include it briefly and plainly.

Sign-off: Always "Kind Regards". Never "Warm regards", "Thanks", "Cheers", or "Best".

Compliance rules — non-negotiable:
- This is a regulated tax agent environment. Never make specific recommendations without appropriate caveats.
- Frame observations as "general in nature", "subject to your specific circumstances", or "worth discussing further".
- Never commit to specific tax outcomes or refund amounts without reviewing all relevant documents.
- Do not include a signature block — it is added automatically.
- Do not include a subject line."""

DEFAULT_PROMPT_PERSONAL = f"""You are drafting an email reply on behalf of {_AUTHOR_NAME} for his personal email account ({_PERSONAL_EMAIL}).

This account handles personal correspondence — family, friends, personal services, subscriptions, and matters unrelated to {_AUTHOR_FIRST}'s professional businesses.

{_AUTHOR_FIRST}'s voice and style — follow this precisely:

Tone: Warm and direct. Natural and human. No corporate language. No compliance framing. Relaxed but not sloppy.

Sentences: Short and clear. One idea per sentence. Conversational rhythm — reads like something a real person wrote, not a template.

Structure: Prose always. No bullet points, no numbered lists unless genuinely needed. Keep it brief — personal emails rarely need more than a few sentences.

Content: Get to the point quickly. Acknowledge the human moment first when one exists. Match the register of the incoming email — if someone writes casually, reply casually.

Sign-off: "Kind Regards" for professional-personal (accountants, lawyers, services). "Thanks" or "Cheers" acceptable for genuine personal correspondence. Never "Warm regards" or "Best".

Rules:
- No compliance disclaimers or financial advice caveats — this is personal, not professional.
- Do not include a signature block — it is added automatically.
- Do not include a subject line."""

EMAIL_SUFFIX = """\n\n{crm_block}Email to reply to:
Account: {account}
From: {sender}
Subject: {subject}
Body:
{body}

Draft the reply now. Write only the email body.{guidance_block}"""

COMPOSE_SUFFIX = """\n\n{crm_block}Compose a new outgoing email (not a reply).

Recipient: {to}
User instructions: {prompt}{meeting_block}

Write only the email body. Do not include a subject line. Do not include a signature block — it is added automatically.{subject_line}"""

VOICE_PROFILE_BLOCK = f"""Style profile derived from {_AUTHOR_FIRST}'s own sent emails — follow this in addition to the instructions above:

{{profile}}

"""


async def _get_voice_block(account: str) -> str:
    try:
        from db.database import get_voice_profile
        profile = await get_voice_profile(account)
        if not profile:
            return ""
        return VOICE_PROFILE_BLOCK.format(profile=profile)
    except Exception as e:
        logger.warning(f"Could not load voice profile: {e}")
        return ""


async def _get_base_prompt(account: str) -> str:
    try:
        stored = await get_setting(f"prompt_{account}")
        if stored:
            return stored
    except Exception as e:
        logger.warning(f"Could not load prompt setting: {e}")
    if account == "financial":
        return DEFAULT_PROMPT_FINANCIAL
    if account == "personal":
        return DEFAULT_PROMPT_PERSONAL
    return DEFAULT_PROMPT_GMAIL


async def _get_footer(account: str) -> str:
    try:
        stored = await get_setting(f"footer_{account}")
        return stored or ""
    except Exception as e:
        logger.warning(f"Could not load footer setting: {e}")
        return ""


async def _call_claude(prompt: str) -> str:
    api_key = os.getenv("ANTHROPIC_API_KEY") or ANTHROPIC_API_KEY
    if not api_key:
        raise Exception("ANTHROPIC_API_KEY not set")
    async with httpx.AsyncClient(timeout=30.0) as client:
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
                "messages": [{"role": "user", "content": prompt}],
            },
        )
        response.raise_for_status()
        return response.json()["content"][0]["text"].strip()


async def _call_openai(prompt: str) -> str:
    api_key = os.getenv("OPENAI_API_KEY") or OPENAI_API_KEY
    if not api_key:
        raise Exception("OPENAI_API_KEY not set")
    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.post(
            OPENAI_URL,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": OPENAI_MODEL,
                "max_tokens": 1024,
                "messages": [{"role": "user", "content": prompt}],
            },
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()


async def generate_draft(account: str, sender: str, subject: str, body: str,
                         guidance: str = "", crm_context: str = "") -> str:
    voice_block = await _get_voice_block(account)
    base_prompt = await _get_base_prompt(account)
    guidance_block = f"\n\nAdditional instruction: {guidance.strip()}" if guidance.strip() else ""
    crm_block = crm_context.strip() + "\n\n" if crm_context.strip() else ""

    prompt = base_prompt + EMAIL_SUFFIX.format(
        account=account,
        sender=sender,
        subject=subject,
        body=_sanitize_body(body)[:4000],
        guidance_block=guidance_block,
        crm_block=crm_block,
    )
    if voice_block:
        prompt = voice_block + "\n" + prompt

    draft = ""
    last_claude_err = None
    for attempt in range(2):
        try:
            draft = await _call_claude(prompt)
            logger.info(f"Draft generated via Claude for: {subject[:50]}")
            break
        except Exception as e:
            last_claude_err = e
            if attempt == 0:
                await asyncio.sleep(2)
    if not draft:
        logger.warning(f"Claude failed after retry ({last_claude_err}), trying OpenAI fallback")
        try:
            draft = await _call_openai(prompt)
            logger.info(f"Draft generated via OpenAI for: {subject[:50]}")
        except Exception as e2:
            logger.error(f"OpenAI fallback also failed: {e2}")
            return ""

    footer = await _get_footer(account)
    if footer:
        draft = draft + "\n\n" + footer
    return draft


_SCHEDULING_EXTRACT_PROMPT = """Analyse this email and detect if the sender is making a scheduling request.

Today is {today} (Australia/Brisbane, UTC+10).

Return JSON only, no explanation.
Format: {{
  "is_scheduling": true,
  "proposed_times": ["verbatim phrase 1", "verbatim phrase 2"],
  "proposed_slots": [{{"start": "2026-05-20T14:00:00+10:00", "end": "2026-05-20T15:00:00+10:00"}}],
  "topic": "meeting purpose"
}}
Or if not a scheduling request: {{"is_scheduling": false}}

Rules:
- is_scheduling is true when the sender proposes specific times/dates OR asks for your availability
- proposed_times: verbatim time/date phrases from the email (empty array if none)
- proposed_slots: resolved ISO 8601 datetime ranges (UTC+10) for each proposed time:
  - Only include if the date can be determined from context (skip "sometime next week" etc.)
  - Assume 1-hour duration if no end time stated
  - "afternoon" = 14:00 start; "morning" = 09:00 start; "end of day" = 16:00 start
  - If only a day is given with no time, do NOT include in proposed_slots
- topic: brief meeting purpose (empty string if unclear)

Subject: {subject}
From: {sender}
Body: {body}"""


async def extract_scheduling_intent(subject: str, sender: str, body: str, today: str = "") -> dict:
    prompt = _SCHEDULING_EXTRACT_PROMPT.format(subject=subject, sender=sender, body=_sanitize_body(body)[:2000], today=today)
    try:
        raw = await _call_claude(prompt)
        match = re.search(r'\{.*\}', raw, re.DOTALL)
        if match:
            return json.loads(match.group())
    except Exception as e:
        logger.debug(f"extract_scheduling_intent failed: {e}")
    return {"is_scheduling": False, "proposed_times": [], "topic": ""}


async def generate_compose_draft(
    account: str,
    to: str,
    subject: str = "",
    prompt: str = "",
    crm_context: str = "",
    meeting_slots: list[str] = None,
    duration_minutes: int = 60,
) -> tuple[str, str]:
    """
    Generate a new outgoing email (not a reply).
    Returns (draft_body, subject_line).
    If subject is blank, Claude is asked to generate one embedded in the response.
    """
    voice_block = await _get_voice_block(account)
    base_prompt = await _get_base_prompt(account)
    crm_block = crm_context.strip() + "\n\n" if crm_context.strip() else ""

    # Strip the "do not include a subject line" rule from base prompt for compose
    stripped_base = base_prompt.replace(
        "- Do not include a subject line.", ""
    ).rstrip()

    # Build meeting block
    meeting_block = ""
    if meeting_slots:
        lines = [f"\n\nMeeting proposal — duration: {duration_minutes} minutes"]
        if len(meeting_slots) == 1:
            lines.append(f"Proposed time: {meeting_slots[0]}")
            lines.append("Present this single time in the email body. State that a calendar invite with meeting details has been sent.")
        else:
            lines.append("Proposed times (list these in the email body and ask the client to confirm which suits):")
            for i, s in enumerate(meeting_slots, 1):
                lines.append(f"  {i}. {s}")
        meeting_block = "\n".join(lines)

    # Subject generation — ask Claude to append it if blank
    need_subject = not subject.strip()
    subject_line = ""
    if need_subject:
        subject_line = '\n\nAfter the email body, on a new line write exactly:\nSUBJECT: <concise subject of 6 words or fewer>'

    full_prompt = stripped_base + COMPOSE_SUFFIX.format(
        crm_block=crm_block,
        to=to,
        prompt=prompt or "Write a professional email.",
        meeting_block=meeting_block,
        subject_line=subject_line,
    )
    if voice_block:
        full_prompt = voice_block + "\n" + full_prompt

    raw = ""
    last_claude_err = None
    for attempt in range(2):
        try:
            raw = await _call_claude(full_prompt)
            logger.info(f"Compose draft generated via Claude for: {to[:50]}")
            break
        except Exception as e:
            last_claude_err = e
            if attempt == 0:
                await asyncio.sleep(2)
    if not raw:
        logger.warning(f"Claude failed for compose after retry ({last_claude_err}), trying OpenAI fallback")
        try:
            raw = await _call_openai(full_prompt)
            logger.info(f"Compose draft generated via OpenAI for: {to[:50]}")
        except Exception as e2:
            logger.error(f"OpenAI compose fallback also failed: {e2}")
            return "", ""

    # Parse subject from tail if we asked for one
    parsed_subject = subject.strip()
    body = raw
    if need_subject:
        if "\nSUBJECT:" in raw:
            parts = raw.rsplit("\nSUBJECT:", 1)
            body = parts[0].strip()
            parsed_subject = parts[1].strip()
        elif raw.upper().startswith("SUBJECT:"):
            first_newline = raw.find("\n")
            if first_newline != -1:
                parsed_subject = raw[8:first_newline].strip()
                body = raw[first_newline:].strip()

    footer = await _get_footer(account)
    if footer:
        body = body + "\n\n" + footer

    return body, parsed_subject
