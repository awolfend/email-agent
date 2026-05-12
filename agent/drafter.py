import os
import httpx
import logging
import anthropic

logger = logging.getLogger(__name__)

CLAUDE_MODEL = "claude-sonnet-4-6"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODEL = "gpt-4o"

DEFAULT_PROMPT_FINANCIAL = """You are drafting an email reply on behalf of Anthony Wolfenden, a financial planner based in Queensland, Australia.

Anthony runs ***FINANCIAL_BUSINESS_NAME***, specialising in superannuation, investments, insurance, cashflow management, and retirement planning. He is an Authorised Representative (AR ***AR_NUMBER***) operating under an AFSL.

Anthony's voice and style — follow this precisely:

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

DEFAULT_PROMPT_GMAIL = """You are drafting an email reply on behalf of Anthony Wolfenden, a financial (tax) adviser based in Queensland, Australia.

Anthony runs ***TAX_BUSINESS_NAME***, specialising in income tax returns, tax planning, BAS lodgements, SMSF tax compliance, and business structuring. He is a registered Tax Agent (TRN ***TAX_TRN_NUMBER***) and Authorised Representative (AR ***AR_NUMBER***).

Anthony's voice and style — follow this precisely:

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

DEFAULT_PROMPT_PERSONAL = """You are drafting an email reply on behalf of Anthony Wolfenden for his personal email account (***PERSONAL_EMAIL***).

This account handles personal correspondence — family, friends, personal services, subscriptions, and matters unrelated to Anthony's professional businesses.

Anthony's voice and style — follow this precisely:

Tone: Warm and direct. Natural and human. No corporate language. No compliance framing. Relaxed but not sloppy.

Sentences: Short and clear. One idea per sentence. Conversational rhythm — reads like something a real person wrote, not a template.

Structure: Prose always. No bullet points, no numbered lists unless genuinely needed. Keep it brief — personal emails rarely need more than a few sentences.

Content: Get to the point quickly. Acknowledge the human moment first when one exists. Match the register of the incoming email — if someone writes casually, reply casually.

Sign-off: "Kind Regards" for professional-personal (accountants, lawyers, services). "Thanks" or "Cheers" acceptable for genuine personal correspondence. Never "Warm regards" or "Best".

Rules:
- No compliance disclaimers or financial advice caveats — this is personal, not professional.
- Do not include a signature block — it is added automatically.
- Do not include a subject line."""

VOICE_PROFILE_BLOCK = """Style profile derived from Anthony's own sent emails — follow this in addition to the instructions above:

{profile}

"""

USER_TEMPLATE = """Email to reply to:
Account: {account}
From: {sender}
Subject: {subject}
Body:
{body}

Draft the reply now. Write only the email body.{guidance_block}"""


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
        from db.database import get_setting
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
        from db.database import get_setting
        stored = await get_setting(f"footer_{account}")
        return stored or ""
    except Exception as e:
        logger.warning(f"Could not load footer setting: {e}")
        return ""


async def _call_claude(system_prompt: str, user_content: str) -> str:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise Exception("ANTHROPIC_API_KEY not set")
    client = anthropic.AsyncAnthropic(api_key=api_key)
    message = await client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=1024,
        system=[{
            "type": "text",
            "text": system_prompt,
            "cache_control": {"type": "ephemeral"},
        }],
        messages=[{"role": "user", "content": user_content}],
    )
    return message.content[0].text.strip()


async def _call_openai(system_prompt: str, user_content: str) -> str:
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
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_content},
                ],
            },
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()


async def generate_draft(account: str, sender: str, subject: str, body: str, guidance: str = "") -> str:
    voice_block = await _get_voice_block(account)
    base_prompt = await _get_base_prompt(account)
    guidance_block = f"\n\nAdditional instruction: {guidance.strip()}" if guidance.strip() else ""

    # System: writing instructions + voice profile (cached — same across all drafts per account)
    system_prompt = base_prompt
    if voice_block:
        system_prompt = voice_block + "\n" + system_prompt

    # User: the specific email to reply to
    user_content = USER_TEMPLATE.format(
        account=account,
        sender=sender,
        subject=subject,
        body=body[:4000],
        guidance_block=guidance_block,
    )

    draft = ""
    try:
        draft = await _call_claude(system_prompt, user_content)
        logger.info(f"Draft generated via Claude for: {subject[:50]}")
    except Exception as e:
        logger.warning(f"Claude failed ({e}), trying OpenAI fallback")
        try:
            draft = await _call_openai(system_prompt, user_content)
            logger.info(f"Draft generated via OpenAI for: {subject[:50]}")
        except Exception as e2:
            logger.error(f"OpenAI fallback also failed: {e2}")
            return ""

    footer = await _get_footer(account)
    if footer:
        draft = draft + "\n\n" + footer
    return draft
