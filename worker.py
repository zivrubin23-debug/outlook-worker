#!/usr/bin/env python3
"""
My Secret Sex Toy Delivery — Automated Email Support Worker
Railway Cron Job — runs every 5 minutes

Rules:
- Only process emails received on or after May 11, 2026
- Emails older than that are skipped WITHOUT marking as read
- If Claude API fails, do NOT mark email as read
- Shopify contact form emails (mailer@shopify.com) are real customers
- Never auto-send return address — every return needs an RMA first
- Replies are sent as HTML for proper formatting
"""

import os
import json
import re
import requests
import anthropic
from datetime import datetime, timezone

# ── Configuration ─────────────────────────────────────────────────────────────

CLIENT_ID     = os.environ["MS365_MCP_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS365_MCP_CLIENT_SECRET"]
TENANT_ID     = os.environ["MS365_MCP_TENANT_ID"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
MAILBOX_USER  = os.environ["MAILBOX_USER"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

CUTOFF_DATE = datetime(2026, 5, 11, 0, 0, 0, tzinfo=timezone.utc)

SKIP_SENDERS = {
    "funnymee.com",
    "163.com",
    "bezlya.com",
    "info@mysecretsextoydelivery.com",
    "mailer@klaviyo.com",
    "hello.klaviyo.com",
    "shop.tiktok.com",
    "taylor@shop.tiktok.com",
    "sam@shop.tiktok.com",
    "partners@buddify.app",
    "brand.faire.com",
    "entervending.com",
    "sextoydistributing.com",
    "lovetoyus.com",
    "info-dysotoys.com",
    "secomapp.com",
    "clients.myguestlist.com.au",
    "swaysucker.com",
    "melodeem.com",
    "blq.com",
    "vscnovelty.com",
    "cucupie.com",
    "purepleasure-4.com",
    "fairyl.com",
    "foxmail.com",
    "adamssceptre.com",
    "getsweetums.com",
    "miamidistro1.com",
    "tabs.co",
    "swisstransfer.com",
}

# ── System Prompt ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are an automated customer support agent for My Secret Sex Toy Delivery
(mysecretsextoydelivery.com).

━━━ TONE & STYLE ━━━

- Warm, professional, and discreet
- Friendly but never overly casual
- Short paragraphs with a blank line between them for readability
- Always reply in the SAME LANGUAGE the customer wrote in
- Never mention internal systems, supplier names, or costs

━━━ STEP 1 — IS THIS A REAL CUSTOMER EMAIL? ━━━

Set is_customer_email = true ONLY if the email is clearly from a customer
about an existing or potential order. This includes:
- Order cancellation requests
- Shipping / tracking / delivery questions
- Address changes or corrections
- Return or refund requests
- Product questions from real shoppers
- Complaints about a received or missing item
- Questions about same-day / express delivery
- Questions about whether there is a physical store
- Shopify contact form messages (these are always real customers)
- Any follow-up in an ongoing customer conversation

Set is_customer_email = false (skip silently) for:
- Supplier offers, wholesale catalogs, manufacturer outreach
- Marketing newsletters or platform notifications
- Cold B2B outreach or partnership proposals
- Spam or automated system emails unrelated to a customer order
- Sex advice questions with no order involved

When in doubt, set is_customer_email = false.

━━━ STEP 2 — AUTO-SEND vs DRAFT ━━━

Set should_auto_send = true ONLY for these simple cases:
- Customer asking where their order is / tracking update
- Customer asking about delivery timing (same day, express)
- Customer asking if there is a physical store / pickup location
- Simple product question (what do you carry, do you have X)
- Address confirmation request (we reached out, they reply with address)
- Cancellation request (use the soft wording below)
- Order status inquiry with no complications

Set should_auto_send = false (save as DRAFT) for ALL of these:
- Any refund request or discussion
- Any return request (RMA must be arranged — never auto-send return address)
- Dispute, chargeback, or credit card claim
- Customer accusing the store of being fake, fraud, or a scam
- Legal threats or mentions of lawyers / consumer protection agencies
- Package marked delivered but customer says it never arrived
- Double charge or payment issues
- Very angry, aggressive, or threatening language
- Anything that requires checking order details or making a judgment call
- Anything unusual you are not fully confident handling automatically

━━━ KEY POLICIES ━━━

CANCELLATIONS:
  Use this soft wording — never say it is impossible:
  "We will do our best to help. Please note that we cannot guarantee
  cancellation at this stage, as the order may already be with the
  shipping provider. We will check and update you as soon as possible."

ORDER STATUS / WHERE IS MY ORDER:
  If no order number provided, ask for it politely.
  If tracking is available, share it. If not:
  "Your order has been passed to our shipping provider.
  As soon as we receive the tracking number, we will send it to you."

SAME-DAY / EXPRESS DELIVERY QUESTIONS:
  Explain clearly: "Same-day shipping means your order is dispatched
  the same day. Final delivery timing depends on the carrier's schedule."

RETURNS:
  NEVER include the return address in an auto-reply.
  Always say: "Please reply to this email and we will arrange a Return
  Merchandise Authorization (RMA) number and provide return instructions."
  Save as DRAFT so the team can issue the RMA.

REFUNDS:
  Always save as DRAFT. Never confirm a refund amount automatically.

NO PHYSICAL STORE:
  "We are an online-only store. All orders ship directly to your
  delivery address. We do not have a walk-in location."

━━━ EMAIL SIGNATURE ━━━

Always include this at the very end of every reply, exactly as shown:

___________________

Best regards,

John
Service Team Leader | MS

Online Store: www.mysecretsextoydelivery.com

━━━ OUTPUT FORMAT ━━━

Respond ONLY with valid JSON — no markdown, no extra text:
{
  "is_customer_email": true,
  "should_auto_send": true,
  "sensitivity_reason": "",
  "reply": "Full reply text here, with blank lines between short paragraphs, ending with the signature block exactly as shown above"
}
"""

# ── HTML Formatting ───────────────────────────────────────────────────────────

def format_reply_html(reply_text: str) -> str:
    """
    Convert plain-text reply into clean HTML.

    - Content before the signature is split on blank lines into <p> tags.
    - The signature block (starting with ___________________) uses <br> tags.
    - A horizontal rule separates body from signature.
    """
    SIGNATURE_MARKER = "___________________"

    # Split body from signature
    if SIGNATURE_MARKER in reply_text:
        body_part, sig_part = reply_text.split(SIGNATURE_MARKER, 1)
    else:
        body_part = reply_text
        sig_part  = ""

    # Convert body paragraphs (split on blank lines)
    body_html = ""
    paragraphs = re.split(r"\n{2,}", body_part.strip())
    for para in paragraphs:
        para = para.strip()
        if para:
            # Replace single newlines within a paragraph with <br>
            para = para.replace("\n", "<br>")
            body_html += f"<p>{para}</p>\n"

    # Convert signature lines to <br>-separated block
    sig_html = ""
    if sig_part:
        sig_lines = sig_part.strip().split("\n")
        sig_html_lines = []
        for line in sig_lines:
            escaped = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            sig_html_lines.append(escaped)
        sig_inner = "<br>".join(sig_html_lines)
        sig_html = (
            f'<hr style="border:none;border-top:1px solid #ccc;margin:16px 0;">\n'
            f'<p style="color:#555;font-size:13px;line-height:1.6;">'
            f'{sig_inner}'
            f'</p>\n'
        )

    html = (
        '<!DOCTYPE html><html><body '
        'style="font-family:Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">\n'
        f'{body_html}'
        f'{sig_html}'
        '</body></html>'
    )
    return html


# ── Microsoft Graph helpers ───────────────────────────────────────────────────

def get_token() -> str:
    resp = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
        },
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def graph(method: str, path: str, token: str, **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    return requests.request(
        method,
        f"{GRAPH_BASE}{path}",
        headers=headers,
        timeout=20,
        **kwargs,
    )


def get_unread_inbox(token: str) -> list:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/inbox/messages", token, params={
        "$filter": "isRead eq false",
        "$select": "id,subject,from,body,bodyPreview,receivedDateTime,conversationId",
        "$orderby": "receivedDateTime asc",
        "$top":     50,
    })
    resp.raise_for_status()
    return resp.json().get("value", [])


def already_replied(token: str, conversation_id: str, received_dt: str) -> bool:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/sentitems/messages", token, params={
        "$filter": f"conversationId eq '{conversation_id}'",
        "$select": "sentDateTime",
        "$top":    10,
    })
    if not resp.ok:
        return False
    for msg in resp.json().get("value", []):
        if msg["sentDateTime"] > received_dt:
            return True
    return False


def draft_exists(token: str, conversation_id: str) -> bool:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/drafts/messages", token, params={
        "$filter": f"conversationId eq '{conversation_id}'",
        "$select": "id",
        "$top":    5,
    })
    if not resp.ok:
        return False
    return len(resp.json().get("value", [])) > 0


def mark_read(token: str, message_id: str):
    graph("PATCH", f"/users/{MAILBOX_USER}/messages/{message_id}", token,
          headers={"Content-Type": "application/json"},
          json={"isRead": True})


def create_reply_draft(token: str, message_id: str) -> str:
    """Create a blank reply draft and return its draft message ID."""
    r = graph("POST", f"/users/{MAILBOX_USER}/messages/{message_id}/createReply", token,
              headers={"Content-Type": "application/json"},
              json={})
    r.raise_for_status()
    return r.json()["id"]


def update_draft_body(token: str, draft_id: str, html_body: str):
    """Update the draft's body with HTML content."""
    r = graph("PATCH", f"/users/{MAILBOX_USER}/messages/{draft_id}", token,
              headers={"Content-Type": "application/json"},
              json={
                  "body": {
                      "contentType": "html",
                      "content": html_body,
                  }
              })
    r.raise_for_status()


def send_draft(token: str, draft_id: str):
    """Send an existing draft message."""
    r = graph("POST", f"/users/{MAILBOX_USER}/messages/{draft_id}/send", token,
              headers={"Content-Type": "application/json"})
    r.raise_for_status()


def send_reply_html(token: str, message_id: str, reply_text: str):
    """
    Full 3-step HTML send:
    1. Create reply draft
    2. Update draft body as HTML
    3. Send the draft
    """
    html_body  = format_reply_html(reply_text)
    draft_id   = create_reply_draft(token, message_id)
    update_draft_body(token, draft_id, html_body)
    send_draft(token, draft_id)


def save_reply_draft(token: str, message_id: str, reply_text: str):
    """
    Create a draft for owner review (sensitive cases).
    Uses same 2-step flow but does NOT send.
    """
    html_body = format_reply_html(reply_text)
    draft_id  = create_reply_draft(token, message_id)
    update_draft_body(token, draft_id, html_body)


# ── Utilities ─────────────────────────────────────────────────────────────────

def strip_html(html: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;",  "&", text)
    text = re.sub(r"&lt;",   "<", text)
    text = re.sub(r"&gt;",   ">", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def parse_dt(dt_str: str) -> datetime:
    return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))


# ── Claude integration ────────────────────────────────────────────────────────

def ask_claude(subject: str, body: str, sender: str) -> dict:
    """Call Claude API. Raises exception on failure — caller must NOT mark email as read."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=1500,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": (
                f"From: {sender}\n"
                f"Subject: {subject}\n\n"
                f"Body:\n{body}"
            ),
        }],
    )

    raw = message.content[0].text.strip()

    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.lower().startswith("json"):
            raw = raw[4:]

    return json.loads(raw.strip())


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    ts = datetime.now(timezone.utc).isoformat()
    print(f"\n{'='*60}")
    print(f"Email Worker — {ts}")
    print(f"Mailbox: {MAILBOX_USER}")
    print(f"Cutoff:  {CUTOFF_DATE.isoformat()}")
    print(f"{'='*60}")

    token  = get_token()
    emails = get_unread_inbox(token)
    print(f"Unread emails found: {len(emails)}")

    for email in emails:
        sender_obj     = email["from"]["emailAddress"]
        sender_address = sender_obj.get("address", "").lower()
        sender_domain  = sender_address.split("@")[-1] if "@" in sender_address else ""
        subject        = email.get("subject") or "(no subject)"
        body_obj       = email.get("body", {})
        body_raw       = body_obj.get("content", email.get("bodyPreview", ""))
        body_text      = strip_html(body_raw) if body_obj.get("contentType") == "html" else body_raw
        conv_id        = email["conversationId"]
        msg_id         = email["id"]
        received_dt    = email["receivedDateTime"]

        print(f"\n-> '{subject}' | from: {sender_address} | received: {received_dt}")

        # ── Skip emails older than cutoff (do NOT mark as read) ──
        msg_time = parse_dt(received_dt)
        if msg_time < CUTOFF_DATE:
            print("  Skipped — older than cutoff (not marked read)")
            continue

        # ── Skip non-customer senders (mark as read to clean inbox) ──
        if sender_domain in SKIP_SENDERS or sender_address in SKIP_SENDERS:
            print("  Skipped — non-customer sender")
            mark_read(token, msg_id)
            continue

        # ── Skip if already replied ──
        if already_replied(token, conv_id, received_dt):
            print("  Skipped — already replied")
            mark_read(token, msg_id)
            continue

        # ── Skip if draft already exists for this conversation ──
        if draft_exists(token, conv_id):
            print("  Skipped — draft already exists")
            continue

        # ── Ask Claude (do NOT mark as read if this fails) ──
        try:
            result = ask_claude(subject, body_text, sender_address)
        except Exception as exc:
            print(f"  ERROR — Claude API failed: {exc}")
            print("  Email NOT marked as read (will retry next cycle)")
            continue

        if not result.get("is_customer_email", True):
            print("  Skipped — Claude: not a customer email")
            mark_read(token, msg_id)
            continue

        reply_text = result.get("reply", "")

        if result.get("should_auto_send"):
            try:
                send_reply_html(token, msg_id, reply_text)
                mark_read(token, msg_id)
                print("  AUTO-REPLY SENT (HTML)")
            except Exception as exc:
                print(f"  ERROR — failed to send reply: {exc}")
        else:
            try:
                save_reply_draft(token, msg_id, reply_text)
                reason = result.get("sensitivity_reason", "requires review")
                print(f"  DRAFT SAVED (HTML) — reason: {reason}")
            except Exception as exc:
                print(f"  ERROR — failed to create draft: {exc}")

    print(f"\nWorker finished — {datetime.now(timezone.utc).isoformat()}")


if __name__ == "__main__":
    main()