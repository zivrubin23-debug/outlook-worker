#!/usr/bin/env python3
"""
My Secret Sex Toy Delivery — Automated Email Support Worker
Deployed as a Railway Cron Job (every 5 minutes)

Logic:
  - Fetches unread customer emails from Outlook inbox
  - Skips supplier / internal / already-processed emails
  - Sends each email to Claude for a response decision
  - Simple cases  -> auto-reply sent immediately
  - Sensitive cases -> saved as Outlook draft for owner approval
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
MAILBOX_USER  = os.environ["MAILBOX_USER"]  # e.g. info@mysecretsextoydelivery.com

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Senders / domains to always skip (suppliers, spam, self)
SKIP_SENDERS = {
    "funnymee.com",
    "163.com",
    "info@mysecretsextoydelivery.com",
}

# ── System prompt for Claude ──────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are an automated customer support agent for My Secret Sex Toy Delivery
(mysecretsextoydelivery.com). You handle all incoming customer emails.

BRAND TONE:
Professional, warm, discreet, and helpful. Friendly but not overly casual.
Always reply in the SAME LANGUAGE the customer used.

KEY POLICIES:

CANCELLATIONS:
  Once an order is placed it is immediately transferred to our shipping provider
  and CANNOT be cancelled. Always explain this clearly. Offer return/refund
  instructions for after the package is received.

ORDER STATUS / WHERE IS MY ORDER:
  If the customer did not include an order number, ask them for it politely.

RETURNS:
  A 20% disposal / restocking fee applies. The customer ships the item back
  at their own cost. Once received we process the refund.

REFUNDS:
  Allow 5-10 business days after the return is received. Explain calmly.

PRODUCT QUESTIONS:
  Answer helpfully and discreetly.

EMAIL SIGNATURE — always include at the very end:

___________________

Best regards,

John
Service Team Leader | MS

Online Store: www.mysecretsextoydelivery.com

SENSITIVE CASES — must be flagged as draft, NEVER auto-sent:
Flag should_auto_send = false if the email contains ANY of:
  - Refund disputes or disagreements with a previous decision
  - Legal threats (lawyers, lawsuits, consumer protection agencies)
  - Chargeback or payment-dispute mentions
  - Very angry, aggressive, or threatening language
  - Complaints that require investigation or a judgment call
  - Anything unusual that you are not confident handling automatically

NON-CUSTOMER EMAILS:
Set is_customer_email = false for: supplier offers, spam, automated
notifications, or any email that is clearly not from a real customer.

OUTPUT FORMAT:
Respond ONLY with valid JSON — no markdown, no extra text:
{
  "is_customer_email": true,
  "should_auto_send": true,
  "sensitivity_reason": "",
  "reply": "Full reply text here, including signature"
}
"""

# ── Microsoft Graph helpers ───────────────────────────────────────────────────

def get_token() -> str:
    """Obtain an app-only access token via client credentials."""
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
    """Generic Microsoft Graph request."""
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
        "$top":     25,
    })
    resp.raise_for_status()
    return resp.json().get("value", [])


def already_replied(token: str, conversation_id: str, received_dt: str) -> bool:
    """Return True if we sent a message in this conversation after received_dt."""
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
    """Return True if a draft already exists for this conversation."""
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


def send_reply(token: str, message_id: str, reply_text: str):
    r = graph("POST", f"/users/{MAILBOX_USER}/messages/{message_id}/reply", token,
              headers={"Content-Type": "application/json"},
              json={"comment": reply_text})
    r.raise_for_status()


def create_draft_reply(token: str, message_id: str, reply_text: str):
    r = graph("POST", f"/users/{MAILBOX_USER}/messages/{message_id}/createReply", token,
              headers={"Content-Type": "application/json"},
              json={"comment": reply_text})
    r.raise_for_status()


# ── Utilities ─────────────────────────────────────────────────────────────────

def strip_html(html: str) -> str:
    """Very lightweight HTML to plain-text conversion."""
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;",  "&", text)
    text = re.sub(r"&lt;",   "<", text)
    text = re.sub(r"&gt;",   ">", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ── Claude integration ────────────────────────────────────────────────────────

def ask_claude(subject: str, body: str, sender: str) -> dict:
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

    # Strip accidental markdown code fences
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
    print(f"{'='*60}")

    token  = get_token()
    emails = get_unread_inbox(token)
    print(f"Unread emails found: {len(emails)}")

    for email in emails:
        sender_address = email["from"]["emailAddress"]["address"].lower()
        sender_domain  = sender_address.split("@")[-1]
        subject        = email.get("subject", "(no subject)")
        body_obj       = email.get("body", {})
        body_raw       = body_obj.get("content", email.get("bodyPreview", ""))
        body_text      = strip_html(body_raw) if body_obj.get("contentType") == "html" else body_raw
        conv_id        = email["conversationId"]
        msg_id         = email["id"]
        received_dt    = email["receivedDateTime"]

        print(f"\n-> '{subject}' | from: {sender_address}")

        # Skip non-customer senders
        if sender_domain in SKIP_SENDERS or sender_address in SKIP_SENDERS:
            print("  Skipped — non-customer sender")
            mark_read(token, msg_id)
            continue

        # Skip if already replied
        if already_replied(token, conv_id, received_dt):
            print("  Skipped — already replied")
            mark_read(token, msg_id)
            continue

        # Skip if draft already exists
        if draft_exists(token, conv_id):
            print("  Skipped — draft already exists")
            continue

        # Ask Claude
        try:
            result = ask_claude(subject, body_text, sender_address)
        except Exception as exc:
            print(f"  ERROR - Claude: {exc}")
            continue

        if not result.get("is_customer_email", True):
            print("  Skipped — Claude: not a customer email")
            mark_read(token, msg_id)
            continue

        reply_text = result.get("reply", "")

        if result.get("should_auto_send"):
            try:
                send_reply(token, msg_id, reply_text)
                mark_read(token, msg_id)
                print("  AUTO-REPLY SENT")
            except Exception as exc:
                print(f"  ERROR sending reply: {exc}")
        else:
            try:
                create_draft_reply(token, msg_id, reply_text)
                reason = result.get("sensitivity_reason", "sensitive case")
                print(f"  DRAFT SAVED — reason: {reason}")
            except Exception as exc:
                print(f"  ERROR creating draft: {exc}")

    print(f"\nWorker finished — {datetime.now(timezone.utc).isoformat()}")


if __name__ == "__main__":
    main()
