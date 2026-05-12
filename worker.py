#!/usr/bin/env python3
"""
My Secret Sex Toy Delivery — Automated Email Support Worker
Railway Cron Job — runs every 5 minutes

Two workflows:
  1. Customer support emails  → auto-reply or draft via Claude
  2. Supplier tracking emails → update Shopify fulfillment automatically
"""

import os
import json
import re
import requests
import anthropic
from datetime import datetime, timezone
from urllib.parse import urlparse, parse_qs, quote

# ── Configuration ─────────────────────────────────────────────────────────────

CLIENT_ID             = os.environ["MS365_MCP_CLIENT_ID"]
CLIENT_SECRET         = os.environ["MS365_MCP_CLIENT_SECRET"]
TENANT_ID             = os.environ["MS365_MCP_TENANT_ID"]
ANTHROPIC_KEY         = os.environ["ANTHROPIC_API_KEY"]
MAILBOX_USER          = os.environ["MAILBOX_USER"]
SHOPIFY_CLIENT_ID     = os.environ["SHOPIFY_CLIENT_ID"]
SHOPIFY_CLIENT_SECRET = os.environ["SHOPIFY_CLIENT_SECRET"]
SHOPIFY_LOCATION      = os.environ.get("SHOPIFY_LOCATION_ID", "")

SHOPIFY_STORE   = "40026a.myshopify.com"
SHOPIFY_VERSION = "2024-01"
GRAPH_BASE      = "https://graph.microsoft.com/v1.0"

CUTOFF_DATE = datetime(2026, 5, 11, 0, 0, 0, tzinfo=timezone.utc)

SUPPLIER_FOLDER_ID = "AAMkADBhNDBkMzJjLWRiZTMtNDE2NC1iNDQ1LTQ4ZjVlMWE3ZGFkYwAuAAAAAAC5ypHMZNNXRo-eNkzzCxTpAQCygDv4EiYcSqbyqT-RbGLGAALMrinTAAA="

TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

SKIP_SENDERS = {
    "funnymee.com", "163.com", "bezlya.com",
    "info@mysecretsextoydelivery.com",
    "mailer@klaviyo.com", "hello.klaviyo.com",
    "shop.tiktok.com", "taylor@shop.tiktok.com", "sam@shop.tiktok.com",
    "partners@buddify.app", "brand.faire.com", "entervending.com",
    "sextoydistributing.com", "lovetoyus.com", "info-dysotoys.com",
    "secomapp.com", "clients.myguestlist.com.au", "swaysucker.com",
    "melodeem.com", "blq.com", "vscnovelty.com", "cucupie.com",
    "purepleasure-4.com", "fairyl.com", "foxmail.com", "adamssceptre.com",
    "getsweetums.com", "miamidistro1.com", "tabs.co", "swisstransfer.com",
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
- Shopify contact form messages (always real customers)
- Any follow-up in an ongoing customer conversation

Set is_customer_email = false for:
- Supplier offers, wholesale catalogs, manufacturer outreach
- Marketing newsletters or platform notifications
- Cold B2B outreach or partnership proposals
- Spam or automated system emails
- Sex advice questions with no order involved

When in doubt, set is_customer_email = false.

━━━ STEP 2 — AUTO-SEND vs DRAFT ━━━

Set should_auto_send = true ONLY for:
- Order status / tracking questions
- Delivery timing questions (same day, express)
- Physical store / pickup questions
- Simple product questions
- Address confirmation replies
- Cancellation requests (use soft wording below)

Set should_auto_send = false (DRAFT) for:
- Any refund request or discussion
- Any return request (never auto-send address — RMA required)
- Dispute, chargeback, or credit card claim
- Fraud / scam accusations
- Legal threats
- Package delivered but not received
- Double charge or payment issues
- Angry or aggressive language
- Anything requiring judgment or investigation

━━━ KEY POLICIES ━━━

CANCELLATIONS:
  Soft wording only — never say impossible:
  "We will do our best to help. Please note that we cannot guarantee
  cancellation at this stage, as the order may already be with the
  shipping provider. We will check and update you as soon as possible."

ORDER STATUS:
  If no order number provided, ask politely.
  If no tracking yet: "Your order has been passed to our shipping provider.
  As soon as we receive the tracking number, we will send it to you."

RETURNS:
  NEVER include return address. Always say:
  "Please reply and we will arrange an RMA number and return instructions."
  Save as DRAFT.

REFUNDS: Always DRAFT. Never confirm amounts automatically.

NO PHYSICAL STORE:
  "We are an online-only store. All orders ship directly to your address."

━━━ EMAIL SIGNATURE ━━━

Always include at the very end:

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
  "reply": "Full reply text, blank lines between paragraphs, ending with signature"
}
"""

# ── HTML Formatting ───────────────────────────────────────────────────────────

def format_reply_html(reply_text: str) -> str:
    SIGNATURE_MARKER = "___________________"

    if SIGNATURE_MARKER in reply_text:
        body_part, sig_part = reply_text.split(SIGNATURE_MARKER, 1)
    else:
        body_part = reply_text
        sig_part  = ""

    body_html = ""
    for para in re.split(r"\n{2,}", body_part.strip()):
        para = para.strip()
        if para:
            body_html += f"<p>{para.replace(chr(10), '<br>')}</p>\n"

    sig_html = ""
    if sig_part:
        sig_inner = "<br>".join(
            l.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            for l in sig_part.strip().split("\n")
        )
        sig_html = (
            '<hr style="border:none;border-top:1px solid #ccc;margin:16px 0;">\n'
            '<p style="color:#555;font-size:13px;line-height:1.6;">'
            f'{sig_inner}</p>\n'
        )

    return (
        '<!DOCTYPE html><html><body '
        'style="font-family:Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">\n'
        f'{body_html}{sig_html}</body></html>'
    )

# ── Notifications ─────────────────────────────────────────────────────────────

def notify(message: str):
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        return
    try:
        requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            json={"chat_id": TELEGRAM_CHAT_ID, "text": message, "parse_mode": "HTML"},
            timeout=10,
        )
    except Exception:
        pass

# ── Microsoft Graph helpers ───────────────────────────────────────────────────

def get_ms_token() -> str:
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
        method, f"{GRAPH_BASE}{path}",
        headers=headers, timeout=20, **kwargs,
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


def get_unread_supplier_emails(token: str) -> list:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/{SUPPLIER_FOLDER_ID}/messages", token, params={
        "$filter": "isRead eq false",
        "$select": "id,subject,from,body,bodyPreview,receivedDateTime",
        "$orderby": "receivedDateTime asc",
        "$top":     25,
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
    return any(m["sentDateTime"] > received_dt for m in resp.json().get("value", []))


def draft_exists(token: str, conversation_id: str) -> bool:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/drafts/messages", token, params={
        "$filter": f"conversationId eq '{conversation_id}'",
        "$select": "id", "$top": 5,
    })
    if not resp.ok:
        return False
    return len(resp.json().get("value", [])) > 0


def mark_read(token: str, message_id: str):
    graph("PATCH", f"/users/{MAILBOX_USER}/messages/{message_id}", token,
          headers={"Content-Type": "application/json"},
          json={"isRead": True})


def create_reply_draft(token: str, message_id: str) -> str:
    r = graph("POST", f"/users/{MAILBOX_USER}/messages/{message_id}/createReply", token,
              headers={"Content-Type": "application/json"}, json={})
    r.raise_for_status()
    return r.json()["id"]


def update_draft_body(token: str, draft_id: str, html_body: str):
    r = graph("PATCH", f"/users/{MAILBOX_USER}/messages/{draft_id}", token,
              headers={"Content-Type": "application/json"},
              json={"body": {"contentType": "html", "content": html_body}})
    r.raise_for_status()


def send_draft(token: str, draft_id: str):
    r = graph("POST", f"/users/{MAILBOX_USER}/messages/{draft_id}/send", token,
              headers={"Content-Type": "application/json"})
    r.raise_for_status()


def send_reply_html(token: str, message_id: str, reply_text: str):
    draft_id = create_reply_draft(token, message_id)
    update_draft_body(token, draft_id, format_reply_html(reply_text))
    send_draft(token, draft_id)


def save_reply_draft(token: str, message_id: str, reply_text: str):
    draft_id = create_reply_draft(token, message_id)
    update_draft_body(token, draft_id, format_reply_html(reply_text))

# ── Shopify Auth ──────────────────────────────────────────────────────────────

def get_shopify_token() -> str:
    resp = requests.post(
        f"https://{SHOPIFY_STORE}/admin/oauth/access_token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     SHOPIFY_CLIENT_ID,
            "client_secret": SHOPIFY_CLIENT_SECRET,
        },
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]

# ── Shopify API helpers ───────────────────────────────────────────────────────

def shopify_get(path: str, shopify_token: str) -> dict:
    resp = requests.get(
        f"https://{SHOPIFY_STORE}/admin/api/{SHOPIFY_VERSION}/{path}",
        headers={
            "X-Shopify-Access-Token": shopify_token,
            "Content-Type": "application/json",
        },
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()


def shopify_post(path: str, data: dict, shopify_token: str) -> dict:
    resp = requests.post(
        f"https://{SHOPIFY_STORE}/admin/api/{SHOPIFY_VERSION}/{path}",
        headers={
            "X-Shopify-Access-Token": shopify_token,
            "Content-Type": "application/json",
        },
        json=data, timeout=15,
    )
    resp.raise_for_status()
    return resp.json()

# ── Shopify fulfillment helpers ───────────────────────────────────────────────

def detect_carrier(url: str) -> str:
    u = url.lower()
    if "usps.com"   in u: return "USPS"
    if "fedex.com"  in u: return "FedEx"
    if "ups.com"    in u: return "UPS"
    if "ontrac.com" in u: return "OnTrac"
    return "Other"


def extract_tracking_number(tracking_url: str, carrier: str) -> str:
    try:
        params = parse_qs(urlparse(tracking_url).query)
        mapping = {
            "USPS":   ["tLabels"],
            "FedEx":  ["trknbr", "tracknumbers"],
            "UPS":    ["tracknum", "tracknumbers"],
            "OnTrac": ["number"],
        }
        for key in mapping.get(carrier, []):
            if key in params:
                return params[key][0]
    except Exception:
        pass
    return ""


def extract_supplier_email_data(body_text: str) -> dict:
    name_match = re.search(r"Hi\s+(.+?),\s+your order has shipped", body_text, re.IGNORECASE)
    customer_name = name_match.group(1).strip() if name_match else ""

    tracking_url = ""
    carrier_domains = [
        "tools.usps.com", "usps.com/go",
        "fedex.com/fedextrack", "fedex.com/tracking",
        "ups.com/track", "ontrac.com/tracking",
    ]
    for url in re.findall(r'https?://[^\s\)\>\"\'\<]+', body_text):
        if any(d in url.lower() for d in carrier_domains):
            tracking_url = url.strip().rstrip(".")
            break

    return {"customer_name": customer_name, "tracking_url": tracking_url}


def find_unfulfilled_order(customer_name: str, shopify_token: str) -> dict:
    try:
        customers = shopify_get(
            f"customers/search.json?query={quote(customer_name)}&limit=10&fields=id,first_name,last_name",
            shopify_token,
        ).get("customers", [])

        if not customers:
            return {}

        name_lower = customer_name.strip().lower()
        matched_id = None

        for c in customers:
            full = f"{c.get('first_name','')} {c.get('last_name','')}".strip().lower()
            if full == name_lower:
                matched_id = c["id"]
                break

        if not matched_id:
            matched_id = customers[0]["id"]

        orders = shopify_get(
            f"orders.json?customer_id={matched_id}"
            f"&fulfillment_status=unfulfilled&status=open&limit=5",
            shopify_token,
        ).get("orders", [])

        return orders[0] if orders else {}

    except Exception as exc:
        print(f"   ERROR searching Shopify: {exc}")
        return {}


def create_fulfillment(order_id: int, tracking_number: str,
                       tracking_url: str, carrier: str,
                       shopify_token: str) -> bool:
    try:
        fulfillment_orders = shopify_get(
            f"orders/{order_id}/fulfillment_orders.json",
            shopify_token,
        ).get("fulfillment_orders", [])

        open_fos = [fo for fo in fulfillment_orders if fo["status"] == "open"]
        if not open_fos:
            print(f"   No open fulfillment orders for order {order_id}")
            return False

        result = shopify_post("fulfillments.json", {
            "fulfillment": {
                "line_items_by_fulfillment_order": [
                    {"fulfillment_order_id": open_fos[0]["id"]}
                ],
                "tracking_info": {
                    "number":  tracking_number,
                    "url":     tracking_url,
                    "company": carrier,
                },
                "notify_customer": True,
            }
        }, shopify_token)

        return "fulfillment" in result

    except Exception as exc:
        print(f"   ERROR creating fulfillment: {exc}")
        return False

# ── Utilities ─────────────────────────────────────────────────────────────────

def strip_html(html: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    for ent, char in [("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">")]:
        text = text.replace(ent, char)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def parse_dt(dt_str: str) -> datetime:
    return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))

# ── Claude ────────────────────────────────────────────────────────────────────

def ask_claude(subject: str, body: str, sender: str) -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=1500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"From: {sender}\nSubject: {subject}\n\nBody:\n{body}"}],
    )
    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.lower().startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())

# ── Workflow 1: Supplier tracking emails ──────────────────────────────────────

def process_supplier_emails(ms_token: str, shopify_token: str):
    print(f"\n{'─'*60}")
    print("WORKFLOW 1 — Supplier tracking emails")
    print(f"{'─'*60}")

    emails = get_unread_supplier_emails(ms_token)
    print(f"Unread supplier emails: {len(emails)}")

    for email in emails:
        subject     = email.get("subject", "")
        body_obj    = email.get("body", {})
        body_raw    = body_obj.get("content", email.get("bodyPreview", ""))
        body_text   = strip_html(body_raw) if body_obj.get("contentType") == "html" else body_raw
        msg_id      = email["id"]
        received_dt = email["receivedDateTime"]

        print(f"\n-> '{subject}' | {received_dt}")

        if parse_dt(received_dt) < CUTOFF_DATE:
            print("   Skipped — older than cutoff (not marked read)")
            continue

        data          = extract_supplier_email_data(body_text)
        customer_name = data["customer_name"]
        tracking_url  = data["tracking_url"]

        if not customer_name:
            print("   Could not extract customer name — not marked read")
            continue

        if not tracking_url:
            print(f"   No tracking URL found for '{customer_name}' — not marked read")
            continue

        carrier         = detect_carrier(tracking_url)
        tracking_number = extract_tracking_number(tracking_url, carrier)

        print(f"   Customer:  {customer_name}")
        print(f"   Carrier:   {carrier}")
        print(f"   Tracking:  {tracking_number}")

        order = find_unfulfilled_order(customer_name, shopify_token)
        if not order:
            print(f"   WARNING — No unfulfilled order found for '{customer_name}'")
            print("   Not marking as read — needs manual check")
            notify(
                f"⚠️ <b>Tracking not matched</b>\n"
                f"Customer: {customer_name}\n"
                f"Carrier: {carrier} | {tracking_number}\n"
                f"No unfulfilled Shopify order found — manual action needed"
            )
            continue

        order_name = order.get("name", "")
        order_id   = order["id"]
        print(f"   Shopify order: {order_name} (ID: {order_id})")

        success = create_fulfillment(order_id, tracking_number, tracking_url, carrier, shopify_token)

        if success:
            mark_read(ms_token, msg_id)
            print(f"   ✅ Fulfillment created — Shopify will notify customer")
            notify(
                f"📦 <b>Tracking updated</b>\n"
                f"Customer: {customer_name}\n"
                f"Order: {order_name}\n"
                f"Carrier: {carrier} | {tracking_number}"
            )
        else:
            print(f"   ❌ Fulfillment failed — not marking as read")
            notify(
                f"❌ <b>Fulfillment failed</b>\n"
                f"Customer: {customer_name} | Order: {order_name}\n"
                f"Manual action needed"
            )

# ── Workflow 2: Customer support emails ───────────────────────────────────────

def process_customer_emails(ms_token: str):
    print(f"\n{'─'*60}")
    print("WORKFLOW 2 — Customer support emails")
    print(f"{'─'*60}")

    emails = get_unread_inbox(ms_token)
    print(f"Unread inbox emails: {len(emails)}")

    for email in emails:
        sender_address = email["from"]["emailAddress"].get("address", "").lower()
        sender_domain  = sender_address.split("@")[-1] if "@" in sender_address else ""
        subject        = email.get("subject") or "(no subject)"
        body_obj       = email.get("body", {})
        body_raw       = body_obj.get("content", email.get("bodyPreview", ""))
        body_text      = strip_html(body_raw) if body_obj.get("contentType") == "html" else body_raw
        conv_id        = email["conversationId"]
        msg_id         = email["id"]
        received_dt    = email["receivedDateTime"]

        print(f"\n-> '{subject}' | from: {sender_address} | {received_dt}")

        if parse_dt(received_dt) < CUTOFF_DATE:
            print("   Skipped — older than cutoff (not marked read)")
            continue

        if sender_domain in SKIP_SENDERS or sender_address in SKIP_SENDERS:
            print("   Skipped — non-customer sender")
            mark_read(ms_token, msg_id)
            continue

        if already_replied(ms_token, conv_id, received_dt):
            print("   Skipped — already replied")
            mark_read(ms_token, msg_id)
            continue

        if draft_exists(ms_token, conv_id):
            print("   Skipped — draft already exists")
            continue

        try:
            result = ask_claude(subject, body_text, sender_address)
        except Exception as exc:
            print(f"   ERROR — Claude failed: {exc}")
            print("   Not marked as read (will retry)")
            continue

        if not result.get("is_customer_email", True):
            print("   Skipped — not a customer email")
            mark_read(ms_token, msg_id)
            continue

        reply_text = result.get("reply", "")

        if result.get("should_auto_send"):
            try:
                send_reply_html(ms_token, msg_id, reply_text)
                mark_read(ms_token, msg_id)
                print("   AUTO-REPLY SENT")
                notify(f"📨 <b>Auto-reply sent</b>\nTo: {sender_address}\nSubject: {subject}")
            except Exception as exc:
                print(f"   ERROR sending reply: {exc}")
        else:
            try:
                save_reply_draft(ms_token, msg_id, reply_text)
                reason = result.get("sensitivity_reason", "requires review")
                print(f"   DRAFT SAVED — {reason}")
                notify(f"📝 <b>Draft needs review</b>\nFrom: {sender_address}\nSubject: {subject}\nReason: {reason}")
            except Exception as exc:
                print(f"   ERROR saving draft: {exc}")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    ts = datetime.now(timezone.utc).isoformat()
    print(f"\n{'='*60}")
    print(f"Email Worker — {ts}")
    print(f"Mailbox:  {MAILBOX_USER}")
    print(f"Shopify:  {SHOPIFY_STORE}")
    print(f"Cutoff:   {CUTOFF_DATE.isoformat()}")
    print(f"Model:    claude-sonnet-4-5")
    print(f"{'='*60}")

    ms_token = get_ms_token()

    print("\nAuthenticating with Shopify...")
    try:
        shopify_token = get_shopify_token()
        print("Shopify token: OK")
    except Exception as exc:
        print(f"ERROR — Shopify auth failed: {exc}")
        print("Skipping supplier workflow. Customer emails will still be processed.")
        shopify_token = None

    if shopify_token:
        process_supplier_emails(ms_token, shopify_token)

    process_customer_emails(ms_token)

    print(f"\n{'='*60}")
    print(f"Worker finished — {datetime.now(timezone.utc).isoformat()}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()