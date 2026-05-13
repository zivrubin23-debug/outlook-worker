#!/usr/bin/env python3
"""
My Secret Sex Toy Delivery — Automated Email Support Worker
Railway Cron Job — runs every 30 minutes

Optimizations:
  - Pre-filter skips obvious spam/marketing before Claude
  - Haiku for classification, Sonnet only for reply generation
  - Truncated email body (max 1500 chars) sent to Claude
  - Short system prompt
"""

import os
import json
import re
import requests
import anthropic
from datetime import datetime, timezone
from urllib.parse import urlparse, parse_qs

# ── Configuration ─────────────────────────────────────────────────────────────

CLIENT_ID             = os.environ["MS365_MCP_CLIENT_ID"]
CLIENT_SECRET         = os.environ["MS365_MCP_CLIENT_SECRET"]
TENANT_ID             = os.environ["MS365_MCP_TENANT_ID"]
ANTHROPIC_KEY         = os.environ["ANTHROPIC_API_KEY"]
MAILBOX_USER          = os.environ["MAILBOX_USER"]
SHOPIFY_CLIENT_ID     = os.environ["SHOPIFY_CLIENT_ID"]
SHOPIFY_CLIENT_SECRET = os.environ["SHOPIFY_CLIENT_SECRET"]

SHOPIFY_STORE   = "40026a.myshopify.com"
SHOPIFY_VERSION = "2024-01"
GRAPH_BASE      = "https://graph.microsoft.com/v1.0"

CUTOFF_DATE = datetime(2026, 5, 11, 0, 0, 0, tzinfo=timezone.utc)

SUPPLIER_FOLDER_ID = "AAMkADBhNDBkMzJjLWRiZTMtNDE2NC1iNDQ1LTQ4ZjVlMWE3ZGFkYwAuAAAAAAC5ypHMZNNXRo-eNkzzCxTpAQCygDv4EiYcSqbyqT-RbGLGAALMrinTAAA="

TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# ── Pre-filter: skip without Claude ──────────────────────────────────────────

SKIP_SENDERS = {
    "funnymee.com", "163.com", "bezlya.com",
    "info@mysecretsextoydelivery.com",
    "klaviyo.com", "hello.klaviyo.com",
    "shop.tiktok.com", "taylor@shop.tiktok.com", "sam@shop.tiktok.com",
    "partners@buddify.app", "brand.faire.com", "entervending.com",
    "sextoydistributing.com", "lovetoyus.com", "info-dysotoys.com",
    "secomapp.com", "clients.myguestlist.com.au", "swaysucker.com",
    "melodeem.com", "blq.com", "vscnovelty.com", "cucupie.com",
    "purepleasure-4.com", "fairyl.com", "foxmail.com", "adamssceptre.com",
    "getsweetums.com", "miamidistro1.com", "tabs.co", "swisstransfer.com",
    "shop.tiktok.com", "lovetoyus.com", "entervending.com",
    "mail.zapier.com", "buddify.app", "creativeoutdoor.com",
}

SKIP_SUBJECT_KEYWORDS = {
    "newsletter", "unsubscribe", "webinar", "live training",
    "affiliate", "% off", "promo", "coupon", "overstock",
    "new arrivals", "hot sellers", "tiktok shop", "free ads",
    "your first 60 days", "automatic reply", "out of office",
    "undeliverable", "delivery failure", "data deletion",
    "scale global", "drive revenue", "klaviyo", "buddify",
    "shopify markets", "integration spotlight", "new products",
    "brand new", "shop now", "view as webpage",
}


def should_skip_without_claude(subject: str, body_preview: str) -> bool:
    """Return True if email is obviously not a real customer — skip Claude."""
    s = (subject or "").lower()
    b = (body_preview or "")[:300].lower()
    return any(kw in s or kw in b for kw in SKIP_SUBJECT_KEYWORDS)

# ── System Prompts (short) ────────────────────────────────────────────────────

CLASSIFY_PROMPT = """You are a classifier for a sex toy delivery store's support inbox.

Decide if an email is from a REAL CUSTOMER about an order (cancellation, tracking, shipping, return, refund, product question, complaint) OR not (spam, supplier, marketing, B2B).

Reply ONLY with JSON:
{"is_customer_email": true/false, "email_type": "cancellation|tracking|return|refund|complaint|product_question|other"}"""

REPLY_PROMPT = """You are a customer support agent for My Secret Sex Toy Delivery (mysecretsextoydelivery.com).

Reply warmly, professionally, discreetly. Same language as customer. Short paragraphs.

POLICIES:
- CANCELLATION: "We cannot guarantee cancellation as the order may already be with our shipping provider. We will check and update you."
- TRACKING: Ask for order number if missing. If no tracking yet: "Your order is with our shipping provider. We'll send tracking as soon as we have it."
- RETURNS: NEVER give return address. Say: "Reply and we'll arrange an RMA number and return instructions." → DRAFT only.
- REFUNDS: DRAFT only. Never confirm amounts.
- NO PHYSICAL STORE: Online only.

DRAFT (should_auto_send=false) for: refunds, returns, disputes, chargebacks, angry/legal/fraud, package not received, double charges.
AUTO-SEND for: tracking questions, cancellations, order status, product questions, store location.

SIGNATURE (always at end):
___________________

Best regards,

John
Service Team Leader | MS

Online Store: www.mysecretsextoydelivery.com

Reply ONLY with JSON:
{"should_auto_send": true/false, "sensitivity_reason": "", "reply": "full reply text with signature"}"""

# ── HTML Formatting ───────────────────────────────────────────────────────────

def format_reply_html(reply_text: str) -> str:
    MARKER = "___________________"
    if MARKER in reply_text:
        body_part, sig_part = reply_text.split(MARKER, 1)
    else:
        body_part, sig_part = reply_text, ""

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
            f'<p style="color:#555;font-size:13px;line-height:1.6;">{sig_inner}</p>\n'
        )

    return (
        '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">\n'
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

# ── Microsoft Graph ───────────────────────────────────────────────────────────

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
    return requests.request(method, f"{GRAPH_BASE}{path}", headers=headers, timeout=20, **kwargs)


def get_unread_inbox(token: str) -> list:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/inbox/messages", token, params={
        "$filter": "isRead eq false",
        "$select": "id,subject,from,body,bodyPreview,receivedDateTime,conversationId,categories",
        "$orderby": "receivedDateTime asc",
        "$top": 50,
    })
    resp.raise_for_status()
    return resp.json().get("value", [])


def get_unread_supplier_emails(token: str) -> list:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/{SUPPLIER_FOLDER_ID}/messages", token, params={
        "$filter": "isRead eq false",
        "$select": "id,subject,from,body,bodyPreview,receivedDateTime",
        "$orderby": "receivedDateTime asc",
        "$top": 25,
    })
    resp.raise_for_status()
    return resp.json().get("value", [])


def already_replied(token: str, conversation_id: str, received_dt: str) -> bool:
    resp = graph("GET", f"/users/{MAILBOX_USER}/mailFolders/sentitems/messages", token, params={
        "$filter": f"conversationId eq '{conversation_id}'",
        "$select": "sentDateTime",
        "$top": 10,
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
          headers={"Content-Type": "application/json"}, json={"isRead": True})


def mark_failed(token: str, message_id: str):
    """Tag email so we don't retry forever."""
    graph("PATCH", f"/users/{MAILBOX_USER}/messages/{message_id}", token,
          headers={"Content-Type": "application/json"},
          json={"categories": ["worker-failed"]})


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

# ── Shopify ───────────────────────────────────────────────────────────────────

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


def shopify_get(path: str, shopify_token: str) -> dict:
    resp = requests.get(
        f"https://{SHOPIFY_STORE}/admin/api/{SHOPIFY_VERSION}/{path}",
        headers={"X-Shopify-Access-Token": shopify_token, "Content-Type": "application/json"},
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()


def shopify_post(path: str, data: dict, shopify_token: str) -> dict:
    resp = requests.post(
        f"https://{SHOPIFY_STORE}/admin/api/{SHOPIFY_VERSION}/{path}",
        headers={"X-Shopify-Access-Token": shopify_token, "Content-Type": "application/json"},
        json=data, timeout=15,
    )
    resp.raise_for_status()
    return resp.json()

# ── Supplier email helpers ────────────────────────────────────────────────────

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


def extract_supplier_email_data(body_text: str, body_raw: str = "") -> dict:
    name_match = re.search(r"Hi\s+(.+?),\s+your order has shipped", body_text, re.IGNORECASE)
    customer_name = name_match.group(1).strip() if name_match else ""

    carrier_domains = [
        "tools.usps.com", "usps.com/go",
        "fedex.com/fedextrack", "fedex.com/tracking",
        "ups.com/track", "ontrac.com/tracking",
    ]

    tracking_url = ""
    for source in [body_raw, body_text]:
        if not source:
            continue
        for url in re.findall(r'https?://[^\s\)\>\"\'<\|]+', source):
            url = url.strip().rstrip(".,;>")
            if any(d in url.lower() for d in carrier_domains):
                tracking_url = url
                break
        if tracking_url:
            break

    return {"customer_name": customer_name, "tracking_url": tracking_url}


def find_unfulfilled_order(customer_name: str, shopify_token: str) -> dict:
    try:
        name_parts = customer_name.strip().split()
        if len(name_parts) == 2 and name_parts[0].lower() == name_parts[1].lower():
            search_name = name_parts[0].lower()
            print(f"   Duplicate name — searching as: '{search_name}'")
        else:
            search_name = customer_name.strip().lower()

        orders = shopify_get(
            "orders.json?fulfillment_status=unfulfilled&status=open"
            "&limit=100&fields=id,name,customer",
            shopify_token,
        ).get("orders", [])

        for order in orders:
            customer = order.get("customer") or {}
            first = (customer.get("first_name") or "").strip().lower()
            last  = (customer.get("last_name")  or "").strip().lower()
            full  = f"{first} {last}".strip()
            if first == search_name or last == search_name or full == search_name:
                return order

        return {}

    except Exception as exc:
        print(f"   ERROR searching orders: {exc}")
        return {}


def create_fulfillment(order_id: int, tracking_number: str,
                       tracking_url: str, carrier: str,
                       shopify_token: str) -> bool:
    try:
        fulfillment_orders = shopify_get(
            f"orders/{order_id}/fulfillment_orders.json", shopify_token,
        ).get("fulfillment_orders", [])

        open_fos = [fo for fo in fulfillment_orders if fo["status"] == "open"]
        if not open_fos:
            print(f"   No open fulfillment orders for {order_id}")
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

        if "fulfillment" in result:
            f = result["fulfillment"]
            print(f"   Fulfillment ID: {f.get('id')} | status: {f.get('status')}")
            return True

        return False

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


def truncate_body(text: str, max_chars: int = 1500) -> str:
    """Limit email body to reduce Claude token usage."""
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n[... truncated]"

# ── Claude (two-stage) ────────────────────────────────────────────────────────

def classify_email(subject: str, body: str, sender: str) -> dict:
    """Stage 1: Haiku — cheap classification only."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    message = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=100,
        system=CLASSIFY_PROMPT,
        messages=[{
            "role": "user",
            "content": f"From: {sender}\nSubject: {subject}\n\nBody (first 500 chars):\n{body[:500]}"
        }],
    )
    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.lower().startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


def generate_reply(subject: str, body: str, sender: str) -> dict:
    """Stage 2: Sonnet — full reply generation (only for real customers)."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=1000,
        system=REPLY_PROMPT,
        messages=[{
            "role": "user",
            "content": f"From: {sender}\nSubject: {subject}\n\nBody:\n{body}"
        }],
    )
    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.lower().startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())

# ── Workflow 1: Supplier tracking ─────────────────────────────────────────────

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
            print("   Skipped — older than cutoff")
            continue

        data          = extract_supplier_email_data(body_text, body_raw)
        customer_name = data["customer_name"]
        tracking_url  = data["tracking_url"]

        if not customer_name:
            print("   Could not extract customer name")
            continue

        if not tracking_url:
            print(f"   No tracking URL found for '{customer_name}'")
            continue

        carrier         = detect_carrier(tracking_url)
        tracking_number = extract_tracking_number(tracking_url, carrier)

        print(f"   Customer:  {customer_name}")
        print(f"   Carrier:   {carrier}")
        print(f"   Tracking:  {tracking_number}")

        order = find_unfulfilled_order(customer_name, shopify_token)
        if not order:
            print(f"   WARNING — No unfulfilled order for '{customer_name}'")
            notify(f"⚠️ <b>No order match</b>\nCustomer: {customer_name}\n{carrier} | {tracking_number}")
            continue

        order_name = order.get("name", "")
        order_id   = order["id"]
        print(f"   Shopify order: {order_name} (ID: {order_id})")

        success = create_fulfillment(order_id, tracking_number, tracking_url, carrier, shopify_token)

        if success:
            mark_read(ms_token, msg_id)
            print(f"   ✅ Fulfilled — customer will receive tracking email")
            notify(f"📦 <b>Fulfilled</b>\nCustomer: {customer_name} | Order: {order_name}\n{carrier} | {tracking_number}")
        else:
            print(f"   ❌ Fulfillment failed — not marked read")
            notify(f"❌ <b>Fulfillment failed</b>\nCustomer: {customer_name} | Order: {order_name}")

# ── Workflow 2: Customer support ──────────────────────────────────────────────

def process_customer_emails(ms_token: str):
    print(f"\n{'─'*60}")
    print("WORKFLOW 2 — Customer support emails")
    print(f"{'─'*60}")

    emails = get_unread_inbox(ms_token)
    print(f"Unread inbox emails: {len(emails)}")

    skipped_prefilter  = 0
    skipped_noncustomer = 0
    processed          = 0

    for email in emails:
        sender_address = email["from"]["emailAddress"].get("address", "").lower()
        sender_domain  = sender_address.split("@")[-1] if "@" in sender_address else ""
        subject        = email.get("subject") or ""
        body_obj       = email.get("body", {})
        body_raw       = body_obj.get("content", email.get("bodyPreview", ""))
        body_text      = strip_html(body_raw) if body_obj.get("contentType") == "html" else body_raw
        body_preview   = email.get("bodyPreview", "")
        conv_id        = email["conversationId"]
        msg_id         = email["id"]
        received_dt    = email["receivedDateTime"]
        categories     = email.get("categories", [])

        print(f"\n-> '{subject}' | from: {sender_address}")

        if parse_dt(received_dt) < CUTOFF_DATE:
            print("   Skipped — older than cutoff")
            continue

        # Skip if previously failed
        if "worker-failed" in categories:
            print("   Skipped — previously failed, needs manual review")
            continue

        # Pre-filter: known non-customer senders
        if sender_domain in SKIP_SENDERS or sender_address in SKIP_SENDERS:
            print("   Skipped — known non-customer sender")
            mark_read(ms_token, msg_id)
            skipped_prefilter += 1
            continue

        # Pre-filter: obvious non-customer subjects/content
        if should_skip_without_claude(subject, body_preview):
            print("   Skipped — pre-filter (no Claude used)")
            mark_read(ms_token, msg_id)
            skipped_prefilter += 1
            continue

        if already_replied(ms_token, conv_id, received_dt):
            print("   Skipped — already replied")
            mark_read(ms_token, msg_id)
            continue

        if draft_exists(ms_token, conv_id):
            print("   Skipped — draft already exists")
            continue

        # Stage 1: Haiku classification (cheap)
        body_truncated = truncate_body(body_text)
        try:
            classification = classify_email(subject, body_truncated, sender_address)
        except Exception as exc:
            print(f"   ERROR — Haiku classification failed: {exc}")
            mark_failed(ms_token, msg_id)
            continue

        if not classification.get("is_customer_email", False):
            print(f"   Skipped — not a customer email (Haiku)")
            mark_read(ms_token, msg_id)
            skipped_noncustomer += 1
            continue

        print(f"   Customer email detected: {classification.get('email_type', 'other')}")

        # Stage 2: Sonnet reply generation (only for real customers)
        try:
            result = generate_reply(subject, body_truncated, sender_address)
        except Exception as exc:
            print(f"   ERROR — Sonnet reply failed: {exc}")
            mark_failed(ms_token, msg_id)
            continue

        reply_text = result.get("reply", "")
        processed += 1

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

    print(f"\n   Summary: {skipped_prefilter} pre-filtered | {skipped_noncustomer} non-customer (Haiku) | {processed} processed (Sonnet)")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    ts = datetime.now(timezone.utc).isoformat()
    print(f"\n{'='*60}")
    print(f"Email Worker — {ts}")
    print(f"Mailbox:  {MAILBOX_USER}")
    print(f"Shopify:  {SHOPIFY_STORE}")
    print(f"Cutoff:   {CUTOFF_DATE.isoformat()}")
    print(f"{'='*60}")

    ms_token = get_ms_token()

    print("\nAuthenticating with Shopify...")
    try:
        shopify_token = get_shopify_token()
        print("Shopify token: OK")
    except Exception as exc:
        print(f"ERROR — Shopify auth failed: {exc}")
        shopify_token = None

    if shopify_token:
        process_supplier_emails(ms_token, shopify_token)

    process_customer_emails(ms_token)

    print(f"\n{'='*60}")
    print(f"Worker finished — {datetime.now(timezone.utc).isoformat()}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()

גם עדכן את railway.toml:
[build]
builder = "NIXPACKS"

[deploy]
startCommand = "python worker.py"
cronSchedule = "*/30 * * * *"
restartPolicyType = "NEVER"