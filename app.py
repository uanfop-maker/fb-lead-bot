import os, hmac, hashlib, json, threading, time
import requests
from flask import Flask, request, abort

app = Flask(__name__)

FB_APP_SECRET        = os.environ["FB_APP_SECRET"]
FB_VERIFY_TOKEN      = os.environ["FB_VERIFY_TOKEN"]
FB_PAGE_ACCESS_TOKEN = os.environ["FB_PAGE_ACCESS_TOKEN"]
TG_BOT_TOKEN         = os.environ["TG_BOT_TOKEN"]
TG_CHAT_ID           = os.environ["TG_CHAT_ID"]
SHEET_ID             = os.environ.get("SHEET_ID", "")
_SA_JSON             = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "{}")

GRAPH = "https://graph.facebook.com/v19.0"

# Google Sheets client (lazy init)
_sheets = None

def _get_sheets():
    global _sheets
    if _sheets is not None:
        return _sheets
    try:
        sa = json.loads(_SA_JSON)
        if not sa or not sa.get("private_key"):
            return None
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        creds = service_account.Credentials.from_service_account_info(
            sa, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        _sheets = build("sheets", "v4", credentials=creds).spreadsheets()
    except Exception as e:
        print(f"Sheets init error: {e}")
        _sheets = None
    return _sheets


def sheet_append(row: list):
    if not SHEET_ID:
        return
    svc = _get_sheets()
    if not svc:
        return
    try:
        svc.values().append(
            spreadsheetId=SHEET_ID,
            range="A1",
            valueInputOption="USER_ENTERED",
            body={"values": [row]},
        ).execute()
    except Exception as e:
        print(f"Sheets append error: {e}")


def tg_send(text):
    requests.post(
        f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage",
        json={"chat_id": TG_CHAT_ID, "text": text, "parse_mode": "HTML"},
        timeout=10,
    )


def fetch_lead(lead_id: str) -> dict:
    r = requests.get(
        f"{GRAPH}/{lead_id}",
        params={"access_token": FB_PAGE_ACCESS_TOKEN},
        timeout=10,
    )
    return r.json()


LABEL_MAP = {
    "full_name": "姓名", "email": "Email", "phone_number": "電話",
    "first_name": "名", "last_name": "姓",
}


def format_lead(data: dict, form_id: str, page_id: str) -> str:
    fields = {f["name"]: f["values"][0] for f in data.get("field_data", [])}
    lines = ["🔔 <b>新潛在顧客</b>"]
    for key, val in fields.items():
        lines.append(f"• {LABEL_MAP.get(key, key)}：{val}")
    lines.append(f"\n🕐 {data.get('created_time', '')[:16].replace('T', ' ')}")
    lines.append(f"📋 表單 ID：{form_id}")
    return "\n".join(lines)


def verify_signature(payload: bytes, sig_header: str) -> bool:
    if not sig_header or not sig_header.startswith("sha256="):
        return False
    expected = hmac.new(FB_APP_SECRET.encode(), payload, hashlib.sha256).hexdigest()
    return hmac.compare_digest(expected, sig_header[7:])


@app.get("/webhook")
def webhook_verify():
    mode      = request.args.get("hub.mode")
    token     = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if mode == "subscribe" and token == FB_VERIFY_TOKEN:
        print("Webhook verified")
        return challenge, 200
    abort(403)


@app.post("/webhook")
def webhook_event():
    sig = request.headers.get("X-Hub-Signature-256", "")
    if not verify_signature(request.data, sig):
        abort(403)

    body = request.get_json(force=True)
    for entry in body.get("entry", []):
        for change in entry.get("changes", []):
            if change.get("field") != "leadgen":
                continue
            val     = change["value"]
            lead_id = val.get("leadgen_id")
            form_id = val.get("form_id", "")
            page_id = val.get("page_id", "")

            def notify(lid=lead_id, fid=form_id, pid=page_id):
                time.sleep(2)
                data   = fetch_lead(lid)
                fields = {f["name"]: f["values"][0] for f in data.get("field_data", [])}
                msg    = format_lead(data, fid, pid)
                tg_send(msg)
                # Log to Google Sheet: time, name, email, phone, form_id, lead_id
                ts = data.get("created_time", "")[:16].replace("T", " ")
                row = [
                    ts,
                    fields.get("full_name") or f"{fields.get('first_name','')} {fields.get('last_name','')}".strip(),
                    fields.get("email", ""),
                    fields.get("phone_number", ""),
                    fid,
                    lid,
                ]
                sheet_append(row)

            threading.Thread(target=notify, daemon=True).start()

    return "ok", 200


@app.get("/health")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
