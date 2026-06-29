import os, hmac, hashlib, json, threading, time
import requests
from flask import Flask, request, abort

app = Flask(__name__)

FB_APP_SECRET       = os.environ["FB_APP_SECRET"]
FB_VERIFY_TOKEN     = os.environ["FB_VERIFY_TOKEN"]
FB_PAGE_ACCESS_TOKEN = os.environ["FB_PAGE_ACCESS_TOKEN"]
TG_BOT_TOKEN        = os.environ["TG_BOT_TOKEN"]
TG_CHAT_ID          = os.environ["TG_CHAT_ID"]

GRAPH = "https://graph.facebook.com/v19.0"


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


def format_lead(data: dict, form_id: str, page_id: str) -> str:
    fields = {f["name"]: f["values"][0] for f in data.get("field_data", [])}
    lines = ["🔔 <b>新潛在顧客</b>"]
    label_map = {
        "full_name": "姓名", "email": "Email", "phone_number": "電話",
        "first_name": "名", "last_name": "姓",
    }
    for key, val in fields.items():
        label = label_map.get(key, key)
        lines.append(f"• {label}：{val}")
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
                time.sleep(2)  # wait for FB to have the lead ready
                data = fetch_lead(lid)
                msg  = format_lead(data, fid, pid)
                tg_send(msg)

            threading.Thread(target=notify, daemon=True).start()

    return "ok", 200


@app.get("/health")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
