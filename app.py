import os, hmac, hashlib, json, threading, time, io
import requests
from datetime import datetime, timezone, timedelta
from flask import Flask, request, abort

app = Flask(__name__)

FB_APP_SECRET        = os.environ["FB_APP_SECRET"]
FB_VERIFY_TOKEN      = os.environ["FB_VERIFY_TOKEN"]
FB_PAGE_ACCESS_TOKEN = os.environ["FB_PAGE_ACCESS_TOKEN"]
TG_BOT_TOKEN         = os.environ["TG_BOT_TOKEN"]
TG_CHAT_ID           = os.environ["TG_CHAT_ID"]
SHEET_ID             = os.environ.get("SHEET_ID", "")
_SA_JSON             = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "{}")
FB_AD_ACCOUNT_ID     = os.environ.get("FB_AD_ACCOUNT_ID", "")
FB_MARKETING_TOKEN   = os.environ.get("FB_MARKETING_TOKEN", "")

PAGE_ID   = "101292869018588"
GRAPH     = "https://graph.facebook.com/v19.0"
TZ_TAIPEI = timezone(timedelta(hours=8))

LABEL_MAP = {
    "full_name": "姓名", "email": "Email", "phone_number": "電話",
    "first_name": "名", "last_name": "姓",
}
PLATFORM_ICON = {"FB": "📘", "IG": "📷"}

# ── Google Sheets ─────────────────────────────────────────────────────────────

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
            sa, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        _sheets = build("sheets", "v4", credentials=creds).spreadsheets()
    except Exception as e:
        print(f"Sheets init error: {e}")
    return _sheets


def sheet_append(row: list):
    if not SHEET_ID:
        return
    svc = _get_sheets()
    if not svc:
        return
    try:
        svc.values().append(
            spreadsheetId=SHEET_ID, range="A1",
            valueInputOption="USER_ENTERED",
            body={"values": [row]},
        ).execute()
    except Exception as e:
        print(f"Sheets append error: {e}")


# ── Telegram helpers ──────────────────────────────────────────────────────────

def tg_send(chat_id: str, text: str):
    requests.post(
        f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage",
        json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
        timeout=10,
    )


def tg_send_doc(chat_id: str, file_bytes: bytes, filename: str, caption: str = ""):
    requests.post(
        f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendDocument",
        data={"chat_id": chat_id, "caption": caption},
        files={"document": (filename, io.BytesIO(file_bytes),
               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        timeout=60,
    )


# ── Business day helpers ──────────────────────────────────────────────────────

def business_day_since_ts() -> int:
    now = datetime.now(TZ_TAIPEI)
    anchor = now.replace(hour=22, minute=0, second=0, microsecond=0)
    if now.hour < 22:
        anchor -= timedelta(days=1)
    return int(anchor.timestamp())


# ── Facebook API helpers ──────────────────────────────────────────────────────

def fetch_lead(lead_id: str) -> dict:
    r = requests.get(
        f"{GRAPH}/{lead_id}",
        params={"access_token": FB_PAGE_ACCESS_TOKEN,
                "fields": "id,created_time,field_data,platform,ad_name,form_id"},
        timeout=10)
    return r.json()


def fetch_all_leads(platform: str = None, since_ts: int = None, until_ts: int = None) -> list:
    forms_r = requests.get(
        f"{GRAPH}/{PAGE_ID}/leadgen_forms",
        params={"access_token": FB_PAGE_ACCESS_TOKEN, "fields": "id,name", "limit": 100},
        timeout=15)
    forms = forms_r.json().get("data", [])

    all_leads = []
    for form in forms:
        url = f"{GRAPH}/{form['id']}/leads"
        params = {
            "access_token": FB_PAGE_ACCESS_TOKEN,
            "fields": "id,created_time,field_data,platform,ad_name",
            "limit": 100,
        }
        if since_ts:
            params["since"] = since_ts
        if until_ts:
            params["until"] = until_ts
        while url:
            r = requests.get(url, params=params, timeout=15)
            data = r.json()
            for lead in data.get("data", []):
                lead["form_id"] = form["id"]
                lead["form_name"] = form.get("name", "")
                all_leads.append(lead)
            url = data.get("paging", {}).get("next")
            params = {}
    if platform:
        all_leads = [l for l in all_leads if l.get("platform", "").upper() == platform.upper()]
    return all_leads


def fetch_insights_stats(since_ts: int, until_ts: int) -> dict:
    if not FB_AD_ACCOUNT_ID or not FB_MARKETING_TOKEN:
        return None
    since_date = datetime.fromtimestamp(since_ts, TZ_TAIPEI).strftime("%Y-%m-%d")
    until_date = datetime.fromtimestamp(until_ts, TZ_TAIPEI).strftime("%Y-%m-%d")
    r = requests.get(
        f"{GRAPH}/{FB_AD_ACCOUNT_ID}/insights",
        params={
            "access_token": FB_MARKETING_TOKEN,
            "fields": "clicks,actions",
            "time_range": json.dumps({"since": since_date, "until": until_date}),
            "level": "account",
        }, timeout=15)
    data = r.json()
    if "error" in data:
        return {"error": data["error"].get("message", "API 錯誤")}
    result = {"clicks": 0, "lead_clicks": 0}
    for row in data.get("data", []):
        result["clicks"] += int(row.get("clicks", 0))
        for action in row.get("actions", []):
            if action["action_type"] == "onsite_conversion.lead_grouped":
                result["lead_clicks"] += int(action.get("value", 0))
    return result


# ── XLS generation ────────────────────────────────────────────────────────────

def leads_to_xlsx(leads: list, sheet_title: str = "潛在顧客") -> bytes:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(["時間(UTC+8)", "姓名", "Email", "電話", "平台", "廣告名稱", "表單", "Lead ID"])
    for lead in sorted(leads, key=lambda x: x.get("created_time", "")):
        ts = lead.get("created_time", "")
        ts_str = ""
        if ts:
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(TZ_TAIPEI)
            ts_str = dt.strftime("%Y-%m-%d %H:%M")
        fd = {f["name"]: f.get("values", [""])[0] for f in lead.get("field_data", [])}
        name = fd.get("full_name") or f"{fd.get('first_name', '')} {fd.get('last_name', '')}".strip()
        ws.append([ts_str, name, fd.get("email", ""), fd.get("phone_number", ""),
                   lead.get("platform", ""), lead.get("ad_name", ""),
                   lead.get("form_name", ""), lead.get("id", "")])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Duplicate detection ───────────────────────────────────────────────────────

def check_duplicate(new_lead_id: str, fields: dict) -> str:
    email = fields.get("email", "")
    phone = fields.get("phone_number", "")
    if not email and not phone:
        return ""
    try:
        all_leads = fetch_all_leads()
        for lead in all_leads:
            if lead.get("id") == new_lead_id:
                continue
            fd = {f["name"]: f.get("values", [""])[0] for f in lead.get("field_data", [])}
            if (email and fd.get("email") == email) or (phone and fd.get("phone_number") == phone):
                ts = lead.get("created_time", "")
                ts_str = ""
                if ts:
                    dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(TZ_TAIPEI)
                    ts_str = dt.strftime("%Y/%m/%d %H:%M")
                return f"⚠️ 重複：上次填於 {ts_str}"
    except Exception as e:
        print(f"Duplicate check error: {e}")
    return ""


# ── TG command handlers ───────────────────────────────────────────────────────

def handle_stats(chat_id: str):
    since_ts = business_day_since_ts()
    now_ts   = int(time.time())
    since_str = datetime.fromtimestamp(since_ts, TZ_TAIPEI).strftime("%m/%d %H:%M")
    now_str   = datetime.now(TZ_TAIPEI).strftime("%m/%d %H:%M")

    leads    = fetch_all_leads(since_ts=since_ts, until_ts=now_ts)
    fb_leads = [l for l in leads if l.get("platform", "").upper() == "FB"]
    ig_leads = [l for l in leads if l.get("platform", "").upper() == "IG"]
    stats    = fetch_insights_stats(since_ts, now_ts)

    lines = [f"📊 <b>統計報表</b>  {since_str} → {now_str}", ""]

    if stats and "error" not in stats:
        lines.append(f"👆 點擊 (Click)：<b>{stats['clicks']}</b>")
        lc = stats['lead_clicks']
        pct1 = f"  ({round(lc / stats['clicks'] * 100)}%)" if stats['clicks'] else ""
        lines.append(f"📝 Lead 點擊（開表單）：<b>{lc}</b>{pct1}")
        pct2 = f"  ({round(len(leads) / lc * 100)}%)" if lc else ""
        lines.append(f"✅ Lead（填完送出）：<b>{len(leads)}</b>{pct2}")
    else:
        if stats and "error" in stats:
            lines.append("⚠️ 廣告數據暫無法取得")
        else:
            lines.append("⚠️ 點擊數據未設定（需 FB_AD_ACCOUNT_ID）")
        lines.append(f"✅ Lead（填完送出）：<b>{len(leads)}</b>")

    lines += ["", f"📘 FB：{len(fb_leads)} 筆", f"📷 IG：{len(ig_leads)} 筆"]
    tg_send(chat_id, "\n".join(lines))


def handle_download(chat_id: str, platform: str = None):
    label = {"FB": "FB", "IG": "IG"}.get(platform, "全部")
    tg_send(chat_id, f"⏳ 準備 {label} 表單資料...")
    leads = fetch_all_leads(platform=platform)
    if not leads:
        tg_send(chat_id, f"⚠️ 找不到 {label} Lead 資料")
        return
    now_str  = datetime.now(TZ_TAIPEI).strftime("%Y%m%d_%H%M")
    filename = f"慢序選物所_{label}表單_{now_str}.xlsx"
    xlsx     = leads_to_xlsx(leads, f"{label}潛在顧客")
    tg_send_doc(chat_id, xlsx, filename, f"📥 {label} 潛在顧客（共 {len(leads)} 筆）")


def handle_find(chat_id: str, query: str):
    if not query:
        tg_send(chat_id, "用法：/find 電話 或 /find Email")
        return
    tg_send(chat_id, f"🔍 搜尋 {query}...")
    leads = fetch_all_leads()
    results = []
    q = query.lower()
    for lead in leads:
        fd = {f["name"]: f.get("values", [""])[0] for f in lead.get("field_data", [])}
        if q in fd.get("email", "").lower() or q in fd.get("phone_number", "").lower():
            results.append((lead, fd))
    if not results:
        tg_send(chat_id, "找不到符合的 Lead")
        return
    lines = [f"🔍 找到 {len(results)} 筆"]
    for lead, fd in results[:5]:
        ts = lead.get("created_time", "")
        ts_str = ""
        if ts:
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(TZ_TAIPEI)
            ts_str = dt.strftime("%Y/%m/%d %H:%M")
        name = fd.get("full_name") or f"{fd.get('first_name', '')} {fd.get('last_name', '')}".strip()
        icon = PLATFORM_ICON.get(lead.get("platform", ""), "")
        lines.append(f"{icon} {ts_str}  {name}  {fd.get('email', '')}  {fd.get('phone_number', '')}")
    tg_send(chat_id, "\n".join(lines))


def handle_adrank(chat_id: str):
    tg_send(chat_id, "⏳ 統計廣告效益...")
    leads = fetch_all_leads()
    counts: dict = {}
    for lead in leads:
        key = lead.get("ad_name") or "（無廣告名稱）"
        counts[key] = counts.get(key, 0) + 1
    if not counts:
        tg_send(chat_id, "暫無廣告資料")
        return
    lines = ["🏆 廣告 Lead 排行"]
    for i, (name, cnt) in enumerate(sorted(counts.items(), key=lambda x: -x[1])[:10], 1):
        lines.append(f"{i}. {name}：{cnt} 筆")
    tg_send(chat_id, "\n".join(lines))


# ── Auto daily report ─────────────────────────────────────────────────────────

def daily_report():
    since_ts  = business_day_since_ts()
    now_ts    = int(time.time())
    leads     = fetch_all_leads(since_ts=since_ts, until_ts=now_ts)
    fb_leads  = [l for l in leads if l.get("platform", "").upper() == "FB"]
    ig_leads  = [l for l in leads if l.get("platform", "").upper() == "IG"]
    since_str = datetime.fromtimestamp(since_ts, TZ_TAIPEI).strftime("%m/%d %H:%M")
    now_str   = datetime.now(TZ_TAIPEI).strftime("%m/%d %H:%M")
    stats     = fetch_insights_stats(since_ts, now_ts)

    lines = [f"🌙 <b>每日報表</b>  {since_str} → {now_str}", ""]
    if stats and "error" not in stats:
        lines.append(f"👆 點擊：<b>{stats['clicks']}</b>")
        lines.append(f"📝 開表單：<b>{stats['lead_clicks']}</b>")
    lines.append(f"✅ Lead：<b>{len(leads)}</b>（📘FB {len(fb_leads)} / 📷IG {len(ig_leads)}）")

    tg_send(TG_CHAT_ID, "\n".join(lines))

    if leads:
        xlsx    = leads_to_xlsx(leads, "當日潛在顧客")
        now_fn  = datetime.now(TZ_TAIPEI).strftime("%Y%m%d")
        tg_send_doc(TG_CHAT_ID, xlsx, f"慢序選物所_{now_fn}日報.xlsx",
                    caption=f"📎 {now_fn} 日報 XLS")


def _start_scheduler():
    from apscheduler.schedulers.background import BackgroundScheduler
    from apscheduler.triggers.cron import CronTrigger
    scheduler = BackgroundScheduler(timezone="Asia/Taipei")
    scheduler.add_job(daily_report, CronTrigger(hour=22, minute=0))
    scheduler.start()


# ── Webhook verification helpers ──────────────────────────────────────────────

def format_lead_msg(data: dict, form_id: str) -> str:
    fields = {f["name"]: f.get("values", [""])[0] for f in data.get("field_data", [])}
    platform = data.get("platform", "")
    icon = PLATFORM_ICON.get(platform, "")
    lines = [f"🔔 <b>新潛在顧客</b> {icon}"]
    for key, val in fields.items():
        lines.append(f"• {LABEL_MAP.get(key, key)}：{val}")
    ts = data.get("created_time", "")
    if ts:
        dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(TZ_TAIPEI)
        lines.append(f"\n🕐 {dt.strftime('%Y-%m-%d %H:%M')} (UTC+8)")
    lines.append(f"📋 表單 ID：{form_id}")
    return "\n".join(lines)


def verify_signature(payload: bytes, sig_header: str) -> bool:
    if not sig_header or not sig_header.startswith("sha256="):
        return False
    expected = hmac.new(FB_APP_SECRET.encode(), payload, hashlib.sha256).hexdigest()
    return hmac.compare_digest(expected, sig_header[7:])


# ── Flask routes ──────────────────────────────────────────────────────────────

@app.get("/webhook")
def webhook_verify():
    mode      = request.args.get("hub.mode")
    token     = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if mode == "subscribe" and token == FB_VERIFY_TOKEN:
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

            def notify(lid=lead_id, fid=form_id):
                time.sleep(2)
                data   = fetch_lead(lid)
                fields = {f["name"]: f.get("values", [""])[0] for f in data.get("field_data", [])}
                dup    = check_duplicate(lid, fields)
                msg    = format_lead_msg(data, fid)
                if dup:
                    msg += f"\n{dup}"
                tg_send(TG_CHAT_ID, msg)
                ts = data.get("created_time", "")
                ts_str = ""
                if ts:
                    dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(TZ_TAIPEI)
                    ts_str = dt.strftime("%Y-%m-%d %H:%M")
                name = fields.get("full_name") or f"{fields.get('first_name', '')} {fields.get('last_name', '')}".strip()
                sheet_append([ts_str, name, fields.get("email", ""),
                              fields.get("phone_number", ""), data.get("platform", ""), fid, lid])

            threading.Thread(target=notify, daemon=True).start()
    return "ok", 200


@app.post("/tg-update")
def tg_update():
    body = request.get_json(force=True)
    msg  = body.get("message", {})
    if not msg:
        return "ok", 200
    chat_id = str(msg.get("chat", {}).get("id", ""))
    text    = (msg.get("text") or "").strip()
    if not text.startswith("/"):
        return "ok", 200
    parts = text.split()
    cmd   = parts[0].lower()
    if "@" in cmd:
        cmd = cmd[:cmd.index("@")]
    args = parts[1:]

    dispatch = {
        "/stats":  lambda: threading.Thread(target=handle_stats,    args=(chat_id,),              daemon=True).start(),
        "/fb":     lambda: threading.Thread(target=handle_download,  args=(chat_id, "FB"),         daemon=True).start(),
        "/ig":     lambda: threading.Thread(target=handle_download,  args=(chat_id, "IG"),         daemon=True).start(),
        "/all":    lambda: threading.Thread(target=handle_download,  args=(chat_id, None),         daemon=True).start(),
        "/find":   lambda: threading.Thread(target=handle_find,      args=(chat_id, " ".join(args)), daemon=True).start(),
        "/adrank": lambda: threading.Thread(target=handle_adrank,   args=(chat_id,),              daemon=True).start(),
    }
    if cmd in dispatch:
        dispatch[cmd]()
    return "ok", 200


@app.get("/health")
def health():
    return "ok"


@app.get("/bulk-import")
def bulk_import():
    svc = _get_sheets()
    if not svc:
        return "Google Sheets not configured", 503
    leads = fetch_all_leads()
    if not leads:
        return "no leads", 200
    rows = [["時間(UTC+8)", "姓名", "Email", "電話", "平台", "廣告名稱", "表單", "Lead ID"]]
    for lead in sorted(leads, key=lambda x: x.get("created_time", "")):
        ts = lead.get("created_time", "")
        ts_str = ""
        if ts:
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(TZ_TAIPEI)
            ts_str = dt.strftime("%Y-%m-%d %H:%M")
        fd = {f["name"]: f.get("values", [""])[0] for f in lead.get("field_data", [])}
        name = fd.get("full_name") or f"{fd.get('first_name', '')} {fd.get('last_name', '')}".strip()
        rows.append([ts_str, name, fd.get("email", ""), fd.get("phone_number", ""),
                     lead.get("platform", ""), lead.get("ad_name", ""),
                     lead.get("form_name", ""), lead.get("id", "")])
    svc.values().update(
        spreadsheetId=SHEET_ID, range="A1",
        valueInputOption="USER_ENTERED", body={"values": rows}
    ).execute()
    return f"imported {len(leads)} leads", 200


# ── Startup (runs when gunicorn imports this module) ──────────────────────────

def _setup_tg_webhook():
    url = "https://fb-lead-bot.zeabur.app/tg-update"
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{TG_BOT_TOKEN}/setWebhook",
            json={"url": url, "allowed_updates": ["message"]},
            timeout=10,
        )
        print(f"TG webhook: {r.json()}")
    except Exception as e:
        print(f"TG webhook setup error: {e}")


_setup_tg_webhook()
_start_scheduler()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
