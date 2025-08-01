import os
from datetime import date, datetime
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from boxsdk import JWTAuth, Client
import openpyxl
import dateparser

print("🚀 Starting app.py")

# ----------------- 1. ENV VAR -----------------
load_dotenv()
BOX_FILE_ID = os.getenv("BOX_FILE_ID")
assert BOX_FILE_ID, "BOX_FILE_ID must be set"
print(f"📦 BOX_FILE_ID={BOX_FILE_ID}")

# ----------------- 2. CONFIG PATH -------------
SECRET_PATH = "/secrets/box_config/box_config.json"
CONFIG_PATH = SECRET_PATH if os.path.exists(SECRET_PATH) else "box_config.json"
print(f"📁 Using Box config: {CONFIG_PATH}")

# ----------------- 3. BOX AUTH ----------------
auth   = JWTAuth.from_settings_file(CONFIG_PATH)
client = Client(auth)
print("✅ Box client ready")

# ----------------- 4. HELPERS -----------------
def to_date(dt):
    """Return a `date` or None."""
    if dt is None:
        return None
    return dt.date() if hasattr(dt, "date") else dt

def download_excel():
    try:
        print("📥 Downloading Excel …")
        content = client.file(BOX_FILE_ID).content()
        with open("oncall.xlsx", "wb") as f:
            f.write(content)
        return True
    except Exception as e:
        print("❌ download_excel:", e)
        return str(e)

def parse_schedule():
    try:
        wb      = openpyxl.load_workbook("oncall.xlsx")
        sheet   = wb.active
        headers = [c.value for c in sheet[1]]
        return [dict(zip(headers, r))
                for r in sheet.iter_rows(min_row=2, values_only=True)]
    except Exception as e:
        print("❌ parse_schedule:", e)
        return []

# ----------------- 5. FLASK -------------------
app = Flask(__name__)

@app.route('/')
def home():
    return "Box On-Call App ✅", 200

@app.route('/check-document', methods=['POST'])
def check_document():
    try:
        week_query = (request.get_json() or {}).get("week_query")
        if not week_query:
            return jsonify({"error": "Missing 'week_query'"}), 400

        if download_excel() is not True:
            return jsonify({"error": "Excel download failed"}), 500

        target_dt = to_date(dateparser.parse(week_query))
        if not target_dt:
            return jsonify({"error": "Bad date"}), 400

        for row in parse_schedule():
            start = to_date(dateparser.parse(str(row.get("Start"))))
            end   = to_date(dateparser.parse(str(row.get("End"))))
            if start and end and start <= target_dt <= end:
                return jsonify({
                    "start": str(start),
                    "end":   str(end),
                    "names": {
                        "primary":   row.get("Primary"),
                        "secondary": row.get("Secondary")
                    }
                })
        return jsonify({"message": "No match"}), 404
    except Exception as e:
        print("❌ /check-document:", e)
        return jsonify({"error": str(e)}), 500

@app.route('/when-am-i-on-call', methods=['POST'])
def when_am_i_on_call():
    try:
        name = (request.get_json() or {}).get("name")
        if not name:
            return jsonify({"error": "Missing 'name'"}), 400

        if download_excel() is not True:
            return jsonify({"error": "Excel download failed"}), 500

        today = date.today()
        upcoming = []
        for row in parse_schedule():
            start = to_date(dateparser.parse(str(row.get("Start"))))
            end   = to_date(dateparser.parse(str(row.get("End"))))
            if (start and end and end >= today and
                (row.get("Primary") == name or row.get("Secondary") == name)):
                upcoming.append({
                    "start":     str(start),
                    "end":       str(end),
                    "primary":   row.get("Primary"),
                    "secondary": row.get("Secondary")
                })
        return jsonify({"name": name, "upcoming_oncall": upcoming})
    except Exception as e:
        print("❌ /when-am-i-on-call:", e)
        return jsonify({"error": str(e)}), 500

@app.route('/slack/events', methods=['POST'])
def slack_events():
    data = request.get_json()
    if data.get('type') == 'url_verification':
        return data.get('challenge'), 200
    print("📥 Slack event:", data)
    return '', 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)

