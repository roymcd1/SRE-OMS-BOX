import os
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from boxsdk import JWTAuth, Client
import openpyxl
import dateparser

print("🚀 Starting app.py")

# -------------------------------------------------
# 1. Environment variable for Box file ID
# -------------------------------------------------
load_dotenv()
BOX_FILE_ID = os.getenv("BOX_FILE_ID")
assert BOX_FILE_ID, "BOX_FILE_ID must be set in the app's environment variables"
print(f"📦 BOX_FILE_ID loaded: {BOX_FILE_ID}")

# -------------------------------------------------
# 2. Locate Box config JSON (secret mount or local)
# -------------------------------------------------
SECRET_PATH = "/secrets/box_config/box_config.json"
LOCAL_PATH  = "box_config.json"

CONFIG_PATH = SECRET_PATH if os.path.exists(SECRET_PATH) else LOCAL_PATH
print(f"📁 Using Box config file at: {CONFIG_PATH}")

# -------------------------------------------------
# 3. Authenticate with Box using the full settings file
# -------------------------------------------------
print("🔐 Authenticating with Box …")
auth   = JWTAuth.from_settings_file(CONFIG_PATH)
client = Client(auth)
print("✅ Box client initialized")

# -------------------------------------------------
# 4. Flask app & helpers
# -------------------------------------------------
app = Flask(__name__)


def download_excel_file():
    """Download the Excel schedule from Box to oncall_schedule.xlsx."""
    try:
        print(f"📥 Downloading Box file {BOX_FILE_ID} …")
        file_content = client.file(BOX_FILE_ID).content()
        with open("oncall_schedule.xlsx", "wb") as f:
            f.write(file_content)
        print("✅ Excel file saved locally")
        return True
    except Exception as e:
        print("❌ Error downloading file:", e)
        return str(e)


def parse_oncall_schedule():
    """Return a list of dicts from the Excel sheet."""
    try:
        wb      = openpyxl.load_workbook("oncall_schedule.xlsx")
        sheet   = wb.active
        headers = [cell.value for cell in sheet[1]]
        return [dict(zip(headers, row))
                for row in sheet.iter_rows(min_row=2, values_only=True)]
    except Exception as e:
        print("❌ Error parsing Excel:", e)
        return []


@app.route('/', methods=['GET'])
def home():
    return 'Box On-Call App is running ✅', 200


@app.route('/check-document', methods=['POST'])
def check_document():
    try:
        week_query = (request.get_json() or {}).get("week_query")
        if not week_query:
            return jsonify({"error": "Missing 'week_query' field"}), 400

        if download_excel_file() is not True:
            return jsonify({"error": "Excel download failed"}), 500

        data        = parse_oncall_schedule()
        target_date = dateparser.parse(week_query)
        if not (data and target_date):
            return jsonify({"error": "Could not parse data or date"}), 500

        for entry in data:
            start = dateparser.parse(str(entry.get("Start"))).date()
            end   = dateparser.parse(str(entry.get("End"))).date()
            if start <= target_date.date() <= end:
                return jsonify({
                    "start": str(start),
                    "end":   str(end),
                    "names": {
                        "primary":   entry.get("Primary"),
                        "secondary": entry.get("Secondary")
                    }
                })
        return jsonify({"message": "No match found"}), 404

    except Exception as e:
        print("❌ /check-document error:", e)
        return jsonify({"error": str(e)}), 500


@app.route('/when-am-i-on-call', methods=['POST'])
def when_am_i_on_call():
    try:
        name = (request.get_json() or {}).get("name")
        if not name:
            return jsonify({"error": "Missing 'name' field"}), 400

        if download_excel_file() is not True:
            return jsonify({"error": "Excel download failed"}), 500

        data     = parse_oncall_schedule()
        today    = dateparser.parse("today").date()
        upcoming = []

        for entry in data:
            start = dateparser.parse(str(entry.get("Start"))).date()
            end   = dateparser.parse(str(entry.get("End"))).date()
            if end >= today and (entry.get("Primary") == name or entry.get("Secondary") == name):
                upcoming.append({
                    "start":     str(start),
                    "end":       str(end),
                    "primary":   entry.get("Primary"),
                    "secondary": entry.get("Secondary")
                })

        return jsonify({"name": name, "upcoming_oncall": upcoming})

    except Exception as e:
        print("❌ /when-am-i-on-call error:", e)
        return jsonify({"error": str(e)}), 500


@app.route('/slack/events', methods=['POST'])
def slack_events():
    data = request.get_json()
    if data.get('type') == 'url_verification':
        return data.get('challenge'), 200
    print("📥 Received Slack event:", data)
    return '', 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)

