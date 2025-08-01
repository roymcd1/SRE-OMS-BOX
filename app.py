import os
import json
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from boxsdk import JWTAuth, Client
import openpyxl
import dateparser
from datetime import timedelta

print("🚀 Starting app.py")

# Load environment variables
load_dotenv()
BOX_FILE_ID = os.getenv("BOX_FILE_ID")
assert BOX_FILE_ID, "BOX_FILE_ID must be set in .env"
print(f"📦 BOX_FILE_ID loaded: {BOX_FILE_ID}")

# Load Box config from mounted secret or local file
if os.path.exists("/secrets/box_config/box_config.json"):
    CONFIG_PATH = "/secrets/box_config/box_config.json"
    print("📁 Loading box config from secret mount...")
else:
    CONFIG_PATH = "box_config.json"
    print("📁 Loading box config from local file...")

with open(CONFIG_PATH) as config_file:
    full_config = json.load(config_file)

box_config = {
    "clientID": full_config["boxAppSettings"]["clientID"],
    "clientSecret": full_config["boxAppSettings"]["clientSecret"],
    "appAuth": full_config["boxAppSettings"]["appAuth"],
    "enterpriseID": full_config["enterpriseID"]
}

print("🔐 Authenticating with Box...")
auth = JWTAuth.from_settings_dictionary(box_config)
auth.authenticate_instance()
client = Client(auth)
print("✅ Box client initialized")

app = Flask(__name__)


def download_excel_file():
    try:
        print(f"📥 Attempting to download file with BOX_FILE_ID={BOX_FILE_ID}")
        file_content = client.file(BOX_FILE_ID).content()
        with open("oncall_schedule.xlsx", "wb") as f:
            f.write(file_content)
        print("✅ Excel file downloaded and saved locally")
        return True
    except Exception as e:
        print("❌ Error downloading Excel file from Box:", e)
        return str(e)


def parse_oncall_schedule():
    try:
        wb = openpyxl.load_workbook("oncall_schedule.xlsx")
        sheet = wb.active
        headers = [cell.value for cell in sheet[1]]
        data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_dict = dict(zip(headers, row))
            data.append(row_dict)
        return data
    except Exception as e:
        print("❌ Error parsing Excel file:", e)
        return []


@app.route('/', methods=['GET'])
def home():
    return 'Box On-Call App is running ✅', 200


@app.route('/check-document', methods=['POST'])
def check_document():
    try:
        content = request.get_json()
        week_query = content.get("week_query")

        if not week_query:
            return jsonify({"error": "Missing 'week_query' field"}), 400

        download_result = download_excel_file()
        if download_result is not True:
            return jsonify({
                "error": "Excel file could not be downloaded",
                "detail": download_result
            }), 500

        data = parse_oncall_schedule()
        if not data:
            return jsonify({"error": "Could not parse Excel file"}), 500

        target_date = dateparser.parse(week_query)
        if not target_date:
            return jsonify({"error": "Could not parse date"}), 400

        for entry in data:
            try:
                start = dateparser.parse(str(entry.get("Start"))).date()
                end = dateparser.parse(str(entry.get("End"))).date()
                if start <= target_date.date() <= end:
                    return jsonify({
                        "start": str(start),
                        "end": str(end),
                        "names": {
                            "primary": entry.get("Primary"),
                            "secondary": entry.get("Secondary")
                        }
                    })
            except Exception as e:
                print("❌ Error parsing entry:", entry, e)
                continue

        return jsonify({"message": "No match found"}), 404

    except Exception as e:
        print("❌ Error in /check-document:", e)
        return jsonify({"error": str(e)}), 500


@app.route('/when-am-i-on-call', methods=['POST'])
def when_am_i_on_call():
    try:
        content = request.get_json()
        name = content.get("name")

        if not name:
            return jsonify({"error": "Missing 'name' field"}), 400

        download_result = download_excel_file()
        if download_result is not True:
            return jsonify({
                "error": "Excel file could not be downloaded",
                "detail": download_result
            }), 500

        data = parse_oncall_schedule()
        if not data:
            return jsonify({"error": "Could not parse Excel file"}), 500

        upcoming = []
        today = dateparser.parse("today").date()

        for entry in data:
            try:
                start = dateparser.parse(str(entry.get("Start"))).date()
                end = dateparser.parse(str(entry.get("End"))).date()
                if end >= today and (entry.get("Primary") == name or entry.get("Secondary") == name):
                    upcoming.append({
                        "start": str(start),
                        "end": str(end),
                        "primary": entry.get("Primary"),
                        "secondary": entry.get("Secondary")
                    })
            except Exception:
                continue

        return jsonify({
            "name": name,
            "upcoming_oncall": upcoming
        })

    except Exception as e:
        print("❌ Error in /when-am-i-on-call:", e)
        return jsonify({"error": str(e)}), 500


@app.route('/slack/events', methods=['POST'])
def slack_events():
    data = request.get_json()
    if data.get('type') == 'url_verification':
        return data.get('challenge'), 200
    print("📥 Received Slack event:", json.dumps(data, indent=2))
    return '', 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)

