import os
import json
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from boxsdk import JWTAuth, Client
import openpyxl
import dateparser
from datetime import timedelta

load_dotenv()
app = Flask(__name__)

# Load environment variables
box_config_path = os.getenv("BOX_CONFIG_PATH", "box_config.json")
excel_filename = os.getenv("EXCEL_FILE_NAME", "oncall_schedule.xlsx")
folder_id = os.getenv("BOX_FOLDER_ID")

# Authenticate with Box using JWT
auth = JWTAuth.from_settings_file(box_config_path)
client = Client(auth)


def download_excel_file():
    try:
        print(f"Looking for file '{excel_filename}' in folder '{folder_id}'...")
        items = client.folder(folder_id).get_items()
        for item in items:
            print(f"Found item: {item.name}")
            if item.name == excel_filename:
                print("Downloading Excel file...")
                file_content = client.file(item.id).content()
                with open(excel_filename, 'wb') as f:
                    f.write(file_content)
                return True
        print("Excel file not found.")
        return False
    except Exception as e:
        print("Error downloading file from Box:", e)
        return False


def parse_oncall_schedule():
    try:
        wb = openpyxl.load_workbook(excel_filename)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1]]
        data = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_dict = dict(zip(headers, row))
            data.append(row_dict)
        return data
    except Exception as e:
        print("Error parsing Excel file:", e)
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

        if not download_excel_file():
            return jsonify({"error": "Excel file could not be downloaded"}), 500

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
                print("Error parsing entry:", entry, e)
                continue

        return jsonify({"message": "No match found"}), 404

    except Exception as e:
        print("Error in /check-document:", e)
        return jsonify({"error": str(e)}), 500


@app.route('/when-am-i-on-call', methods=['POST'])
def when_am_i_on_call():
    try:
        content = request.get_json()
        name = content.get("name")

        if not name:
            return jsonify({"error": "Missing 'name' field"}), 400

        if not download_excel_file():
            return jsonify({"error": "Excel file could not be downloaded"}), 500

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
            except Exception as e:
                print("Error parsing entry:", entry, e)
                continue

        return jsonify({
            "name": name,
            "upcoming_oncall": upcoming
        })

    except Exception as e:
        print("Error in /when-am-i-on-call:", e)
        return jsonify({"error": str(e)}), 500


@app.route('/slack/events', methods=['POST'])
def slack_events():
    data = request.get_json()
    if data.get('type') == 'url_verification':
        return data.get('challenge'), 200
    print("Received Slack event:", json.dumps(data, indent=2))
    return '', 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)

