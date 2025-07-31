import os
import json
from flask import Flask, request, jsonify
from dotenv import load_dotenv
from boxsdk import JWTAuth, Client
import openpyxl
import dateparser

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
    items = client.folder(folder_id).get_items()
    for item in items:
        if item.name == excel_filename:
            file_content = client.file(item.id).content()
            with open(excel_filename, 'wb') as f:
                f.write(file_content)
            return True
    return False


def parse_oncall_schedule():
    wb = openpyxl.load_workbook(excel_filename)
    sheet = wb.active
    headers = [cell.value for cell in sheet[1]]
    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        row_dict = dict(zip(headers, row))
        data.append(row_dict)
    return data


def find_week_range(week_query):
    dt = dateparser.parse(week_query)
    if not dt:
        return None, None
    start = dt - dt.weekday() * timedelta(days=1)  # Monday
    end = start + timedelta(days=6)  # Sunday
    return start.date(), end.date()


@app.route('/check-document', methods=['POST'])
def check_document():
    content = request.get_json()
    week_query = content.get("week_query")

    if not week_query:
        return jsonify({"error": "Missing 'week_query' field"}), 400

    download_excel_file()
    data = parse_oncall_schedule()

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
            continue

    return jsonify({"message": "No match found"}), 404


@app.route('/when-am-i-on-call', methods=['POST'])
def when_am_i_on_call():
    content = request.get_json()
    name = content.get("name")

    if not name:
        return jsonify({"error": "Missing 'name' field"}), 400

    download_excel_file()
    data = parse_oncall_schedule()

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
            continue

    return jsonify({
        "name": name,
        "upcoming_oncall": upcoming
    })


@app.route('/slack/events', methods=['POST'])
def slack_events():
    data = request.get_json()

    # Slack Event API verification
    if data.get('type') == 'url_verification':
        return data.get('challenge'), 200

    # Handle other Slack events here if needed
    return '', 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)

