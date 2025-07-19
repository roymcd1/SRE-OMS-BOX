import os
from flask import Flask, request, jsonify
from boxsdk import JWTAuth, Client
from dotenv import load_dotenv
from datetime import datetime, timedelta
import openpyxl
from io import BytesIO
import dateparser

# Load environment variables
load_dotenv()
BOX_FILE_ID = os.getenv("BOX_FILE_ID")
assert BOX_FILE_ID, "BOX_FILE_ID must be set in .env"

# Box authentication
auth = JWTAuth.from_settings_file("box_config.json")
auth.authenticate_instance()
client = Client(auth)

# Flask app
app = Flask(__name__)

def download_excel_file(file_id):
    try:
        file_content = client.file(file_id).content()
        return BytesIO(file_content)
    except Exception as e:
        print(f"🔥 Error downloading file: {e}")
        return None

def parse_week_query(week_query):
    parsed_date = dateparser.parse(week_query)
    if not parsed_date:
        return None, None
    start_of_week = parsed_date - timedelta(days=parsed_date.weekday())
    end_of_week = start_of_week + timedelta(days=6)
    return start_of_week.date(), end_of_week.date()

def extract_names(excel_bytes, start_date, end_date):
    wb = openpyxl.load_workbook(excel_bytes)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2, values_only=True):
        try:
            excel_start = row[1].date() if hasattr(row[1], 'date') else row[1]
            excel_end = row[2].date() if hasattr(row[2], 'date') else row[2]
        except Exception:
            continue

        if excel_start == start_date and excel_end == end_date:
            return {
                "primary": row[3],
                "secondary": row[5]  # Assuming column F is Secondary
            }
    return {"primary": None, "secondary": None}

def find_upcoming_oncall(excel_bytes, person_name, today=None):
    if today is None:
        today = datetime.today().date()

    wb = openpyxl.load_workbook(excel_bytes)
    sheet = wb.active
    matches = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        try:
            excel_start = row[1].date() if hasattr(row[1], 'date') else row[1]
            excel_end = row[2].date() if hasattr(row[2], 'date') else row[2]
            primary = row[3]
            secondary = row[5]
        except Exception:
            continue

        if not excel_start or excel_start < today:
            continue

        if person_name.lower() in str(primary).lower() or person_name.lower() in str(secondary).lower():
            matches.append({
                "start": str(excel_start),
                "end": str(excel_end),
                "primary": primary,
                "secondary": secondary,
            })

        if len(matches) == 3:
            break

    return matches

@app.route("/check-document", methods=["POST"])
def check_document():
    data = request.get_json()
    week_query = data.get("week_query")

    if not week_query:
        return jsonify({"error": "Missing 'week_query' field"}), 400

    start_date, end_date = parse_week_query(week_query)
    if not start_date or not end_date:
        return jsonify({"error": "Could not understand week_query"}), 400

    excel_bytes = download_excel_file(BOX_FILE_ID)
    if not excel_bytes:
        return jsonify({"error": "Failed to download Excel file"}), 500

    names = extract_names(excel_bytes, start_date, end_date)
    return jsonify({
        "start": str(start_date),
        "end": str(end_date),
        "names": names
    })

@app.route("/when-am-i-on-call", methods=["POST"])
def when_am_i_on_call():
    data = request.get_json()
    person = data.get("name")

    if not person:
        return jsonify({"error": "Missing 'name' field"}), 400

    excel_bytes = download_excel_file(BOX_FILE_ID)
    if not excel_bytes:
        return jsonify({"error": "Failed to download Excel file"}), 500

    upcoming = find_upcoming_oncall(excel_bytes, person)
    return jsonify({
        "name": person,
        "upcoming_oncall": upcoming
    })

if __name__ == "__main__":
    print(f"📦 BOX_FILE_ID = {BOX_FILE_ID}")
    app.run(host="0.0.0.0", port=8080)

