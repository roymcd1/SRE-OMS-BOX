from flask import Flask, request, jsonify
from datetime import datetime, timedelta
from boxsdk import JWTAuth, Client
from openpyxl import load_workbook
import io
import requests
import dateparser
import os
import re

app = Flask(__name__)

def get_box_client():
    auth = JWTAuth.from_settings_file('box_config.json')
    access_token = auth.authenticate_instance()
    return Client(auth)

def get_schedule_file():
    client = get_box_client()
    file_id = os.environ.get("BOX_FILE_ID")
    file_content = client.file(file_id).content()
    return load_workbook(filename=io.BytesIO(file_content), data_only=True)

def parse_date_range(raw_input):
    if not raw_input:
        return None, None

    query = raw_input.lower().strip()
    query = query.replace("on-call", "on call").replace("oncall", "on call")
    query = re.sub(r'[^\w\s]', '', query)  # remove punctuation

    match = re.search(
        r"(this|next|last)?\s?(week|month|monday|tuesday|wednesday|thursday|friday|saturday|sunday|today|tomorrow|yesterday)",
        query
    )

    if not match:
        return None, None

    phrase = match.group().strip()
    parsed_date = dateparser.parse(phrase, settings={'RELATIVE_BASE': datetime.today()})
    if not parsed_date:
        return None, None

    if 'week' in phrase:
        start = parsed_date - timedelta(days=parsed_date.weekday())
        end = start + timedelta(days=6)
    else:
        start = parsed_date.date()
        end = start

    return start, end

@app.route("/check-document", methods=["POST"])
def check_document():
    if request.content_type == "application/x-www-form-urlencoded":
        query = request.form.get("text")
        slack_mode = True
    else:
        data = request.get_json()
        query = data.get("week_query")
        slack_mode = False

    start, end = parse_date_range(query)
    if not start or not end:
        message = f"Could not understand week_query: '{query}'"
        return jsonify(text=message) if slack_mode else jsonify({"error": message}), 400

    wb = get_schedule_file()
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_start, row_end, primary, secondary = row
        if row_start <= start <= row_end:
            result = {
                "start": str(row_start),
                "end": str(row_end),
                "names": {
                    "primary": primary,
                    "secondary": secondary
                }
            }
            if slack_mode:
                message = (
                    f"📅 On-call from *{row_start}* to *{row_end}*\n"
                    f"👨‍🚒 Primary: *{primary}*\n"
                    f"🧯 Secondary: *{secondary}*"
                )
                return jsonify(text=message)
            return jsonify(result)

    message = "No schedule found for that week."
    return jsonify(text=message) if slack_mode else jsonify({"error": message}), 404

@app.route("/when-am-i-on-call", methods=["POST"])
def when_am_i_on_call():
    if request.content_type == "application/x-www-form-urlencoded":
        name = request.form.get("text")
        slack_mode = True
    else:
        data = request.get_json()
        name = data.get("name")
        slack_mode = False

    if not name:
        message = "No name provided. Try `/whenami Jane Doe`"
        return jsonify(text=message) if slack_mode else jsonify({"error": message}), 400

    wb = get_schedule_file()
    ws = wb.active
    today = datetime.today().date()
    upcoming = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_start, row_end, primary, secondary = row
        if row_end >= today and (primary == name or secondary == name):
            upcoming.append({
                "start": str(row_start),
                "end": str(row_end),
                "primary": primary,
                "secondary": secondary
            })

    if not upcoming:
        message = f"No upcoming on-call slots found for *{name}*."
        return jsonify(text=message) if slack_mode else jsonify({"name": name, "upcoming_oncall": []}), 404

    if slack_mode:
        message = f"📟 On-call slots for *{name}*:\n"
        for slot in upcoming:
            message += (
                f"• {slot['start']} → {slot['end']} "
                f"(👨‍🚒 Primary: {slot['primary']}, 🧯 Secondary: {slot['secondary']})\n"
            )
        return jsonify(text=message)

    return jsonify({"name": name, "upcoming_oncall": upcoming})

if __name__ == "__main__":
    app.run(debug=True)

