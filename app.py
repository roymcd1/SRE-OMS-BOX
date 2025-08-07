import os
from datetime import date, datetime, timedelta
from flask import Flask, request, jsonify, send_file
from dotenv import load_dotenv
from boxsdk import JWTAuth, Client
import openpyxl
import dateparser
import subprocess

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
CACHE_DURATION = timedelta(hours=1)  # Cache TTL for the Excel data
last_download_time = None           # Timestamp of last Excel download
schedule_cache = None               # Cached schedule data (list of rows as dicts)
last_pdf_time = None                # Timestamp of last PDF generation

def to_date(dt):
    """Convert a datetime or date-like object to a date object."""
    if dt is None:
        return None
    return dt.date() if hasattr(dt, "date") else dt

def download_excel():
    """Download the Excel file from Box to local disk."""
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
    """Load and parse the oncall Excel file into a list of dictionaries."""
    try:
        wb = openpyxl.load_workbook("oncall.xlsx", read_only=True, data_only=True)
        sheet = wb.active
        headers = [c.value for c in sheet[1]]
        schedule_data = [dict(zip(headers, row))
                         for row in sheet.iter_rows(min_row=2, values_only=True)]
        wb.close()
        return schedule_data
    except Exception as e:
        print("❌ parse_schedule:", e)
        return []

def get_schedule_data():
    """
    Retrieve the on-call schedule data, using a cached copy if available and fresh.
    Returns a list of schedule rows (each row is a dict) or None if download failed.
    """
    global last_download_time, schedule_cache
    now = datetime.now()
    try:
        # Refresh data if not cached or cache expired
        if schedule_cache is None or last_download_time is None or (now - last_download_time) > CACHE_DURATION:
            print("🔄 Refreshing schedule data...")
            result = download_excel()
            if result is not True:
                # Download failed; return None to indicate an error
                return None
            data = parse_schedule()
            # Parse date fields ("Start" and "End") to date objects for faster comparisons
            for row in data:
                start_val = row.get("Start")
                end_val   = row.get("End")
                if start_val:
                    # If not already a date/datetime, parse it
                    if not isinstance(start_val, date):
                        start_val = dateparser.parse(str(start_val))
                    row["Start"] = to_date(start_val)
                else:
                    row["Start"] = None
                if end_val:
                    if not isinstance(end_val, date):
                        end_val = dateparser.parse(str(end_val))
                    row["End"] = to_date(end_val)
                else:
                    row["End"] = None
            schedule_cache = data
            last_download_time = now
            print(f"✅ Schedule data refreshed at {last_download_time}")
        else:
            print(f"⚡ Using cached schedule data (last refreshed at {last_download_time})")
        return schedule_cache
    except Exception as e:
        print("❌ get_schedule_data:", e)
        return None

# ----------------- 5. FLASK -------------------
app = Flask(__name__)

@app.route('/')
def home():
    return "Box On-Call App ✅", 200

@app.route('/check-document', methods=['POST'])
def check_document():
    try:
        req_data = request.get_json() or {}
        week_query = req_data.get("week_query")
        if not week_query:
            return jsonify({"error": "Missing 'week_query'"}), 400

        schedule_data = get_schedule_data()
        if schedule_data is None:
            return jsonify({"error": "Excel download failed"}), 500

        target_dt = to_date(dateparser.parse(week_query))
        if not target_dt:
            return jsonify({"error": "Bad date"}), 400

        for row in schedule_data:
            start = row.get("Start")
            end   = row.get("End")
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
        req_data = request.get_json() or {}
        name = req_data.get("name")
        if not name:
            return jsonify({"error": "Missing 'name'"}), 400

        schedule_data = get_schedule_data()
        if schedule_data is None:
            return jsonify({"error": "Excel download failed"}), 500

        today = date.today()
        upcoming = []
        for row in schedule_data:
            start = row.get("Start")
            end   = row.get("End")
            if start and end and end >= today and (row.get("Primary") == name or row.get("Secondary") == name):
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

@app.route('/rota-pdf', methods=['GET'])
def rota_pdf():
    try:
        global last_pdf_time
        # Ensure we have an up-to-date Excel file (use cache if within 1 hour)
        schedule_data = get_schedule_data()
        if schedule_data is None:
            return "Failed to download Excel file", 500

        input_path = os.path.abspath("oncall.xlsx")
        output_dir = os.path.abspath(".")
        pdf_path = os.path.join(output_dir, "oncall.pdf")

        # If a recent PDF exists and the Excel data hasn't been refreshed since it was generated, use it
        if last_pdf_time and last_download_time and (datetime.now() - last_pdf_time) < CACHE_DURATION and last_pdf_time >= last_download_time:
            if os.path.exists(pdf_path):
                print(f"⚡ Using cached PDF (last generated at {last_pdf_time})")
                return send_file(pdf_path, as_attachment=True)

        print("🌀 Converting Excel to PDF...")
        result = subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, input_path
        ], capture_output=True)

        if result.returncode != 0:
            print("❌ PDF conversion failed:", result.stderr.decode())
            return "PDF conversion failed", 500
        if not os.path.exists(pdf_path):
            return "PDF not created", 500

        last_pdf_time = datetime.now()
        print("✅ PDF ready:", pdf_path)
        return send_file(pdf_path, as_attachment=True)
    except Exception as e:
        print("❌ /rota-pdf:", e)
        return str(e), 500

@app.route('/slack/events', methods=['POST'])
def slack_events():
    data = request.get_json() or {}
    if data.get('type') == 'url_verification':
        return data.get('challenge'), 200
    print("📥 Slack event:", data)
    return '', 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)

