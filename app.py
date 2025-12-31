from flask import Flask, send_from_directory, request, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime, timedelta
import datetime as dtmod
import time
import os

app = Flask(__name__, static_folder='.')
CORS(app)

# === CONFIG ===
EXCEL_FILE = "Deep Focused Work Tracker(Winter 2026).xlsx"
CURRENT_WEEK = "Week 1"

CATEGORIES = [
    "Classes",
    "Personal",
    "Coursework",
    "Altium",
    "Machine Learning",
    "Academic Clubs",
    "Combat Clubs"
]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# === TIMER STATE ===
timer_state = {
    "start_time": None,
    "active_category": None
}

# === EXCEL HELPERS ===

def get_existing_timedelta(cell_value):
    if isinstance(cell_value, (int, float)):
        return timedelta(seconds=float(cell_value) * 24 * 3600)

    if isinstance(cell_value, dtmod.time):
        return timedelta(hours=cell_value.hour, minutes=cell_value.minute, seconds=cell_value.second)

    try:
        s = str(cell_value).strip()
        if s == "" or s in ("0", "0.0", "0:00:00"):
            return timedelta()
        f = float(s)
    except Exception:
        return parse_time_string(str(cell_value))
    else:
        if abs(f) < 1:
            return timedelta(seconds=f * 24 * 3600)
        return timedelta(seconds=f * 3600)

def get_sheet():
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel file '{EXCEL_FILE}' not found.")
    wb = openpyxl.load_workbook(EXCEL_FILE)
    if CURRENT_WEEK not in wb.sheetnames:
        raise ValueError(f"Sheet '{CURRENT_WEEK}' not found in workbook.")
    return wb, wb[CURRENT_WEEK]

def parse_time_string(time_str):
    if not time_str or time_str.strip() == "" or time_str == "0:00:00":
        return timedelta()

    s = time_str.strip()
    try:
        f = float(s)
    except Exception:
        pass
    else:
        if abs(f) < 1:
            return timedelta(seconds=f * 24 * 3600)
        return timedelta(seconds=f * 3600)

    parts = [p.strip() for p in s.split(":") if p.strip() != ""]
    try:
        if len(parts) == 3:
            h = int(float(parts[0]))
            m = int(float(parts[1]))
            sec = int(float(parts[2]))
        elif len(parts) == 2:
            h = 0
            m = int(float(parts[0]))
            sec = int(float(parts[1]))
        elif len(parts) == 1:
            sec = int(float(parts[0]))
            h = 0
            m = 0
        else:
            raise ValueError(f"Unrecognized time format: '{time_str}'")
    except Exception as e:
        raise ValueError(f"Couldn't parse time string '{time_str}': {e}")

    return timedelta(hours=h, minutes=m, seconds=sec)

def format_timedelta(td):
    total_seconds = int(td.total_seconds())
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h}:{m:02}:{s:02}"

def update_excel(category, day, elapsed):
    wb, ws = get_sheet()

    cat_row = CATEGORIES.index(category) + 2
    day_col = DAYS.index(day) + 2

    cell = ws.cell(row=cat_row, column=day_col)

    td_existing = get_existing_timedelta(cell.value)
    existing_days = td_existing.total_seconds() / (24*3600)

    elapsed_days = elapsed.total_seconds() / (24*3600)

    new_value = existing_days + elapsed_days
    cell.value = new_value
    cell.number_format = '[h]:mm:ss'
    cell.alignment = Alignment(horizontal='right')

    wb.save(EXCEL_FILE)
    
    total_seconds = new_value * 24 * 60 * 60
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)
    
    return f"{hours}:{minutes:02}:{seconds:02}"

# === API ROUTES ===

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/api/start', methods=['POST'])
def api_start():
    data = request.json
    category = data.get('category')
    
    if timer_state["active_category"]:
        return jsonify({
            "success": False,
            "message": f"Timer already running for '{timer_state['active_category']}'. Stop it first."
        }), 400
    
    if category not in CATEGORIES:
        return jsonify({
            "success": False,
            "message": f"Invalid category. Valid options: {CATEGORIES}"
        }), 400
    
    timer_state["start_time"] = time.time()
    timer_state["active_category"] = category
    
    today = DAYS[datetime.today().weekday()]
    
    return jsonify({
        "success": True,
        "message": f"Started tracking '{category}' on {today}",
        "category": category,
        "day": today,
        "start_time": timer_state["start_time"]
    })

@app.route('/api/stop', methods=['POST'])
def api_stop():
    if not timer_state["active_category"]:
        return jsonify({
            "success": False,
            "message": "No timer running."
        }), 400
    
    elapsed = timedelta(seconds=time.time() - timer_state["start_time"])
    today = DAYS[datetime.today().weekday()]
    category = timer_state["active_category"]
    
    try:
        new_total = update_excel(category, today, elapsed)
        
        timer_state["start_time"] = None
        timer_state["active_category"] = None
        
        return jsonify({
            "success": True,
            "message": f"Stopped '{category}'",
            "duration": format_timedelta(elapsed),
            "new_total": new_total,
            "category": category
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"Error updating Excel: {str(e)}"
        }), 500

@app.route('/api/status', methods=['GET'])
def api_status():
    if timer_state["active_category"]:
        elapsed = timedelta(seconds=time.time() - timer_state["start_time"])
        return jsonify({
            "active": True,
            "category": timer_state["active_category"],
            "elapsed": format_timedelta(elapsed),
            "elapsed_seconds": int(elapsed.total_seconds())
        })
    else:
        return jsonify({
            "active": False,
            "message": "No timer currently running."
        })

@app.route('/api/categories', methods=['GET'])
def api_categories():
    return jsonify({
        "categories": CATEGORIES,
        "days": DAYS,
        "current_week": CURRENT_WEEK
    })

if __name__ == '__main__':
    app.run(debug=True, port=5000)