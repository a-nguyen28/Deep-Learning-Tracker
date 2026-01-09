from flask import Flask, send_from_directory, request, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime, timedelta
import datetime as dtmod
import time
import os
import logging
import webbrowser
import threading

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

app = Flask(__name__, static_folder='.')
CORS(app)

# === CONFIG ===
EXCEL_FILE = "Deep Focused Work Tracker(Winter 2026).xlsx"

# IMPORTANT: Set this to the first Monday of your semester
SEMESTER_START = datetime(2026, 1, 5)  # CHANGE THIS DATE!

def get_current_week():
    """Calculate current week number based on semester start date"""
    today = datetime.now()
    days_elapsed = (today - SEMESTER_START).days
    
    # If before semester start, default to Week 1
    if days_elapsed < 0:
        logging.warning(f"Current date is before semester start ({SEMESTER_START.date()}). Using Week 1.")
        return "Week 1"
    
    week_num = (days_elapsed // 7) + 1
    logging.info(f"Calculated current week: Week {week_num}")
    return f"Week {week_num}"

CURRENT_WEEK = get_current_week()

CATEGORIES = [
    "Classes",
    "Personal",
    "Coursework",
    "Altium",
    "UAV Lab",
    "Machine Learning",
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
    """Parse Excel cell value into timedelta, handling multiple formats"""
    
    # Empty cell
    if cell_value is None or cell_value == "":
        return timedelta()
    
    # Excel stores time as decimal fraction of days
    if isinstance(cell_value, (int, float)):
        days = float(cell_value)
        logging.debug(f"Parsed numeric value {cell_value} as {days} days")
        return timedelta(seconds=days * 24 * 3600)

    # Excel time object
    if isinstance(cell_value, dtmod.time):
        td = timedelta(hours=cell_value.hour, minutes=cell_value.minute, seconds=cell_value.second)
        logging.debug(f"Parsed time object {cell_value} as {td}")
        return td

    # Try parsing as string
    try:
        s = str(cell_value).strip()
        if s == "" or s in ("0", "0.0", "0:00:00"):
            return timedelta()
        
        # Try as float first
        try:
            f = float(s)
            if abs(f) < 1:
                # Small number = fraction of day
                return timedelta(seconds=f * 24 * 3600)
            else:
                # Large number = hours
                return timedelta(seconds=f * 3600)
        except ValueError:
            # Not a number, try time string format
            return parse_time_string(s)
            
    except Exception as e:
        logging.error(f"Failed to parse cell value '{cell_value}': {e}")
        return timedelta()

def get_sheet():
    """Load Excel workbook and return current week's sheet"""
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel file '{EXCEL_FILE}' not found in current directory.")
    
    wb = openpyxl.load_workbook(EXCEL_FILE)
    
    if CURRENT_WEEK not in wb.sheetnames:
        available_sheets = ", ".join(wb.sheetnames)
        raise ValueError(
            f"Sheet '{CURRENT_WEEK}' not found in workbook. "
            f"Available sheets: {available_sheets}"
        )
    
    return wb, wb[CURRENT_WEEK]

def parse_time_string(time_str):
    """Parse time string in formats like '2:30:15' or '150:45:30'"""
    if not time_str or time_str.strip() == "" or time_str == "0:00:00":
        return timedelta()

    s = time_str.strip()
    
    # Try as float
    try:
        f = float(s)
        if abs(f) < 1:
            return timedelta(seconds=f * 24 * 3600)
        return timedelta(seconds=f * 3600)
    except ValueError:
        pass

    # Parse as H:M:S format
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
    """Format timedelta as H:MM:SS"""
    total_seconds = int(td.total_seconds())
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h}:{m:02}:{s:02}"

def update_excel(category, day, elapsed):
    """Update Excel cell with new time value"""
    try:
        wb, ws = get_sheet()

        cat_row = CATEGORIES.index(category) + 2
        day_col = DAYS.index(day) + 2

        cell = ws.cell(row=cat_row, column=day_col)

        # Get existing time
        td_existing = get_existing_timedelta(cell.value)
        existing_seconds = td_existing.total_seconds()
        
        # Add new elapsed time
        elapsed_seconds = elapsed.total_seconds()
        total_seconds = existing_seconds + elapsed_seconds
        
        # Convert to days for Excel (Excel stores as decimal fraction of days)
        new_value = total_seconds / (24 * 3600)
        
        logging.info(
            f"Updating {category}/{day}: "
            f"existing={format_timedelta(td_existing)}, "
            f"elapsed={format_timedelta(elapsed)}, "
            f"new_total={format_timedelta(timedelta(seconds=total_seconds))}"
        )
        
        cell.value = new_value
        cell.number_format = '[h]:mm:ss'  # [h] allows hours > 24
        cell.alignment = Alignment(horizontal='right')

        wb.save(EXCEL_FILE)
        
        # Format for display
        hours = int(total_seconds // 3600)
        minutes = int((total_seconds % 3600) // 60)
        seconds = int(total_seconds % 60)
        
        return f"{hours}:{minutes:02}:{seconds:02}"
        
    except Exception as e:
        logging.error(f"Error updating Excel: {e}", exc_info=True)
        raise

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
    
    logging.info(f"Started timer for '{category}' on {today}")
    
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
        
        logging.info(f"Stopped timer for '{category}': duration={format_timedelta(elapsed)}, total={new_total}")
        
        return jsonify({
            "success": True,
            "message": f"Stopped '{category}'",
            "duration": format_timedelta(elapsed),
            "new_total": new_total,
            "category": category
        })
    except Exception as e:
        logging.error(f"Error stopping timer: {e}", exc_info=True)
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

def open_browser():
    """Open browser after brief delay to ensure server is running"""
    time.sleep(1.5)
    webbrowser.open('http://localhost:5000')

if __name__ == '__main__':
    print(f"\n{'='*60}")
    print(f"  Deep Focused Work Tracker")
    print(f"{'='*60}")
    print(f"  Current Week: {CURRENT_WEEK}")
    print(f"  Semester Start: {SEMESTER_START.date()}")
    print(f"  Excel File: {EXCEL_FILE}")
    print(f"{'='*60}")
    print(f"  Starting server at http://localhost:5000")
    print(f"  Browser will open automatically...")
    print(f"  Press Ctrl+C to stop\n")
    
    # Open browser in background thread
    threading.Thread(target=open_browser, daemon=True).start()
    
    # Start Flask server
    app.run(debug=True, port=5000, use_reloader=False)