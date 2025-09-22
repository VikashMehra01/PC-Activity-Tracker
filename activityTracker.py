import win32gui
import win32process
import psutil
import time
from datetime import datetime
import csv
import os
import json

# --- Configuration ---


desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
DATA_FOLDER = os.path.join(desktop_path, 'PC_Activity_Logs')
DAILY_CSV_FOLDER = os.path.join(DATA_FOLDER, 'DailyActivityLogs')
DAILY_SUMMARY_FOLDER = os.path.join(DATA_FOLDER, 'DailyUsage')



# DAILY_CSV_FOLDER = 'DailyActivityLogs'
# DAILY_SUMMARY_FOLDER = 'DailyUsage'




CSV_HEADER = ['start_time', 'process_name', 'window_title', 'duration_seconds']
LOG_INTERVAL_SECONDS = 25
SUMMARY_SAVE_INTERVAL_SECONDS = 60
BROWSERS = ["firefox.exe", "chrome.exe", "msedge.exe"]

def get_active_window_info():
    """Gets the process name and window title of the currently active window."""
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        process_name = process.name()
        window_title = win32gui.GetWindowText(hwnd)
        return process_name, window_title
    except (psutil.NoSuchProcess, psutil.AccessDenied, win32gui.error):
        return None, None

def write_log_entry(log_entry):
    """Determines the daily CSV filename and appends a single entry."""
    today_str = datetime.now().strftime('%Y-%m-%d')
    filename = os.path.join(DAILY_CSV_FOLDER, f"{today_str}-Activity.csv")
    
    file_exists = os.path.exists(filename)
    with open(filename, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not file_exists or os.path.getsize(filename) == 0:
            writer.writerow(CSV_HEADER)
        writer.writerow(log_entry)
    
def save_summary(summary_data, filename):
    """Calculates grand total, wraps data, and saves to the JSON file."""
    grand_total = sum(proc.get("total_time", 0) for proc in summary_data.values())

    output_data = {
        "grand_total_seconds": grand_total,
        "processes": summary_data
    }

    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=4)

def update_summary(summary, proc_name, title, duration_sec):
    """Updates the nested summary dictionary."""
    if proc_name not in summary:
        summary[proc_name] = {"total_time": 0, "details": {}}
    
    summary[proc_name]["total_time"] += duration_sec
    
    details = summary[proc_name]["details"]
    details[title] = details.get(title, 0) + duration_sec

def load_summary(filename):
    """Loads the summary data, unwrapping it from the new structure."""
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if "processes" in data:
                return data["processes"]
            else:
                return data
    except (FileNotFoundError, json.JSONDecodeError):
        return {} 

def main():

    print("🚀 Starting duration tracker... Press Ctrl+C to stop.") 
    """Main loop to track activity and log its duration."""
    os.makedirs(DAILY_SUMMARY_FOLDER, exist_ok=True)
    os.makedirs(DAILY_CSV_FOLDER, exist_ok=True)

    today_str = datetime.now().strftime('%Y-%m-%d')
    summary_filename = os.path.join(DAILY_SUMMARY_FOLDER, f"{today_str}.json")
    time_summary = load_summary(summary_filename)
    current_activity = None
    last_save_time = datetime.now()

    try:
        while True:
            process_name, window_title = get_active_window_info()

            if not window_title:
                time.sleep(LOG_INTERVAL_SECONDS)
                continue

            effective_title = window_title
            if process_name in BROWSERS:
                suffixes_to_remove = [
                    " - Mozilla Firefox", " — Mozilla Firefox",
                    " - Google Chrome", " — Google Chrome",
                    " - Microsoft Edge", " — Microsoft Edge"
                ]
                temp_title = window_title
                for suffix in suffixes_to_remove:
                    if temp_title.endswith(suffix):
                        temp_title = temp_title.removesuffix(suffix).strip()
                        break

                separator=[' | ', ' - ' ,' — ']
                found=False
                for sep in separator:
                    if sep in temp_title:
                        effective_title=temp_title.rsplit(sep,1)[-1].strip()
                        found=True
                        break # Use the first separator found
                if not found:
                    effective_title = temp_title
            
            elif process_name == "Code.exe":
                temp_title = window_title
                suffix = " - Visual Studio Code"
                if temp_title.endswith(suffix):
                    temp_title = temp_title.removesuffix(suffix).strip()

                parts = temp_title.split(' - ')
                if parts:
                    effective_title = parts[0].lstrip('● ').strip()
            
            if current_activity is None:
                current_activity = {"start_time": datetime.now(), "process_name": process_name, "window_title": effective_title}
            
            elif effective_title != current_activity["window_title"] or process_name != current_activity["process_name"]:
                end_time = datetime.now()
                duration = end_time - current_activity["start_time"]
                duration_seconds = int(duration.total_seconds())
                
                if duration_seconds > 0: # Only log activities with a duration
                    log_entry = [
                        current_activity["start_time"].strftime('%Y-%m-%d %H:%M:%S'),
                        current_activity["process_name"],
                        current_activity["window_title"],
                        duration_seconds
                    ]
                    write_log_entry(log_entry)
                    update_summary(time_summary, current_activity["process_name"], current_activity["window_title"], duration_seconds)
                
                current_activity = {"start_time": datetime.now(), "process_name": process_name, "window_title": effective_title}

            time.sleep(LOG_INTERVAL_SECONDS)

            if (datetime.now() - last_save_time).total_seconds() >= SUMMARY_SAVE_INTERVAL_SECONDS:
                save_summary(time_summary, summary_filename)
                last_save_time = datetime.now()

    except KeyboardInterrupt:
        if current_activity:
            end_time = datetime.now()
            duration = end_time - current_activity["start_time"]
            duration_seconds = int(duration.total_seconds())
            
            if duration_seconds > 0:
                log_entry = [
                    current_activity["start_time"].strftime('%Y-%m-%d %H:%M:%S'),
                    current_activity["process_name"],
                    current_activity["window_title"],
                    duration_seconds
                ]
                write_log_entry(log_entry)
                update_summary(time_summary, current_activity["process_name"], current_activity["window_title"], duration_seconds)

        save_summary(time_summary, summary_filename)
       
    print(f"\n🛑 Tracker stopped..")  

if __name__ == "__main__":
    main()