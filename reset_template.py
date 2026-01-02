import os
import time
import json
import psutil
import win32com.client
from datetime import datetime

# === LOAD CONFIGURATION ===
# We determine the folder where this script is running
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")
SESSION_FILE = os.path.join(SCRIPT_DIR, "session.json")
FLAG_FILE = os.path.join(SCRIPT_DIR, "restore_pending.flag")

# Check if setup.py has been run
if not os.path.exists(CONFIG_PATH):
    print("❌ Error: config.json not found.")
    print("   Please run 'setup.py' first to configure your settings.")
    time.sleep(10)
    exit()

# Load user settings
with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

BACKUP_FOLDER = config.get("backup_folder")
APPS_TO_TRACK = config.get("apps_to_track", [])

def get_open_folders():
    """Scans for open Explorer windows and records their paths."""
    paths = []
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        for window in shell.Windows():
            if "Explorer" in window.Name:
                try:
                    paths.append(window.Document.Folder.Self.Path)
                except:
                    continue
    except:
        pass
    return paths

def get_running_apps():
    """Scans for running apps based on the config list."""
    running = []
    current_processes = [p.name().lower() for p in psutil.process_iter(['name'])]
    
    for app in APPS_TO_TRACK:
        if app.lower() in current_processes:
            running.append(app)
    return running

def process_adobe_app(app_name, program_id, extension):
    """Backs up active documents and closes the Adobe app."""
    files_to_reopen = []
    
    try:
        try:
            app = win32com.client.GetActiveObject(program_id)
        except:
            return False, [] # App is not running

        print(f"--- Processing {app_name} ---")
        
        # Create backup folder
        today_folder = os.path.join(BACKUP_FOLDER, datetime.now().strftime("%Y-%m-%d"))
        if not os.path.exists(today_folder):
            os.makedirs(today_folder)

        # Loop through open documents
        if app.Documents.Count > 0:
            for i in range(app.Documents.Count, 0, -1):
                doc = app.Documents(i)
                try:
                    # Save Backup
                    clean_name = os.path.splitext(doc.Name)[0]
                    timestamp = datetime.now().strftime("%H-%M-%S")
                    new_filename = f"{timestamp}_{clean_name}{extension}"
                    backup_path = os.path.join(today_folder, new_filename)
                    
                    print(f"Backing up: {doc.Name}")
                    doc.SaveAs(backup_path)
                    
                    # Record file path
                    try:
                        files_to_reopen.append(doc.FullName)
                    except:
                        files_to_reopen.append(backup_path)

                    doc.Close(2) 
                except Exception as e:
                    print(f"Error saving {doc.Name}: {e}")
        
        app.Quit()
        return True, files_to_reopen

    except Exception as e:
        print(f"Error connecting to {app_name}: {e}")
        return False, []

if __name__ == "__main__":
    os.system("title Nightly Reset Assistant")
    print("Scanning system state...")

    # 1. Capture State
    session_data = {
        "folders": get_open_folders(),
        "apps": get_running_apps(),
        "photoshop": {"was_open": False, "files": []},
        "illustrator": {"was_open": False, "files": []}
    }

    # 2. Handle Adobe
    was_open, files = process_adobe_app("Photoshop", "Photoshop.Application", ".psd")
    session_data["photoshop"] = {"was_open": was_open, "files": files}

    was_open, files = process_adobe_app("Illustrator", "Illustrator.Application", ".ai")
    session_data["illustrator"] = {"was_open": was_open, "files": files}

    # 3. Save Session
    with open(SESSION_FILE, "w") as f:
        json.dump(session_data, f, indent=4)
    
    print("Session recorded successfully.")

    # 4. Set Flag
    with open(FLAG_FILE, "w") as f:
        f.write("active")
    
    # 5. Restart
    print("Restarting in 3 seconds...")
    time.sleep(3)
    os.system("shutdown /r /f /t 0")