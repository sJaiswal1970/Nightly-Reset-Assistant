import os
import time
import json
import psutil
import win32com.client
from datetime import datetime

# ==============================================================================
# ⚙️ USER CONFIGURATION
# ==============================================================================
# 1. Where do you want to save your .psd/.ai backups?
BACKUP_FOLDER = r"C:\Users\YourName\Documents\NightlyBackups"

# 2. Where should the session data be stored? (Keep these in the same folder as this script)
SESSION_FILE = r"C:\Path\To\Script\session.json"
FLAG_FILE = r"C:\Path\To\Script\restore_pending.flag"

# 3. List the executable names of apps you want to reopen (case-insensitive)
#    Use 'Task Manager' -> 'Details' tab to find the exact .exe names.
APPS_TO_TRACK = [
    "chrome.exe", 
    "opera.exe", 
    "WhatsApp.Root.exe", 
    "code.exe", 
    "spotify.exe"
]
# ==============================================================================

def get_open_folders():
    """Scans for open Explorer windows and records their paths."""
    paths = []
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        for window in shell.Windows():
            # Check if the window is a File Explorer window
            if "Explorer" in window.Name:
                try:
                    paths.append(window.Document.Folder.Self.Path)
                except:
                    continue
    except:
        pass
    return paths

def get_running_apps():
    """Scans for running apps from our specific list."""
    running = []
    # Get all process names in lowercase for easier matching
    current_processes = [p.name().lower() for p in psutil.process_iter(['name'])]
    
    for app in APPS_TO_TRACK:
        if app.lower() in current_processes:
            running.append(app)
    return running

def process_adobe_app(app_name, program_id, extension):
    """
    Connects to an active Adobe app, saves all documents to the backup folder,
    records their original paths, and closes the app safely.
    """
    files_to_reopen = []
    
    try:
        # Try to connect to the running app via COM
        try:
            app = win32com.client.GetActiveObject(program_id)
        except:
            return False, [] # App is not running

        print(f"--- Processing {app_name} ---")
        
        # Create a dated folder for backups
        today_folder = os.path.join(BACKUP_FOLDER, datetime.now().strftime("%Y-%m-%d"))
        if not os.path.exists(today_folder):
            os.makedirs(today_folder)

        # Loop through open documents
        if app.Documents.Count > 0:
            for i in range(app.Documents.Count, 0, -1):
                doc = app.Documents(i)
                try:
                    # Create a backup copy
                    clean_name = os.path.splitext(doc.Name)[0]
                    timestamp = datetime.now().strftime("%H-%M-%S")
                    new_filename = f"{timestamp}_{clean_name}{extension}"
                    backup_path = os.path.join(today_folder, new_filename)
                    
                    print(f"Backing up: {doc.Name}")
                    doc.SaveAs(backup_path)
                    
                    # Record the file path for restoration
                    # Try to use original path; if unsaved (Untitled), use the backup path
                    try:
                        files_to_reopen.append(doc.FullName)
                    except:
                        files_to_reopen.append(backup_path)

                    # Close without saving changes (2 = DoNotSaveChanges)
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

    # 2. Handle Adobe Apps
    was_open, files = process_adobe_app("Photoshop", "Photoshop.Application", ".psd")
    session_data["photoshop"] = {"was_open": was_open, "files": files}

    was_open, files = process_adobe_app("Illustrator", "Illustrator.Application", ".ai")
    session_data["illustrator"] = {"was_open": was_open, "files": files}

    # 3. Save Manifest to JSON
    with open(SESSION_FILE, "w") as f:
        json.dump(session_data, f, indent=4)
    
    print("Session recorded successfully.")

    # 4. Set the Flag (The "Sticky Note")
    # This tells the restore script that this was a PLANNED restart
    with open(FLAG_FILE, "w") as f:
        f.write("active")
    
    # 5. Restart Computer
    print("Restarting in 3 seconds...")
    time.sleep(3)
    # /f = force close apps, /r = restart, /t 0 = immediate
    os.system("shutdown /r /f /t 0")