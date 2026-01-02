import os
import time
import json
import subprocess
import win32com.client

# === LOAD CONFIGURATION ===
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")
SESSION_FILE = os.path.join(SCRIPT_DIR, "session.json")
FLAG_FILE = os.path.join(SCRIPT_DIR, "restore_pending.flag")

# Check if setup.py has been run
if not os.path.exists(CONFIG_PATH):
    # If no config exists, we can't run. Exit silently (since this runs on startup).
    exit()

with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

PHOTOSHOP_PATH = config.get("photoshop_path")
ILLUSTRATOR_PATH = config.get("illustrator_path")
OPERA_PATH = config.get("opera_path")

def launch_app_reliable(exe_name):
    """Launches apps using config paths or standard commands."""
    try:
        if "whatsapp" in exe_name.lower():
             os.system("start whatsapp:")
        
        elif "opera" in exe_name.lower() and OPERA_PATH:
            # Use the specific user-defined path
            subprocess.Popen(f'"{OPERA_PATH}"', shell=True)
            
        else:
             subprocess.Popen(f"start {exe_name}", shell=True)
    except Exception as e:
        print(f"Failed to launch {exe_name}: {e}")

def restore_adobe_files(app_name, program_id, file_list, exe_path):
    """Launches Adobe app via path, waits for splash screen, opens files."""
    print(f"Restoring {app_name}...")
    
    # 1. LAUNCH
    try:
        subprocess.Popen(f'"{exe_path}"', shell=True)
    except Exception as e:
        print(f"   Error launching {app_name}: {e}")
        return

    if not file_list:
        return
    
    # 2. WAIT LOOP (60s max)
    print(f"   Waiting for {app_name} to respond...")
    app = None
    for attempt in range(60): 
        try:
            app = win32com.client.GetActiveObject(program_id)
            _ = app.Name 
            break
        except:
            time.sleep(1)
    
    if app:
        print(f"   Connected! Waiting 5s for splash screen...")
        time.sleep(5)
        
        app.Visible = True
        app.UserInteractionLevel = 1 
        
        for file_path in file_list:
            try:
                print(f"   Opening: {os.path.basename(file_path)}")
                app.Open(file_path)
                time.sleep(1) 
            except Exception as e:
                print(f"   Failed to open file: {e}")
    else:
        print(f"   Error: {app_name} took too long to connect.")

def restore_session():
    # 1. Check Flag
    if not os.path.exists(FLAG_FILE):
        return

    if not os.path.exists(SESSION_FILE):
        return

    with open(SESSION_FILE, "r") as f:
        data = json.load(f)

    # 2. Restore Apps
    for app in data.get("apps", []):
        launch_app_reliable(app)

    # 3. Restore Folders
    for path in data.get("folders", []):
        if os.path.exists(path):
            os.startfile(path)

    # 4. Restore Adobe
    ps_data = data.get("photoshop", {})
    if ps_data.get("was_open") and PHOTOSHOP_PATH:
        restore_adobe_files("Photoshop", "Photoshop.Application", ps_data.get("files", []), PHOTOSHOP_PATH)

    ai_data = data.get("illustrator", {})
    if ai_data.get("was_open") and ILLUSTRATOR_PATH:
        restore_adobe_files("Illustrator", "Illustrator.Application", ai_data.get("files", []), ILLUSTRATOR_PATH)

    # 5. Cleanup
    try:
        os.remove(FLAG_FILE)
    except:
        pass
    
    time.sleep(5)

if __name__ == "__main__":
    restore_session()