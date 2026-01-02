import os
import time
import json
import subprocess
import win32com.client

# ==============================================================================
# ⚙️ USER CONFIGURATION
# ==============================================================================
# 1. Must match the paths used in reset.py
SESSION_FILE = r"C:\Path\To\Script\session.json"
FLAG_FILE = r"C:\Path\To\Script\restore_pending.flag"

# 2. EXACT PATHS to your applications
#    Windows Task Scheduler often needs full paths to find apps reliably.
#    Right-click your App Shortcut -> Properties -> Target to find these.
OPERA_PATH = r"C:\Users\YourName\AppData\Local\Programs\Opera GX\opera.exe"
PHOTOSHOP_PATH = r"C:\Program Files\Adobe\Adobe Photoshop 2026\Photoshop.exe"
ILLUSTRATOR_PATH = r"C:\Program Files\Adobe\Adobe Illustrator 2026\Support Files\Contents\Windows\Illustrator.exe"
# ==============================================================================

def launch_app_reliable(exe_name):
    """Launches apps, with special handling for WhatsApp and Opera."""
    try:
        if "whatsapp" in exe_name.lower():
             # WhatsApp often needs the URI scheme to launch correctly
             os.system("start whatsapp:")
        
        elif "opera" in exe_name.lower():
            # Use the specific path for Opera if standard launch fails
            subprocess.Popen(f'"{OPERA_PATH}"', shell=True)
            
        else:
             # Standard launch for Chrome, Edge, etc.
             subprocess.Popen(f"start {exe_name}", shell=True)
    except Exception as e:
        print(f"Failed to launch {exe_name}: {e}")

def restore_adobe_files(app_name, program_id, file_list, exe_path):
    """
    Launches an Adobe app via its full path, waits for the splash screen to finish,
    and then opens the specified files.
    """
    print(f"Restoring {app_name}...")
    
    # 1. LAUNCH from the specific path
    try:
        # Use quotes to handle spaces in path names
        subprocess.Popen(f'"{exe_path}"', shell=True)
    except Exception as e:
        print(f"   Error launching {app_name}: {e}")
        return

    if not file_list:
        return
    
    # 2. WAIT LOOP
    # We wait up to 60 seconds for the app to initialize the COM interface
    print(f"   Waiting for {app_name} to respond...")
    app = None
    for attempt in range(60): 
        try:
            app = win32com.client.GetActiveObject(program_id)
            # Try accessing a property to ensure it's truly ready
            _ = app.Name 
            break
        except:
            time.sleep(1)
    
    if app:
        print(f"   Connected! Waiting 5s for splash screen to clear...")
        # 3. SAFETY BUFFER
        # Essential for preventing "Open" commands from failing during startup
        time.sleep(5)
        
        app.Visible = True
        app.UserInteractionLevel = 1 # 1 = DISPLAYALERTS
        
        # 4. OPEN FILES
        for file_path in file_list:
            try:
                print(f"   Opening: {os.path.basename(file_path)}")
                app.Open(file_path)
                # Small pause prevents command flooding
                time.sleep(1) 
            except Exception as e:
                print(f"   Failed to open file: {e}")
    else:
        print(f"   Error: {app_name} took too long to connect.")

def restore_session():
    # 1. CHECK THE FLAG
    # If the flag file is missing, this is a manual restart -> Do nothing.
    if not os.path.exists(FLAG_FILE):
        return

    if not os.path.exists(SESSION_FILE):
        return

    with open(SESSION_FILE, "r") as f:
        data = json.load(f)

    # 2. RESTORE APPS
    for app in data.get("apps", []):
        launch_app_reliable(app)

    # 3. RESTORE FOLDERS
    for path in data.get("folders", []):
        if os.path.exists(path):
            os.startfile(path)

    # 4. RESTORE ADOBE
    ps_data = data.get("photoshop", {})
    if ps_data.get("was_open"):
        restore_adobe_files("Photoshop", "Photoshop.Application", ps_data.get("files", []), PHOTOSHOP_PATH)

    ai_data = data.get("illustrator", {})
    if ai_data.get("was_open"):
        restore_adobe_files("Illustrator", "Illustrator.Application", ai_data.get("files", []), ILLUSTRATOR_PATH)

    # 5. CLEANUP
    # Remove the flag so the script doesn't run on the next manual restart
    try:
        os.remove(FLAG_FILE)
    except:
        pass
    
    # Keep window open briefly for confirmation
    time.sleep(5)

if __name__ == "__main__":
    restore_session()