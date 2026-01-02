import os
import sys
import json
import subprocess
import shutil
import winreg # Built-in module to find installed apps

def is_admin():
    try:
        return os.getuid() == 0
    except AttributeError:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0

def install_dependencies():
    print("📦 Installing required libraries...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32", "psutil", "winshell", "pypiwin32"])

def find_adobe_app(app_name, year_range=range(2026, 2020, -1)):
    """Auto-detects Adobe paths by checking standard folders."""
    # Check default drives
    drives = ["C:\\", "D:\\", "X:\\"]
    roots = ["Program Files", "Program Files (x86)"]
    
    print(f"🔍 Searching for {app_name}...")
    
    for drive in drives:
        for root in roots:
            base = os.path.join(drive, root, "Adobe")
            if not os.path.exists(base):
                continue
                
            # Look for folder versions (e.g., "Adobe Photoshop 2024")
            for folder in os.listdir(base):
                if app_name.lower() in folder.lower():
                    # Look for the .exe inside
                    full_path = os.path.join(base, folder, f"{app_name}.exe")
                    
                    # Illustrator hides deeper
                    if "Illustrator" in app_name:
                        full_path = os.path.join(base, folder, "Support Files", "Contents", "Windows", "Illustrator.exe")
                    
                    if os.path.exists(full_path):
                        print(f"   ✅ Found: {full_path}")
                        return full_path
    return None

def create_task_scheduler(script_path):
    print("⏰ Creating 'Nightly Restore' task in Windows Scheduler...")
    
    # We use the pythonw.exe (windowless) so no black box appears on startup
    python_exe = sys.executable.replace("python.exe", "pythonw.exe")
    
    cmd = (
        f'schtasks /Create /F '
        f'/TN "Nightly Restore" '
        f'/TR "\'{python_exe}\' \'{script_path}\'" '
        f'/SC ONLOGON '
        f'/RL HIGHEST'
    )
    
    # Execute command
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    if "SUCCESS" in result.stdout:
        print("   ✅ Task Scheduler configured successfully.")
    else:
        print(f"   ❌ Failed to create task: {result.stderr}")

def create_desktop_shortcut(target_script):
    try:
        import winshell
        from win32com.client import Dispatch
        
        desktop = winshell.desktop()
        path = os.path.join(desktop, "Nightly Reset.lnk")
        
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = sys.executable
        shortcut.Arguments = f'"{target_script}"'
        shortcut.IconLocation = "shell32.dll, 27" # Uses a cool 'Shutdown' icon
        shortcut.save()
        print("   ✅ Desktop shortcut created.")
    except Exception as e:
        print(f"   ⚠️ Could not create shortcut: {e}")

def main():
    print("=== NIGHTLY RESET ASSISTANT SETUP ===\n")
    
    if not is_admin():
        print("❌ Error: Please run this script as Administrator!")
        print("   (Right-click Command Prompt -> Run as Administrator)")
        input("Press Enter to exit...")
        return

    # 1. Install Requirements
    try:
        install_dependencies()
    except:
        print("⚠️ Warning: Could not install dependencies automatically.")

    # 2. Configuration Wizard
    config = {}
    
    # Backup Folder
    default_backup = os.path.join(os.path.expanduser("~"), "Documents", "NightlyBackups")
    print(f"\n📂 Where should backups go? [Default: {default_backup}]")
    user_backup = input("   Path: ").strip()
    config["backup_folder"] = user_backup if user_backup else default_backup
    
    if not os.path.exists(config["backup_folder"]):
        os.makedirs(config["backup_folder"])

    # Auto-Detect Adobe
    config["photoshop_path"] = find_adobe_app("Photoshop") or input("   ❌ Photoshop not found. Paste full path: ").strip('"')
    config["illustrator_path"] = find_adobe_app("Illustrator") or input("   ❌ Illustrator not found. Paste full path: ").strip('"')
    
    # Browser / Opera
    print("\n🌐 Paste your Browser Path (Chrome/Opera/Edge):")
    config["opera_path"] = input("   Path: ").strip('"')

    # Apps list
    config["apps_to_track"] = ["chrome.exe", "opera.exe", "WhatsApp.Root.exe", "code.exe", "discord.exe"]

    # 3. Save Config
    with open("config.json", "w") as f:
        json.dump(config, f, indent=4)
    print("\n💾 Configuration saved.")

    # 4. Setup Automation
    current_dir = os.path.dirname(os.path.abspath(__file__))
    restore_script = os.path.join(current_dir, "restore.py")
    reset_script = os.path.join(current_dir, "reset.py")

    create_task_scheduler(restore_script)
    create_desktop_shortcut(reset_script)

    print("\n✨ INSTALLATION COMPLETE! ✨")
    print("1. Use the 'Nightly Reset' shortcut on your desktop tonight.")
    print("2. When you login tomorrow, everything will restore automatically.")
    input("\nPress Enter to exit.")

if __name__ == "__main__":
    main()