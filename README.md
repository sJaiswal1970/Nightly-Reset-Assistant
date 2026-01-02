# 🌙 Nightly Reset Assistant

**Automate your creative workflow.**
Nightly Reset Assistant is a Python-based tool designed for designers and developers on Windows. It safely saves your active Adobe work, snapshots your running applications, and restarts your PC—then automatically restores your entire environment (including specific Photoshop/Illustrator files) the next morning.

## 🚀 Features

* **Adobe Auto-Save:** Connects to Photoshop & Illustrator via COM interface to save all open documents to a timestamped backup folder before closing.
* **Smart State Capture:** Records your open Explorer folders and active applications (Browsers, WhatsApp, VS Code, etc.).
* **Automated Restoration:**
    * Launches apps instantly upon login.
    * Restores folders to their previous locations.
    * **Adobe Deep-Restore:** Launches Photoshop/Illustrator, waits for the splash screen to clear, and re-opens the specific files you were working on.
* **Task Scheduler Integration:** Uses Windows Task Scheduler for high-priority, reliable startup execution.

## 📋 Prerequisites

* **OS:** Windows 10 or Windows 11
* **Language:** Python 3.x
* **Required Libraries:**
    The following libraries are required. The `setup.py` script will attempt to install them automatically, but you can also install them manually:
    * `pywin32` (For Windows & Adobe COM automation)
    * `psutil` (For scanning running processes)
    * `winshell` (For creating desktop shortcuts)
    * `pypiwin32` (Dependency for pywin32)

    **Manual Install Command:**
    ```bash
    pip install pywin32 psutil winshell pypiwin32
    ```

## 🛠️ Installation

You don't need to edit code manually. The included setup script handles the configuration.

1.  **Clone the Repository:**
    ```bash
    git clone [https://github.com/YOUR_USERNAME/Nightly-Reset-Assistant.git](https://github.com/YOUR_USERNAME/Nightly-Reset-Assistant.git)
    cd Nightly-Reset-Assistant
    ```

2.  **Run the Installer:**
    * Right-click `setup.py` and select **Run as Administrator**.
    * *Alternative:* Open a terminal as Admin and run:
        ```bash
        python setup.py
        ```

3.  **Follow the On-Screen Wizard:**
    * The script will ask where you want to store backups.
    * It will auto-detect your Photoshop/Illustrator installation paths.
    * It will create a "Nightly Reset" shortcut on your desktop.
    * It will register the startup task in Windows Task Scheduler.

## 🏃‍♂️ Usage

### 1. The Nightly Shutdown
When you are done for the day, double-click the **Nightly Reset** shortcut on your desktop (or assign it a hotkey like `Ctrl+Alt+R`).

* **What happens:** The script saves your .psd/.ai files, records your session, sets a "restore flag," and restarts your computer.

### 2. The Morning Restore
Simply log in to Windows.

* **What happens:** The Task Scheduler detects the "restore flag" and triggers `restore.py`. It will launch your apps and re-open your Adobe files exactly where you left off.

## ⚙️ Customization

If you need to change your settings later (e.g., you installed a new version of Photoshop), simply edit the `config.json` file generated in the project folder:

```json
{
    "backup_folder": "C:\\Users\\Sahil\\Documents\\Backups",
    "photoshop_path": "C:\\Program Files\\Adobe\\...",
    "apps_to_track": ["chrome.exe", "opera.exe", "code.exe"]
}