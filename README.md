\# 🌙 Nightly Reset Assistant



A Python-based automation tool for Windows creatives. It automatically saves your work, captures your session state, and restarts your PC overnight—then restores everything exactly as you left it (including specific Adobe files) when you log in the next morning.



\## 🚀 Features



\* \*\*Adobe Auto-Save:\*\* Automatically connects to Photoshop \& Illustrator via COM interface to save open documents to a timestamped backup folder.

\* \*\*Smart Shutdown:\*\* Records running apps (Browsers, WhatsApp, etc.) and open Explorer folders before restarting Windows.

\* \*\*Automated Restoration:\*\*

&nbsp;   \* Re-opens exact folder paths.

&nbsp;   \* Launches specific apps (including Opera GX, WhatsApp).

&nbsp;   \* \*\*Adobe Recovery:\*\* Launches Photoshop/Illustrator, waits for initialization, and re-opens the specific files you were working on.

\* \*\*Startup Flag:\*\* Uses a "flag file" system so restoration only happens after a planned nightly reset, not after manual restarts.



\## 🛠️ Prerequisites



\* Windows 10/11

\* Python 3.x

\* \*\*Libraries:\*\* `pip install pywin32 psutil`



\## ⚙️ Configuration



1\.  \*\*Clone the repo:\*\*

&nbsp;   ```bash

&nbsp;   git clone \[https://github.com/yourusername/Nightly-Reset-Assistant.git](https://github.com/yourusername/Nightly-Reset-Assistant.git)

&nbsp;   ```

2\.  \*\*Setup Scripts:\*\*

&nbsp;   \* Rename `reset\_template.py` to `reset.py`.

&nbsp;   \* Rename `restore\_template.py` to `restore.py`.

3\.  \*\*Edit Paths:\*\* Open both files and update the `CONFIGURATION` section with your specific paths:

&nbsp;   \* `BACKUP\_FOLDER`: Where you want your `.psd/.ai` backups stored.

&nbsp;   \* `PHOTOSHOP\_PATH` / `ILLUSTRATOR\_PATH`: The full path to your `.exe` files.



\## 🏃‍♂️ Usage



\### 1. The Nightly Reset

Run `reset.py` when you are done for the day. You can set up a desktop shortcut or global hotkey (e.g., `Ctrl+Alt+R`).

\* It saves all Adobe work.

\* It records your open folders.

\* It restarts your PC.



\### 2. The Morning Restore

Set up `restore.py` to run automatically:

1\.  Open \*\*Task Scheduler\*\*.

2\.  Create a new task triggered \*\*"At log on"\*\*.

3\.  Action: Start Program -> `py` -> Arguments: `"path\\to\\restore.py"`.

4\.  Ensure "Run with highest privileges" is checked.



Now, when you log in after a nightly reset, your environment will reconstruct itself automatically.

