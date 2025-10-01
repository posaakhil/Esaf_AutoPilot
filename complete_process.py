import pyautogui
import time
import webbrowser
import pyperclip
import keyboard
import sys
import os
import glob
import shutil
import subprocess
import json

# ===== LOAD CONFIG FROM JSON =====
with open("esaf_config.json", 'r') as f:
    config = json.load(f)

URL = config["url"]
REQUESTS_ASSIGNED_COORDS = tuple(config["mouse_coords"]["requests_assigned"])
ADVANCED_FILTER_BUTTON = tuple(config["mouse_coords"]["advanced_filter"])
APPLICATION_NAME_FIELD = tuple(config["mouse_coords"]["app_field"])
CLICK_OUTSIDE_TARGET = tuple(config["mouse_coords"]["click_outside"])
APPLY_BUTTON = tuple(config["mouse_coords"]["apply_button"])
EXPORT_XLS_COORDS = tuple(config["mouse_coords"]["export_xls"])
STATUS_FIELD_COORDS = tuple(config["mouse_coords"]["status_field"])
APPLY_2_BUTTON_COORDS = tuple(config["mouse_coords"]["apply_2_button"])

WAIT_AFTER_PAGE_LOAD = config["timings"]["wait_after_page_load"]
WAIT_AFTER_REQUESTS_CLICK = config["timings"]["wait_after_requests_click"]
WAIT_AFTER_ADVANCED_FILTER_OPEN = config["timings"]["wait_after_advanced_filter_open"]
WAIT_AFTER_PASTE = config["timings"]["wait_after_paste"]
WAIT_AFTER_APPLY = config["timings"]["wait_after_apply"]
WAIT_AFTER_EXPORT = config["timings"]["wait_after_export"]
MAX_WAIT_FOR_DOWNLOAD = config["timings"]["max_wait_for_download"]
CHECK_INTERVAL = config["timings"]["check_interval"]

ASSIGNEES = config["assignees"]
KEEP_COLUMNS = config["keep_columns"]
OUTPUT_FILE = config["output_file"]

QUEUES = config.get("queues", [])
KEYWORDS = config.get("keywords", [])

# ✅ New (dynamic + fallback)
downloads_raw = config.get("downloads_folder", "AUTO")
if downloads_raw == "AUTO" or not downloads_raw.strip():
    DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads").replace("\\", "/")
else:
    DOWNLOADS_FOLDER = downloads_raw.replace("\\", "/")

BASE_ASSIGNMENT_FOLDER = config.get("base_assignment_folder", "assignment_")
NO_REQUESTS_TEXT = config.get("no_requests_text", "There are no requests available")


# ===== TRACKER & SETUP =====
downloaded_files = []

def check_abort():
    """Check if user pressed ESC or closed terminal"""
    if keyboard.is_pressed('esc'):
        print("\n�� EMERGENCY STOP: ESC key pressed!")
        sys.exit(0)

def wait_for_download(keyword_index):
    """Wait for .xls OR detect 'no requests' by selecting all text"""
    start_time = time.time()
    initial_files = set(glob.glob(os.path.join(DOWNLOADS_FOLDER, "*.xls")))

    while time.time() - start_time < MAX_WAIT_FOR_DOWNLOAD:
        check_abort()
        time.sleep(CHECK_INTERVAL)

        current_files = set(glob.glob(os.path.join(DOWNLOADS_FOLDER, "*.xls")))
        new_files = current_files - initial_files

        for file_path in new_files:
            filename = os.path.basename(file_path)
            if any(filename[i:i+5].isdigit() for i in range(len(filename) - 4)):
                print(f"✅ Download detected: {filename}")
                downloaded_files.append(file_path)
                return True

        print("�� Checking for 'no requests' text...")
        pyautogui.click(EXPORT_XLS_COORDS)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)

        try:
            copied_text = pyperclip.paste()
            if NO_REQUESTS_TEXT in copied_text:
                print(f"ℹ️ Detected: '{NO_REQUESTS_TEXT}' — skipping export.")
                return True
            else:
                print("�� 'No requests' text not found — continuing wait...")
        except Exception as e:
            print(f"⚠️ Failed to read clipboard: {e}")

    print("⚠️ Timeout: No download and no 'no requests' detected.")
    return False

def create_assignment_folder():
    """Create assignment_1, assignment_2, ... folder in current directory"""
    i = 1
    while True:
        folder_name = f"{BASE_ASSIGNMENT_FOLDER}{i}"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            print(f"�� Created folder: {folder_name}")
            return folder_name
        i += 1

def move_downloaded_files(target_folder):
    """Move all tracked .xls files to target folder"""
    if not downloaded_files:
        print("�� No .xls files to move.")
        return

    print(f"�� Moving {len(downloaded_files)} .xls files to '{target_folder}'...")
    moved_count = 0
    for file_path in downloaded_files:
        try:
            filename = os.path.basename(file_path)
            dest_path = os.path.join(target_folder, filename)
            shutil.move(file_path, dest_path)
            print(f"   → Moved: {filename}")
            moved_count += 1
        except Exception as e:
            print(f"   ❌ Failed to move {filename}: {e}")
    print(f"✅ Successfully moved {moved_count} files.")

# ===== SAFETY & START =====
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

print("✋ MOVE MOUSE TO TOP-LEFT CORNER TO ABORT.")
print("✋ PRESS 'ESC' KEY ANYTIME TO STOP IMMEDIATELY.")
print(f"�� Monitoring Downloads: {DOWNLOADS_FOLDER}")
print("�� Starting ESAF Multi-Phase Automation...")

try:
    # ===== STEP 1: OPEN PAGE & CLICK 'REQUESTS ASSIGNED TO ME' ONCE =====
    print(f"�� Opening: {URL}")
    webbrowser.open(URL)
    time.sleep(WAIT_AFTER_PAGE_LOAD)

    check_abort()
    print(f"��️ Clicking 'Requests Assigned to Me' at {REQUESTS_ASSIGNED_COORDS}")
    pyautogui.click(REQUESTS_ASSIGNED_COORDS)
    time.sleep(WAIT_AFTER_REQUESTS_CLICK)

    # ===== PHASE 1: PROCESS QUEUES =====
    if QUEUES:
        print(f"\n=== PHASE 1: PROCESSING {len(QUEUES)} QUEUES ===")
        for i, queue in enumerate(QUEUES, 1):
            check_abort()
            print(f"\n�� Processing queue {i}/{len(QUEUES)}: '{queue}'")

            pyautogui.click(ADVANCED_FILTER_BUTTON)
            time.sleep(WAIT_AFTER_ADVANCED_FILTER_OPEN)

            print("⚡ HYPER-SCROLL: Blasting mouse wheel down for 4 seconds...")
            start_time = time.time()
            while time.time() - start_time < 4:
                pyautogui.scroll(-150)
            time.sleep(1.5)

            pyautogui.click(STATUS_FIELD_COORDS)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyperclip.copy(queue)
            pyautogui.hotkey('ctrl', 'v')
            print("✅ Queue pasted.")

            pyautogui.click(APPLY_2_BUTTON_COORDS)
            time.sleep(WAIT_AFTER_APPLY)

            pyautogui.click(EXPORT_XLS_COORDS)

            if not wait_for_download(i):
                print("⚠️ Proceeding despite timeout.")

            if i < len(QUEUES):
                print(f"⏳ Waiting {WAIT_AFTER_EXPORT} seconds before next queue...")
                for sec in range(WAIT_AFTER_EXPORT):
                    check_abort()
                    time.sleep(1)

    # ===== PHASE 2: PROCESS KEYWORDS =====
    if KEYWORDS:
        print(f"\n=== PHASE 2: PROCESSING {len(KEYWORDS)} KEYWORDS ===")
        for i, keyword in enumerate(KEYWORDS, 1):
            check_abort()
            print(f"\n�� Processing keyword {i}/{len(KEYWORDS)}: '{keyword}'")

            pyautogui.click(ADVANCED_FILTER_BUTTON)
            time.sleep(WAIT_AFTER_ADVANCED_FILTER_OPEN)

            pyautogui.click(APPLICATION_NAME_FIELD)
            pyautogui.hotkey('ctrl', 'shift', 'right')
            pyautogui.press('backspace')
            pyperclip.copy(keyword)
            pyautogui.hotkey('ctrl', 'v')
            print("✅ Keyword pasted.")

            pyautogui.click(CLICK_OUTSIDE_TARGET)
            time.sleep(0.5)

            print("⚡ HYPER-SCROLL: Blasting mouse wheel down for 4 seconds...")
            start_time = time.time()
            while time.time() - start_time < 4:
                pyautogui.scroll(-150)
            time.sleep(1.5)

            pyautogui.click(APPLY_BUTTON)
            time.sleep(WAIT_AFTER_APPLY)

            pyautogui.click(EXPORT_XLS_COORDS)

            if not wait_for_download(i):
                print("⚠️ Proceeding despite timeout.")

            if i < len(KEYWORDS):
                print(f"⏳ Waiting {WAIT_AFTER_EXPORT} seconds before next keyword...")
                for sec in range(WAIT_AFTER_EXPORT):
                    check_abort()
                    time.sleep(1)

    # ===== FINAL STEP: CREATE FOLDER & MOVE FILES =====
    print("\n�� All phases completed. Preparing to organize files...")
    assignment_folder = create_assignment_folder()
    move_downloaded_files(assignment_folder)

    print(f"\n�� FULL AUTOMATION COMPLETED SUCCESSFULLY! ��")
    print(f"�� All files moved to: {assignment_folder}")
    print("✅ Automation fully completed.")

    # ===== AUTO-TRIGGER STEP 2: MERGE & CLEANUP =====
    print("\n�� Auto-triggering merge_and_cleanup.py...")
    try:
        subprocess.run([sys.executable, "merge_and_cleanup.py", "--auto"], check=True)
        print("✅ Step 2 completed automatically.")
    except Exception as e:
        print(f"❌ Failed to auto-trigger merge_and_cleanup.py: {e}")

    # ===== AUTO-TRIGGER STEP 3: DATA ANALYSIS SPLIT =====
    print("\n�� Auto-triggering Data_Analysis_Split.py...")
    try:
        subprocess.run([sys.executable, "Data_Analysis_Split.py", "--auto"], check=True)
        print("✅ Step 3 completed automatically.")
    except Exception as e:
        print(f"❌ Failed to auto-trigger Data_Analysis_Split.py: {e}")

    # ===== AUTO-TRIGGER STEP 4: SUMMARY PIVOT =====
    print("\n�� Auto-triggering summary_pivot.py...")
    try:
        subprocess.run([sys.executable, "summary_pivot.py", "--auto"], check=True)
        print("✅ Step 4 completed automatically.")
    except Exception as e:
        print(f"❌ Failed to auto-trigger summary_pivot.py: {e}")

    # ===== AUTO-TRIGGER STEP 5: INTERACTIVE DASHBOARD =====
    print("\n�� Auto-triggering Interactive_Dashboard.py...")
    try:
        subprocess.run([sys.executable, "Interactive_Dashboard.py"], check=True)
        print("✅ Step 5 completed automatically.")
    except Exception as e:
        print(f"❌ Failed to auto-trigger Interactive_Dashboard.py: {e}")

    print("\n�� ENTIRE WORKFLOW COMPLETED END-TO-END!")

except Exception as e:
    print(f"❌ ERROR: {e}")
    print("�� Script stopped due to error.")

except KeyboardInterrupt:
    print("\n�� Script manually aborted by user (Ctrl+C).")