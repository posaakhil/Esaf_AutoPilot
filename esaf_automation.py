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
with open("esaf_config.json", 'r', encoding='utf-8') as f:
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

QUEUES = config.get("queues", [])
KEYWORDS = config.get("keywords", [])

downloads_raw = config.get("downloads_folder", "AUTO")
if downloads_raw == "AUTO" or not str(downloads_raw).strip():
    DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads").replace("\\", "/")
else:
    DOWNLOADS_FOLDER = downloads_raw.replace("\\", "/")

BASE_ASSIGNMENT_FOLDER = config.get("base_assignment_folder", "assignment_")
NO_REQUESTS_TEXT = config.get("no_requests_text", "There are no requests available")

# ===== TRACKER & SETUP =====
downloaded_files = []

def check_abort():
    if keyboard.is_pressed('esc'):
        print("\n[EMERGENCY STOP] ESC key pressed!")
        sys.exit(0)

def wait_for_download(keyword_index):
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
                print(f"[SUCCESS] Download detected: {filename}")
                downloaded_files.append(file_path)
                return True

        print("[INFO] Checking for 'no requests' text...")
        pyautogui.click(EXPORT_XLS_COORDS)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)

        try:
            copied_text = pyperclip.paste()
            if NO_REQUESTS_TEXT in copied_text:
                print(f"[INFO] Detected: '{NO_REQUESTS_TEXT}' — skipping export.")
                return True
            else:
                print("[INFO] 'No requests' text not found — continuing wait...")
        except Exception as e:
            print(f"[WARNING] Failed to read clipboard: {e}")

    print("[WARNING] Timeout: No download and no 'no requests' detected.")
    return False

def create_assignment_folder():
    i = 1
    while True:
        folder_name = f"{BASE_ASSIGNMENT_FOLDER}{i}"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            print(f"[INFO] Created folder: {folder_name}")
            return folder_name
        i += 1

def move_downloaded_files(target_folder):
    if not downloaded_files:
        print("[INFO] No .xls files to move.")
        return

    print(f"[INFO] Moving {len(downloaded_files)} .xls files to '{target_folder}'...")
    moved_count = 0
    for file_path in downloaded_files:
        try:
            filename = os.path.basename(file_path)
            dest_path = os.path.join(target_folder, filename)
            shutil.move(file_path, dest_path)
            print(f"   → Moved: {filename}")
            moved_count += 1
        except Exception as e:
            print(f"   [ERROR] Failed to move {filename}: {e}")
    print(f"[SUCCESS] Successfully moved {moved_count} files.")

# ===== SAFETY & START =====
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

print("[STOP] MOVE MOUSE TO TOP-LEFT CORNER TO ABORT.")
print("[STOP] PRESS 'ESC' KEY ANYTIME TO STOP IMMEDIATELY.")
print(f"[INFO] Monitoring Downloads: {DOWNLOADS_FOLDER}")
print("[INFO] Starting ESAF Multi-Phase Automation...")

try:
    print(f"[INFO] Opening: {URL}")
    webbrowser.open(URL)
    time.sleep(WAIT_AFTER_PAGE_LOAD)

    check_abort()
    print(f"[ACTION] Clicking 'Requests Assigned to Me' at {REQUESTS_ASSIGNED_COORDS}")
    pyautogui.click(REQUESTS_ASSIGNED_COORDS)
    time.sleep(WAIT_AFTER_REQUESTS_CLICK)

    # ===== PHASE 1: PROCESS QUEUES =====
    if QUEUES:
        print(f"\n=== PHASE 1: PROCESSING {len(QUEUES)} QUEUES ===")
        for i, queue in enumerate(QUEUES, 1):
            check_abort()
            print(f"\n[PROCESS] Processing queue {i}/{len(QUEUES)}: '{queue}'")

            print(f"[ACTION] Clicking Advanced Filter at {ADVANCED_FILTER_BUTTON}")
            pyautogui.click(ADVANCED_FILTER_BUTTON)
            time.sleep(WAIT_AFTER_ADVANCED_FILTER_OPEN)

            # FIX: Replaced ⚡ with plain text
            print("[ACTION] HYPER-SCROLL: Blasting mouse wheel down for 4 seconds...")
            start_time = time.time()
            while time.time() - start_time < 4:
                pyautogui.scroll(-150)
            time.sleep(1.5)

            print(f"[ACTION] Clicking Status Field at {STATUS_FIELD_COORDS}")
            pyautogui.click(STATUS_FIELD_COORDS)

            print("[ACTION] Clearing field...")
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            print(f"[ACTION] Pasting queue: '{queue}'")
            pyperclip.copy(queue)
            pyautogui.hotkey('ctrl', 'v')
            print("[SUCCESS] Queue pasted.")

            print(f"[ACTION] Clicking Apply_2 Button at {APPLY_2_BUTTON_COORDS}")
            pyautogui.click(APPLY_2_BUTTON_COORDS)
            time.sleep(WAIT_AFTER_APPLY)

            print(f"[ACTION] Clicking Export XLS at {EXPORT_XLS_COORDS}")
            pyautogui.click(EXPORT_XLS_COORDS)

            if not wait_for_download(i):
                print("[WARNING] Proceeding despite timeout.")

            if i < len(QUEUES):
                print(f"[WAIT] Waiting {WAIT_AFTER_EXPORT} seconds before next queue...")
                for sec in range(WAIT_AFTER_EXPORT):
                    check_abort()
                    time.sleep(1)

    # ===== PHASE 2: PROCESS KEYWORDS =====
    if KEYWORDS:
        print(f"\n=== PHASE 2: PROCESSING {len(KEYWORDS)} KEYWORDS ===")
        for i, keyword in enumerate(KEYWORDS, 1):
            check_abort()
            print(f"\n[PROCESS] Processing keyword {i}/{len(KEYWORDS)}: '{keyword}'")

            print(f"[ACTION] Clicking Advanced Filter at {ADVANCED_FILTER_BUTTON}")
            pyautogui.click(ADVANCED_FILTER_BUTTON)
            time.sleep(WAIT_AFTER_ADVANCED_FILTER_OPEN)

            print(f"[ACTION] Clicking Application Name Field at {APPLICATION_NAME_FIELD}")
            pyautogui.click(APPLICATION_NAME_FIELD)

            print("[ACTION] Clearing field with Ctrl+Shift+Right...")
            pyautogui.hotkey('ctrl', 'shift', 'right')
            pyautogui.press('backspace')

            print(f"[ACTION] Pasting: '{keyword}'")
            pyperclip.copy(keyword)
            pyautogui.hotkey('ctrl', 'v')
            print("[SUCCESS] Keyword pasted.")

            print("[ACTION] Clicking outside target to trigger scroll...")
            pyautogui.click(CLICK_OUTSIDE_TARGET)
            time.sleep(0.5)

            # FIX: Replaced ⚡ with plain text
            print("[ACTION] HYPER-SCROLL: Blasting mouse wheel down for 4 seconds...")
            start_time = time.time()
            while time.time() - start_time < 4:
                pyautogui.scroll(-150)
            time.sleep(1.5)

            print(f"[ACTION] Clicking Apply Button at {APPLY_BUTTON}")
            pyautogui.click(APPLY_BUTTON)
            time.sleep(WAIT_AFTER_APPLY)

            print(f"[ACTION] Clicking Export XLS at {EXPORT_XLS_COORDS}")
            pyautogui.click(EXPORT_XLS_COORDS)

            if not wait_for_download(i):
                print("[WARNING] Proceeding despite timeout.")

            if i < len(KEYWORDS):
                print(f"[WAIT] Waiting {WAIT_AFTER_EXPORT} seconds before next keyword...")
                for sec in range(WAIT_AFTER_EXPORT):
                    check_abort()
                    time.sleep(1)

    print("\n[INFO] All phases completed. Preparing to organize files...")
    assignment_folder = create_assignment_folder()
    move_downloaded_files(assignment_folder)

    print(f"\n[SUCCESS] FULL AUTOMATION COMPLETED SUCCESSFULLY!")
    print(f"[INFO] All files moved to: {assignment_folder}")
    print("[SUCCESS] Automation fully completed.")

except Exception as e:
    print(f"[ERROR] {e}")
    print("[FATAL] Script stopped due to error.")

except KeyboardInterrupt:
    print("\n[INFO] Script manually aborted by user (Ctrl+C).")