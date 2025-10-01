import subprocess
import sys
import os
import json
import shutil
import keyboard
from colorama import Fore, Style, init
import pyautogui  # Added for mouse coordinate capture

# Initialize colorama
init(autoreset=True)

CONFIG_FILE = "esaf_config.json"
DEFAULTS_FILE = "esaf_config_defaults.json"
BACKUP_FILE = "esaf_config_backup.json"

# ===== SCRIPT FILENAMES (MUST MATCH YOUR ACTUAL FILE NAMES) =====
SCRIPTS = {
    "step1": "esaf_automation.py",
    "step2": "merge_and_cleanup.py",
    "step3": "Data_Analysis_Split.py",
    "step4": "summary_pivot.py",
    "step5": "Interactive_Dashboard.py"
}

# ===== ENHANCED PROFESSIONAL ASCII ART (UNCHANGED) =====
ASCII_ART = f"""
{Fore.CYAN}╔══════════════════════════════════════════════════════════════════════════════════╗
{Fore.CYAN}║                                                                                  ║
{Fore.CYAN}║  {Fore.BLUE}███████╗███████╗ █████╗ ███████╗     █████╗ ██╗   ██╗████████╗ ██████╗ {Fore.CYAN}  ║
{Fore.CYAN}║  {Fore.BLUE}██╔════╝██╔════╝██╔══██╗██╔════╝    ██╔══██╗██║   ██║╚══██╔══╝██╔═══██╗{Fore.CYAN}  ║
{Fore.CYAN}║  {Fore.BLUE}█████╗  ███████╗███████║█████╗      ███████║██║   ██║   ██║   ██║   ██║{Fore.CYAN}  ║
{Fore.CYAN}║  {Fore.BLUE}██╔══╝  ╚════██║██╔══██║██╔══╝      ██╔══██║██║   ██║   ██║   ██║   ██║{Fore.CYAN}  ║
{Fore.CYAN}║  {Fore.BLUE}███████╗███████║██║  ██║██          ██║  ██║╚██████╔╝   ██║   ╚██████╔╝{Fore.CYAN}  ║
{Fore.CYAN}║  {Fore.BLUE}╚══════╝╚══════╝╚═╝  ╚═╝╚═          ╚═╝  ╚═╝ ╚═════╝    ╚═╝    ╚═════╝ {Fore.CYAN}  ║
{Fore.CYAN}║                                                                                  ║
{Fore.CYAN}║  {Fore.GREEN}╔══════════════════════════════════════════════════════════════════════╗  {Fore.CYAN}║
{Fore.CYAN}║  {Fore.GREEN}║                   AUTOMATED REQUEST MANAGEMENT SYSTEM                ║  {Fore.CYAN}║
{Fore.CYAN}║  {Fore.GREEN}╚══════════════════════════════════════════════════════════════════════╝  {Fore.CYAN}║
{Fore.CYAN}║                                                                                  ║
{Fore.CYAN}╚══════════════════════════════════════════════════════════════════════════════════╝
"""

def check_abort():
    """Check if user pressed ESC — can be called anywhere"""
    if keyboard.is_pressed('esc'):
        print(f"\n{Fore.RED}🛑 EMERGENCY STOP: ESC key pressed!")
        raise KeyboardInterrupt("User pressed ESC")

def ensure_defaults():
    """Create defaults file if it doesn't exist"""
    if not os.path.exists(DEFAULTS_FILE):
        defaults = {
            "url": "https://esaf.hca.corpad.net/",
            "downloads_folder": "AUTO",
            "mouse_coords": {
                "requests_assigned": [168, 212],
                "advanced_filter": [728, 317],
                "app_field": [589, 516],
                "click_outside": [534, 512],
                "apply_button": [1357, 780],
                "export_xls": [1834, 317],
                "status_field": [588, 674],
                "apply_2_button": [1358, 781]
            },
            "timings": {
                "wait_after_page_load": 6,
                "wait_after_requests_click": 3,
                "wait_after_advanced_filter_open": 2,
                "wait_after_paste": 2,
                "wait_after_apply": 5,
                "wait_after_export": 3,
                "max_wait_for_download": 30,
                "check_interval": 1
            },
            "queues": ["CORP-ACCESS-INDIA"],
            "keywords": ["SMART", "SMART Inventory", "Smart admin", "abc", "HOST"],
            "assignees": ["Akhil", "Swathi", "Divya", "Amreen", "Riya", "Madhurima", "Vamsitha"],
            "rules": {
                "india_overflow_threshold": 70,
                "max_per_person_domestic": 15,
                "max_per_person_meditech": 5,
                "requests_per_app_domestic": 2,
                "india_keyword": "CORP-ACCESS-INDIA",
                "meditech_keyword": "MEDITECH_Expanse_CAP_Panhandle_Market"
            },
            "output_file": "Today_Assignment.xlsx",
            "keep_columns": [
                "User name",
                "sAMAccountName",
                "Request date",
                "Last updated time",
                "Requested by",
                "Request",
                "Status",
                "Application"
            ],
            "excel_options": {
                "auto_fit_column_width": True,
                "wrap_text": True
            }
        }
        with open(DEFAULTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(defaults, f, indent=4, ensure_ascii=False)
        print(f"{Fore.GREEN}✅ Created default config: {DEFAULTS_FILE}")

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return None
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading config: {e}")
        return None

def resolve_downloads_folder(config):
    raw = config.get("downloads_folder", "AUTO")
    if raw == "AUTO" or not str(raw).strip():
        return os.path.join(os.path.expanduser("~"), "Downloads").replace("\\", "/")
    return str(raw).replace("\\", "/")

def save_config(config):
    if os.path.exists(CONFIG_FILE):
        shutil.copy2(CONFIG_FILE, BACKUP_FILE)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)
    print(f"{Fore.GREEN}💾 Saved config to {CONFIG_FILE}")

def reset_to_defaults():
    if os.path.exists(DEFAULTS_FILE):
        with open(DEFAULTS_FILE, 'r', encoding='utf-8') as f:
            defaults = json.load(f)
        save_config(defaults)
        print(f"{Fore.GREEN}✅ Config reset to defaults!")
    else:
        print(f"{Fore.YELLOW}⚠️ Defaults file not found. Creating...")
        ensure_defaults()
        reset_to_defaults()

def validate_config():
    config = load_config()
    if not config:
        return False
    required_fields = ['url', 'downloads_folder', 'mouse_coords', 'keywords', 'queues', 'assignees', 'rules']
    for field in required_fields:
        if field not in config:
            print(f"{Fore.RED}❌ Missing required field: {field}")
            return False
    mouse_coords = config.get('mouse_coords', {})
    required_coords = ['requests_assigned', 'advanced_filter', 'app_field', 'status_field',
                      'apply_button', 'apply_2_button', 'export_xls', 'click_outside']
    for coord in required_coords:
        if coord not in mouse_coords or not isinstance(mouse_coords[coord], list) or len(mouse_coords[coord]) != 2:
            print(f"{Fore.RED}❌ Invalid mouse coord: {coord}")
            return False
    return True

def edit_url(config):
    current = config.get("url", "https://esaf.hca.corpad.net/")
    print(f"\n{Fore.CYAN}🌐 Current URL: {Fore.WHITE}{current}")
    new_url = input(f"{Fore.CYAN}✏️ Enter new ESAF URL (or press Enter to keep): ").strip()
    final_url = new_url if new_url else current
    config["url"] = final_url
    save_config(config)
    print(f"{Fore.GREEN}✅ URL updated to: {Fore.WHITE}{final_url}")

def capture_mouse_coordinates(config):
    print(f"\n{Fore.CYAN}🖱️  MOUSE COORDINATE CAPTURE TOOL")
    print(f"{Fore.YELLOW}❗ Set browser zoom to 100% and keep window position fixed!")
    if "mouse_coords" not in config:
        config["mouse_coords"] = {}
    steps = [
        ("Requests Assigned to Me", "requests_assigned"),
        ("Advanced Filter Button", "advanced_filter"),
        ("Application Name Field", "app_field"),
        ("Apply Button", "apply_button"),
        ("Export XLS Button", "export_xls"),
        ("Click Outside Target", "click_outside"),
        ("Status Field", "status_field"),
        ("Apply_2 Button", "apply_2_button")
    ]
    for display_name, key in steps:
        input(f"\n{Fore.CYAN}👉 Hover over '{display_name}', then press ENTER...")
        x, y = pyautogui.position()
        config["mouse_coords"][key] = [x, y]
        print(f"{Fore.GREEN}✅ Captured: {Fore.WHITE}{key} → ({x}, {y})")
    save_config(config)

def edit_list(config, key, name):
    current = config.get(key, [])
    print(f"\n{Fore.CYAN}📋 Current {name}: {Fore.WHITE}{current}")
    print(f"{Fore.YELLOW}1) ➕ Add item(s)")
    print(f"{Fore.YELLOW}2) ➖ Remove item(s)")
    print(f"{Fore.YELLOW}3) 🗑️ Clear all")
    print(f"{Fore.YELLOW}4) ◀️ Back")
    choice = input(f"{Fore.CYAN}Choose option (1-4): ").strip()
    if choice == "1":
        try:
            count = int(input(f"{Fore.CYAN}🔢 How many {name.lower()} to add? ").strip())
            if count <= 0:
                print(f"{Fore.YELLOW}⚠️ Skipping.")
                return
            new_items = []
            for i in range(count):
                item = input(f"{Fore.CYAN}✏️ Enter {name[:-1].lower()} #{i+1}: ").strip()
                if item:
                    new_items.append(item)
            if new_items:
                current.extend(new_items)
                config[key] = current
                save_config(config)
                print(f"{Fore.GREEN}✅ Added {len(new_items)} {name.lower()}.")
        except ValueError:
            print(f"{Fore.RED}❌ Invalid number.")
    elif choice == "2" and current:
        print(f"\n{Fore.CYAN}🗑️ Select items to remove (comma-separated numbers):")
        for i, item in enumerate(current, 1):
            print(f"{Fore.YELLOW}{i}) {item}")
        indices_input = input(f"{Fore.CYAN}🔢 Enter numbers (e.g., 1,3,5) or 0 to cancel: ").strip()
        if indices_input == "0":
            return
        try:
            indices = [int(x.strip()) - 1 for x in indices_input.split(",") if x.strip().isdigit()]
            indices = sorted(set(indices), reverse=True)
            removed = []
            for idx in indices:
                if 0 <= idx < len(current):
                    removed.append(current.pop(idx))
            if removed:
                config[key] = current
                save_config(config)
                print(f"{Fore.GREEN}✅ Removed: {removed}")
            else:
                print(f"{Fore.YELLOW}⚠️ No valid items removed.")
        except Exception as e:
            print(f"{Fore.RED}❌ Invalid input.")
    elif choice == "3":
        config[key] = []
        save_config(config)
        print(f"{Fore.GREEN}✅ Cleared all {name}")

def edit_rules(config):
    rules = config.get("rules", {})
    print(f"\n{Fore.CYAN}⚙️ Current Rules:")
    for key, value in rules.items():
        print(f"{Fore.YELLOW}• {key}: {value}")
    fields = {"1": "india_overflow_threshold", "2": "max_per_person_domestic", "3": "max_per_person_meditech", "4": "requests_per_app_domestic"}
    print(f"\n{Fore.YELLOW}1) 📊 Edit India Overflow Threshold")
    print(f"{Fore.YELLOW}2) 👥 Edit Max Per Person Domestic")
    print(f"{Fore.YELLOW}3) 🏥 Edit Max Per Person Meditech")
    print(f"{Fore.YELLOW}4) 📦 Edit Requests Per App Domestic")
    print(f"{Fore.YELLOW}5) ◀️ Back")
    choice = input(f"{Fore.CYAN}Choose option (1-5): ").strip()
    if choice in fields:
        try:
            value = int(input(f"{Fore.CYAN}✏️ Enter new value for {fields[choice]}: ").strip())
            rules[fields[choice]] = value
            config["rules"] = rules
            save_config(config)
            print(f"{Fore.GREEN}✅ Rule updated!")
        except ValueError:
            print(f"{Fore.RED}❌ Invalid number.")

def view_config(config):
    print(f"\n{Fore.CYAN}📄 CURRENT CONFIGURATION:")
    print(json.dumps(config, indent=2, ensure_ascii=False))

def run_script(script_name, step_name):
    """Run script with ESC monitoring and live output"""
    if not os.path.exists(script_name):
        print(f"{Fore.RED}❌ {script_name} not found!")
        return False

    print(f"\n{Fore.CYAN}🚀 Running {step_name}...")
    print(f"{Fore.YELLOW}🛑 Press ESC anytime to ABORT!")
    try:
        proc = subprocess.Popen(
            [sys.executable, script_name],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True
        )

        while True:
            check_abort()
            output = proc.stdout.readline()
            if output == '' and proc.poll() is not None:
                break
            if output:
                print(output.strip())

        rc = proc.poll()
        if rc == 0:
            print(f"\n{Fore.GREEN}✅ {step_name} completed successfully!")
            return True
        else:
            print(f"\n{Fore.RED}❌ {step_name} failed with return code {rc}")
            return False

    except KeyboardInterrupt:
        print(f"\n{Fore.RED}🛑 Automation interrupted by user (ESC or Ctrl+C).")
        if 'proc' in locals():
            proc.terminate()
            proc.wait()
        return False
    except Exception as e:
        print(f"\n{Fore.RED}❌ Unexpected error: {e}")
        return False

def show_menu():
    """Menu with MAXIMUM TEXT VISIBILITY (simulated 'big text')"""
    try:
        terminal_width = shutil.get_terminal_size().columns
    except:
        terminal_width = 80

    # Clear screen
    os.system('cls' if os.name == 'nt' else 'clear')

    # Print ASCII banner (unchanged)
    print(ASCII_ART)
    print()

    # System status
    config_status = "✅ CONFIGURED" if validate_config() else "❌ NOT CONFIGURED"
    status_lines = [
        f"{Fore.YELLOW}┌────────────────────────────────────────────────────────────────┐",
        f"{Fore.YELLOW}│                    {Fore.WHITE}SYSTEM STATUS: {config_status} {Fore.YELLOW}                   │",
        f"{Fore.YELLOW}└────────────────────────────────────────────────────────────────┘"
    ]
    for line in status_lines:
        print(line.center(terminal_width))
    print()

    # Notices
    notices = [
        f"{Fore.CYAN}💡 First-time user? Configure (Option 1) before running full process!",
        f"{Fore.RED}⚠️  Press ESC anytime during automation to ABORT immediately!",
        f"{Fore.GREEN}🔒 HCA Healthcare Internal Use Only • v2.1"
    ]
    for notice in notices:
        print(notice.center(terminal_width))
    print()

    # === MAIN MENU HEADER ===
    print()
    header_lines = [
        f"{Fore.MAGENTA}{Style.BRIGHT}╔════════════════════════════════════════════════════════════════╗",
        f"{Fore.WHITE}{Style.BRIGHT}║                  MAIN OPERATIONS MENU                          ║",
        f"{Fore.MAGENTA}{Style.BRIGHT}╚════════════════════════════════════════════════════════════════╝"
    ]
    for line in header_lines:
        print(line.center(terminal_width))
    print()

    # 🔢 === "BIG TEXT" STYLE FOR NUMBERING ===
    # This is the closest you can get to "text_size = 20" in a terminal
    BIG_NUMBER = Fore.WHITE + Style.BRIGHT  # Bright white = maximum visibility
    MENU_TEXT = Fore.WHITE + Style.BRIGHT

    # 🔧 === YOUR PADDING VALUES (adjust these numbers) ===
    PAD_1 = 75
    PAD_2 = 75
    PAD_3 = 75
    PAD_4 = 75
    PAD_5 = 75
    PAD_6 = 75
    PAD_7 = 75
    PAD_0 = 75

    def print_item(pad, num, emoji, label):
        # Format: "1)  ⚙️  Configuration Management"
        line = f"{BIG_NUMBER}{num}){Style.RESET_ALL}  {MENU_TEXT}{emoji}  {label}{Style.RESET_ALL}"
        print(" " * pad + line)
        print()  # blank line

    # Print menu
    print_item(PAD_1, "1", "⚙️", "Configuration Management(First-time user? Set your config (Option 1) before starting!)")
    print_item(PAD_2, "2", "🚀", "Run FULL End-to-End Process")
    print_item(PAD_3, "3", "🖥️", "Step 1: ESAF UI Automation")
    print_item(PAD_4, "4", "🧹", "Step 2: Merge & Cleanup")
    print_item(PAD_5, "5", "👥", "Step 3: Assign Requests to Team")
    print_item(PAD_6, "6", "📊", "Step 4: Summary + Pivot")
    print_item(PAD_7, "7", "📈", "Step 5: Interactive Dashboard")
    print_item(PAD_0, "0", "❌", "Exit ESAF AutoPilot")

    print()
    footer = f"{Style.DIM}{Fore.WHITE}ESAF AutoPilot™ • Streamlining Access Request Management"
    print(footer.center(terminal_width))
    print()

def run_full_process():
    print(f"\n{Fore.GREEN}🚀 INITIATING FULL END-TO-END PROCESS")
    print(f"{Fore.CYAN}═══════════════════════════════════════════════════════════════")
    
    if not validate_config():
        print(f"{Fore.RED}❌ Configuration is missing or invalid!")
        print(f"{Fore.RED}Please run Option 1 to set up your configuration first.")
        return

    config = load_config()
    DOWNLOADS_FOLDER = resolve_downloads_folder(config)
    print(f"{Fore.CYAN}📁 Using Downloads folder: {Fore.WHITE}{DOWNLOADS_FOLDER}")

    steps = [
        (SCRIPTS["step1"], "Step 1: ESAF UI Automation"),
        (SCRIPTS["step2"], "Step 2: Merge & Cleanup"),
        (SCRIPTS["step3"], "Step 3: Assign Requests to Team"),
        (SCRIPTS["step4"], "Step 4: Summary + Pivot"),
        (SCRIPTS["step5"], "Step 5: Interactive Dashboard")
    ]

    try:
        for i, (script, step_name) in enumerate(steps, 1):
            print(f"\n{Fore.CYAN}🔹 Step {i}/5: {step_name}")
            print(f"{Fore.CYAN}{'─' * 60}")
            success = run_script(script, step_name)
            if not success:
                print(f"{Fore.RED}🛑 Full process stopped due to error in {step_name}")
                return
            print(f"{Fore.GREEN}✅ {step_name} completed successfully!")
        
        print(f"\n{Fore.GREEN}🎉 FULL END-TO-END PROCESS COMPLETED SUCCESSFULLY!")
        print(f"{Fore.CYAN}═══════════════════════════════════════════════════════════════")
        
    except KeyboardInterrupt:
        print(f"\n{Fore.RED}🛑 Process manually aborted. Returning to menu...")

def crud_menu():
    config = load_config() or {}
    while True:
        print(f"\n{Fore.CYAN}🛠️  CONFIGURATION MANAGEMENT CENTER")
        print(f"{Fore.CYAN}═══════════════════════════════════════════════════════════════")
        print(f"{Fore.YELLOW}0) 🌐 Edit ESAF URL")
        print(f"{Fore.YELLOW}1) 🖱️  Capture Mouse Coordinates")
        print(f"{Fore.YELLOW}2) 🔍 Edit Keywords")
        print(f"{Fore.YELLOW}3) 👥 Edit Assignees")
        print(f"{Fore.YELLOW}4) 📁 Set Downloads Folder")
        print(f"{Fore.YELLOW}5) 📥 Edit Queues")
        print(f"{Fore.YELLOW}6) ⚙️  Edit Rules (Thresholds)")
        print(f"{Fore.YELLOW}7) 🔄 Reset to Defaults")
        print(f"{Fore.YELLOW}8) 📄 View Current Config")
        print(f"{Fore.YELLOW}9) ◀️  Back to Main Menu")
        print(f"{Fore.CYAN}═══════════════════════════════════════════════════════════════")
        choice = input(f"{Fore.CYAN}Choose option (0-9): ").strip()
        if choice == "0":
            edit_url(config)
        elif choice == "1":
            capture_mouse_coordinates(config)
        elif choice == "2":
            edit_list(config, "keywords", "Keywords")
        elif choice == "3":
            edit_list(config, "assignees", "Assignees")
        elif choice == "4":
            folder = input(f"{Fore.CYAN}📁 Enter Downloads folder path (or 'AUTO' for default): ").strip()
            if not folder:
                folder = "AUTO"
            config["downloads_folder"] = folder
            save_config(config)
        elif choice == "5":
            edit_list(config, "queues", "Queues")
        elif choice == "6":
            edit_rules(config)
        elif choice == "7":
            reset_to_defaults()
            config = load_config() or {}
        elif choice == "8":
            view_config(config)
        elif choice == "9":
            break
        else:
            print(f"{Fore.YELLOW}⚠️ Invalid choice.")

def main():
    ensure_defaults()
    while True:
        try:
            show_menu()
            choice = input(f"\n{Fore.CYAN}Enter your choice (0-7): ").strip()
            if choice == "1":
                crud_menu()
            elif choice == "2":
                run_full_process()
            elif choice == "3":
                run_script(SCRIPTS["step1"], "Step 1: ESAF UI Automation")
            elif choice == "4":
                run_script(SCRIPTS["step2"], "Step 2: Merge & Cleanup")
            elif choice == "5":
                run_script(SCRIPTS["step3"], "Step 3: Assign Requests to Team")
            elif choice == "6":
                run_script(SCRIPTS["step4"], "Step 4: Summary + Pivot")
            elif choice == "7":
                run_script(SCRIPTS["step5"], "Step 5: Interactive Dashboard")
            elif choice == "0":
                print(f"\n{Fore.RED}👋 Thank you for using ESAF AutoPilot™. Goodbye!")
                print(f"{Fore.CYAN}ESAF AutoPilot™ • Streamlining Access Request Management")
                break
            else:
                print(f"{Fore.YELLOW}⚠️ Invalid choice. Please enter 0-7.")
            if choice != "0":
                input(f"\n{Fore.CYAN}Press Enter to return to main menu...")
        except KeyboardInterrupt:
            print(f"\n{Fore.RED}🛑 Interrupted. Returning to main menu...")
            input(f"\n{Fore.CYAN}Press Enter to continue...")

if __name__ == "__main__":
    main()