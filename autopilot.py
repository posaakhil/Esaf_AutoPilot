import subprocess
import sys
import os
import json
import shutil
import keyboard
from colorama import Fore, Style, init
import pyautogui
import tempfile

# Initialize colorama with full compatibility
init(autoreset=True, convert=True, strip=False)

CONFIG_FILE = "esaf_config.json"
DEFAULTS_FILE = "esaf_config_defaults.json"
BACKUP_FILE = "esaf_config_backup.json"

# ===== SCRIPT FILENAMES =====
SCRIPTS = {
    "step1": "esaf_automation.py",
    "step2": "merge_and_cleanup.py", 
    "step3": "Data_Analysis_Split.py",
    "step4": "summary_pivot.py",
    "step5": "Interactive_Dashboard.py",
    "complete": "complete_process.py"
}

def get_ascii_banner():
    """Return ASCII banner as plain string (safe for all terminals)"""
    return f"""
{Fore.CYAN}+==============================================================================+
{Fore.CYAN}|                                                                              |
{Fore.CYAN}|  {Fore.BLUE}███████╗███████╗ █████╗ ███████╗     █████╗ ██╗   ██╗████████╗ ██████╗ {Fore.CYAN}  |
{Fore.CYAN}|  {Fore.BLUE}██╔════╝██╔════╝██╔══██╗██╔════╝    ██╔══██╗██║   ██║╚══██╔══╝██╔═══██╗{Fore.CYAN}  |
{Fore.CYAN}|  {Fore.BLUE}█████╗  ███████╗███████║█████╗      ███████║██║   ██║   ██║   ██║   ██║{Fore.CYAN}  |
{Fore.CYAN}|  {Fore.BLUE}██╔══╝  ╚════██║██╔══██║██╔══╝      ██╔══██║██║   ██║   ██║   ██║   ██║{Fore.CYAN}  |
{Fore.CYAN}|  {Fore.BLUE}███████╗███████║██║  ██║██          ██║  ██║╚██████╔╝   ██║   ╚██████╔╝{Fore.CYAN}  |
{Fore.CYAN}|  {Fore.BLUE}╚══════╝╚══════╝╚═╝  ╚═╝╚═          ╚═╝  ╚═╝ ╚═════╝    ╚═╝    ╚═════╝ {Fore.CYAN}  |
{Fore.CYAN}|                                                                              |
{Fore.CYAN}|  {Fore.GREEN}+--------------------------------------------------------------------------+  {Fore.CYAN}|
{Fore.CYAN}|  {Fore.GREEN}|           AUTOMATED REQUEST MANAGEMENT SYSTEM - ESAF AutoPilot™         |  {Fore.CYAN}|
{Fore.CYAN}|  {Fore.GREEN}+--------------------------------------------------------------------------+  {Fore.CYAN}|
{Fore.CYAN}|                                                                              |
{Fore.CYAN}+==============================================================================+
"""

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def extract_and_run_script(script_name, step_name):
    """Extract script from bundle, modify it to remove menu, and run it"""
    try:
        print(f"\n{Fore.CYAN}[RUN] Starting {step_name}...")
        print(f"{Fore.YELLOW}[ABORT] Press ESC anytime to stop!")
        
        # Get the bundled script
        bundled_path = get_resource_path(script_name)
        
        if not os.path.exists(bundled_path):
            print(f"{Fore.RED}[ERROR] Script not found in bundle: {script_name}")
            return False
        
        # Read the script content
        with open(bundled_path, 'r', encoding='utf-8') as f:
            original_content = f.read()
        
        # MODIFY THE SCRIPT: Remove menu code and ensure it runs the main functionality
        modified_content = modify_script_content(original_content, script_name)
        
        # Create temporary file with modified content
        temp_dir = tempfile.mkdtemp()
        temp_script_path = os.path.join(temp_dir, script_name)
        
        with open(temp_script_path, 'w', encoding='utf-8') as f:
            f.write(modified_content)
        
        # Run the modified script
        proc = subprocess.Popen(
            [sys.executable, temp_script_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            encoding='utf-8',
            errors='replace'
        )
        
        # Stream output with colors
        while True:
            check_abort()
            output = proc.stdout.readline()
            if output == '' and proc.poll() is not None:
                break
            if output:
                print(output, end='', flush=True)
        
        rc = proc.poll()
        if rc == 0:
            print(f"\n{Fore.GREEN}[SUCCESS] {step_name} completed!")
            return True
        else:
            print(f"\n{Fore.RED}[FAILED] {step_name} exited with code {rc}")
            return False
            
    except KeyboardInterrupt:
        print(f"\n{Fore.RED}[ABORT] Process interrupted by user.")
        if 'proc' in locals():
            proc.terminate()
            proc.wait()
        return False
    except Exception as e:
        print(f"\n{Fore.RED}[ERROR] Failed to run {step_name}: {e}")
        return False

def modify_script_content(original_content, script_name):
    """Modify script content to ensure it runs the actual functionality, not the menu"""
    
    # Remove any main menu code and ensure the script runs its main function
    lines = original_content.split('\n')
    modified_lines = []
    
    # Keep all lines except the main menu invocation
    in_main_block = False
    main_removed = False
    
    for line in lines:
        # Skip the main menu invocation but keep the function definitions
        if 'show_menu()' in line or 'main()' in line and 'def ' not in line:
            if script_name == "complete_process.py":
                # For complete_process, we want it to run its main function
                if 'main()' in line and 'def ' not in line:
                    modified_lines.append(line)  # Keep main() call for complete_process
            else:
                # For other scripts, replace menu calls with their actual functionality
                if 'merge_and_cleanup.py' in script_name:
                    modified_lines.append("    main()  # Run the actual merge functionality")
                elif 'esaf_automation.py' in script_name:
                    modified_lines.append("    main()  # Run the actual automation")
                elif 'Data_Analysis_Split.py' in script_name:
                    modified_lines.append("    main()  # Run the actual data analysis")
                elif 'summary_pivot.py' in script_name:
                    modified_lines.append("    main()  # Run the actual summary creation")
                elif 'Interactive_Dashboard.py' in script_name:
                    modified_lines.append("    main()  # Run the actual dashboard generation")
                main_removed = True
        else:
            modified_lines.append(line)
    
    # If we didn't find and replace the main call, add it at the end
    if not main_removed and script_name != "complete_process.py":
        modified_lines.append("")
        modified_lines.append("# Auto-execute main functionality")
        modified_lines.append("if __name__ == '__main__':")
        modified_lines.append("    main()")
    
    return '\n'.join(modified_lines)

def check_abort():
    """Check if user pressed ESC - can be called anywhere"""
    if keyboard.is_pressed('esc'):
        print(f"\n{Fore.RED}[ABORT] EMERGENCY STOP: ESC key pressed!")
        raise KeyboardInterrupt("User pressed ESC")

# ... [REST OF YOUR CONFIG FUNCTIONS REMAIN EXACTLY THE SAME] ...

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
            },
            "no_requests_text": "There are no requests available",
            "base_assignment_folder": "assignment_"
        }
        with open(DEFAULTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(defaults, f, indent=4, ensure_ascii=False)
        print(f"{Fore.GREEN}[OK] Created default config: {DEFAULTS_FILE}")

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return None
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"{Fore.RED}[ERROR] Failed to load config: {e}")
        return None

# ... [ALL YOUR OTHER FUNCTIONS REMAIN EXACTLY THE SAME] ...

def show_menu():
    try:
        terminal_width = shutil.get_terminal_size().columns
    except:
        terminal_width = 80
    os.system('cls' if os.name == 'nt' else 'clear')
    print(get_ascii_banner())
    print()
    config_status = "[OK] CONFIGURED" if validate_config() else "[ERROR] NOT CONFIGURED"
    status_color = Fore.GREEN if validate_config() else Fore.RED
    status_line = f"[SYSTEM STATUS: {status_color}{config_status}{Fore.RESET}]"
    print(status_line.center(terminal_width))
    print()
    notices = [
        f"{Fore.CYAN}[TIP] First-time user? Configure (Option 1) before running full process!",
        f"{Fore.RED}[WARNING] Press ESC anytime during automation to ABORT immediately!",
        f"{Fore.GREEN}[INFO] HCA Healthcare Internal Use Only • v2.1"
    ]
    for notice in notices:
        print(notice.center(terminal_width))
    print()
    print(f"{Fore.MAGENTA} MAIN OPERATIONS MENU ".center(terminal_width, "="))
    print()
    menu_items = [
        f"{Fore.WHITE}1) {Fore.CYAN}Configuration Management",
        f"{Fore.WHITE}2) {Fore.GREEN}Run FULL End-to-End Process (complete_process.py)",
        f"{Fore.WHITE}3) {Fore.BLUE}Step 1: ESAF UI Automation (esaf_automation.py)",
        f"{Fore.WHITE}4) {Fore.BLUE}Step 2: Merge & Cleanup (merge_and_cleanup.py)",
        f"{Fore.WHITE}5) {Fore.BLUE}Step 3: Assign Requests to Team (Data_Analysis_Split.py)",
        f"{Fore.WHITE}6) {Fore.BLUE}Step 4: Summary + Pivot (summary_pivot.py)",
        f"{Fore.WHITE}7) {Fore.BLUE}Step 5: Interactive Dashboard (Interactive_Dashboard.py)",
        f"{Fore.WHITE}0) {Fore.RED}Exit ESAF AutoPilot"
    ]
    for item in menu_items:
        print("    " + item)
    print()
    footer = f"{Fore.CYAN}ESAF AutoPilot™ • Streamlining Access Request Management"
    print(footer.center(terminal_width))

def crud_menu():
    """Configuration Management Menu"""
    config = load_config() or {}
    while True:
        print(f"\n{Fore.CYAN}[CONFIG] CONFIGURATION MANAGEMENT")
        print(f"{Fore.CYAN}=" * 60)
        print(f"{Fore.YELLOW}0) {Fore.CYAN}Edit ESAF URL")
        print(f"{Fore.YELLOW}1) {Fore.CYAN}Capture Mouse Coordinates")
        print(f"{Fore.YELLOW}2) {Fore.CYAN}Edit Keywords")
        print(f"{Fore.YELLOW}3) {Fore.CYAN}Edit Assignees")
        print(f"{Fore.YELLOW}4) {Fore.CYAN}Set Downloads Folder")
        print(f"{Fore.YELLOW}5) {Fore.CYAN}Edit Queues")
        print(f"{Fore.YELLOW}6) {Fore.CYAN}Edit Rules (Thresholds)")
        print(f"{Fore.YELLOW}7) {Fore.CYAN}Reset to Defaults")
        print(f"{Fore.YELLOW}8) {Fore.CYAN}View Current Config")
        print(f"{Fore.YELLOW}9) {Fore.RED}Back to Main Menu")
        print(f"{Fore.CYAN}=" * 60)
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
            folder = input(f"{Fore.CYAN}Enter Downloads folder path (or 'AUTO'): ").strip()
            config["downloads_folder"] = folder if folder else "AUTO"
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
            print(f"{Fore.YELLOW}[WARN] Invalid choice.")

def main():
    ensure_defaults()
    while True:
        try:
            show_menu()
            choice = input(f"\n{Fore.CYAN}Enter your choice (0-7): ").strip()
            if choice == "1":
                crud_menu()
            elif choice == "2":
                extract_and_run_script(SCRIPTS["complete"], "Full End-to-End Process")
            elif choice == "3":
                extract_and_run_script(SCRIPTS["step1"], "Step 1: ESAF UI Automation")
            elif choice == "4":
                extract_and_run_script(SCRIPTS["step2"], "Step 2: Merge & Cleanup")
            elif choice == "5":
                extract_and_run_script(SCRIPTS["step3"], "Step 3: Assign Requests to Team")
            elif choice == "6":
                extract_and_run_script(SCRIPTS["step4"], "Step 4: Summary + Pivot")
            elif choice == "7":
                extract_and_run_script(SCRIPTS["step5"], "Step 5: Interactive Dashboard")
            elif choice == "0":
                print(f"\n{Fore.CYAN}Thank you for using ESAF AutoPilot™. Goodbye!")
                break
            else:
                print(f"{Fore.YELLOW}[WARN] Please enter 0-7.")
            
            if choice != "0":
                input(f"\n{Fore.CYAN}Press Enter to return to main menu...")
                
        except KeyboardInterrupt:
            print(f"\n{Fore.RED}[ABORT] Interrupted. Returning to menu...")
            input(f"\n{Fore.CYAN}Press Enter to continue...")
        except Exception as e:
            print(f"\n{Fore.RED}[ERROR] Unexpected error: {e}")
            input(f"\n{Fore.CYAN}Press Enter to continue...")

if __name__ == "__main__":
    # Ensure UTF-8 and color support
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding='utf-8')
    init(autoreset=True, convert=True, strip=False)
    main()
