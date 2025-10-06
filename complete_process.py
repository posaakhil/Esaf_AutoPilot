# complete_process.py — PyInstaller-safe, no subprocess, all-in-one execution
import sys
import os
import json
import tempfile
import importlib.util
from pathlib import Path

# ===== CONFIG & PATH HELPERS =====
def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

CONFIG_FILE = "esaf_config.json"
if not os.path.exists(CONFIG_FILE):
    print("[ERROR] esaf_config.json not found!")
    sys.exit(1)

with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
    config = json.load(f)

# ===== SCRIPT EXECUTION ENGINE =====
def run_script_in_memory(script_name, step_name, extra_args=None):
    """Execute a bundled Python script in the current process (no subprocess)"""
    print(f"\n[INFO] 🚀 Starting {step_name}...")
    
    script_path = get_resource_path(script_name)
    if not os.path.exists(script_path):
        print(f"[ERROR] ❌ Script not found: {script_name}")
        return False

    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            code = f.read()

        # Prepare globals (mimic __main__ environment)
        script_globals = {
            '__name__': '__main__',
            '__file__': script_path,
            'sys': sys,
            'os': os,
            'json': json,
            'tempfile': tempfile,
            'Path': Path,
        }

        # Inject --auto if needed
        if extra_args and '--auto' in extra_args:
            original_argv = sys.argv.copy()
            sys.argv = [script_name, '--auto']
        else:
            original_argv = None

        # Optional: inject common modules if needed
        try:
            import pandas as pd
            script_globals['pd'] = pd
        except ImportError:
            pass
        try:
            from openpyxl import load_workbook
            script_globals['load_workbook'] = load_workbook
        except ImportError:
            pass
        try:
            import plotly.express as px
            import plotly.graph_objects as go
            import plotly.io as pio
            script_globals.update({'px': px, 'go': go, 'pio': pio})
        except ImportError:
            pass

        # Execute!
        exec(code, script_globals)
        print(f"[SUCCESS] ✅ {step_name} completed!")

        # Restore argv
        if original_argv is not None:
            sys.argv = original_argv

        return True

    except KeyboardInterrupt:
        print(f"\n[ABORT] ⚠️ {step_name} interrupted by user.")
        return False
    except SystemExit as e:
        if e.code == 0:
            print(f"[SUCCESS] ✅ {step_name} exited cleanly.")
            return True
        else:
            print(f"[ERROR] ❌ {step_name} exited with error code {e.code}.")
            return False
    except Exception as e:
        print(f"[ERROR] ❌ Failed to run {step_name}: {e}")
        return False

# ===== MAIN WORKFLOW =====
def main():
    print("=" * 70)
    print("🚀 ESAF FULL END-TO-END AUTOMATION (PyInstaller-Safe Version)")
    print("=" * 70)

    steps = [
        ("esaf_automation.py", "Step 1: ESAF UI Automation"),
        ("merge_and_cleanup.py", "Step 2: Merge & Cleanup", ["--auto"]),
        ("Data_Analysis_Split.py", "Step 3: Assign Requests", ["--auto"]),
        ("summary_pivot.py", "Step 4: Summary & Pivot", ["--auto"]),
        ("Interactive_Dashboard.py", "Step 5: Interactive Dashboard", []),
    ]

    for script, name, *args in steps:
        args = args[0] if args else []
        success = run_script_in_memory(script, name, args)
        if not success:
            print(f"\n[CRITICAL] 💥 Workflow halted at {name}.")
            sys.exit(1)

    print("\n" + "=" * 70)
    print("🎉 ENTIRE ESAF WORKFLOW COMPLETED SUCCESSFULLY!")
    print("📁 Output: Today_Assignment.xlsx + esaf_executive_dashboard.html")
    print("=" * 70)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[INFO] Script manually aborted.")
    except Exception as e:
        print(f"\n[FATAL] Unexpected error: {e}")
        sys.exit(1)
