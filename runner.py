import pandas as pd
from datetime import datetime
import json
import os
import subprocess
import logging

SETTINGS_FILE = 'settings.json'
LOG_FILE = 'task_runner.log'

# Setup logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# PowerShell interaction helpers
def run_powershell(command):
    result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)
    if result.returncode != 0:
        logging.error(f"PowerShell error for command '{command}': {result.stderr.strip()}")
        print(f"[ERROR] {result.stderr.strip()}")
    else:
        logging.info(f"Command succeeded: {command}\nOutput: {result.stdout.strip()}")
        print(f"[SUCCESS] {result.stdout.strip()}")

def disable_user(user):
    logging.info(f"Disabling user: {user}")
    run_powershell(f"Disable-ADAccount -Identity \"{user}\"")

def enable_user(user):
    logging.info(f"Enabling user: {user}")
    run_powershell(f"Enable-ADAccount -Identity \"{user}\"")

def add_to_group(user, group):
    logging.info(f"Adding user {user} to group {group}")
    run_powershell(f"Add-ADGroupMember -Identity \"{group}\" -Members \"{user}\"")

def remove_from_group(user, group):
    logging.info(f"Removing user {user} from group {group}")
    run_powershell(f"Remove-ADGroupMember -Identity \"{group}\" -Members \"{user}\" -Confirm:$false")

def run_task(row):
    user = row['user']
    action = row['action']
    value = row['value']

    logging.info(f"Executing task: {action} for {user} with value '{value}'")

    if action == 'disable':
        disable_user(user)
    elif action == 'enable':
        enable_user(user)
    elif action == 'addtogroup':
        add_to_group(user, value)
    elif action == 'removefromgroup':
        remove_from_group(user, value)
    else:
        logging.warning(f"Unknown action: {action}")
        print(f"[WARNING] Unknown action: {action}")

def main():
    logging.info("Task runner started.")

    if not os.path.exists(SETTINGS_FILE):
        logging.error("Settings file not found.")
        print("Settings file not found.")
        return

    with open(SETTINGS_FILE, 'r') as f:
        settings = json.load(f)

    csv_path = settings.get('last_csv_path')
    if not csv_path or not os.path.exists(csv_path):
        logging.error("CSV file path not found or invalid in settings.json.")
        print("CSV file path not found or invalid in settings.json.")
        return

    df = pd.read_csv(csv_path, parse_dates=['datetime'])
    now = datetime.now()
    due_tasks = df[df['datetime'] <= now]

    if due_tasks.empty:
        logging.info("No tasks due at this time.")
        print("No tasks due at this time.")
        return

    print(f"Running {len(due_tasks)} task(s)...")
    logging.info(f"Running {len(due_tasks)} due task(s)...")

    for idx, row in due_tasks.iterrows():
        run_task(row)

    # Remove executed tasks
    df = df[df['datetime'] > now]
    df.to_csv(csv_path, index=False)
    logging.info("Completed all due tasks and updated schedule.")
    print("Completed and updated schedule.")

if __name__ == '__main__':
    main()
