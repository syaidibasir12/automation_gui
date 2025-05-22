import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import json

SCRIPT_DIR = "Documents/AutomationExcel/IPH_Automation_TELE"
SETTINGS_FILE = os.path.join(SCRIPT_DIR, "settings.json")

# Load last used settings
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r") as f:
            return json.load(f)
    return {"downloads_path": "", "report_path": ""}

# Save settings
def save_settings(downloads_path, report_path):
    with open(SETTINGS_FILE, "w") as f:
        json.dump({
            "downloads_path": downloads_path,
            "report_path": report_path
        }, f)

def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        downloads_entry.delete(0, tk.END)
        downloads_entry.insert(0, folder)

def browse_report_file():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file:
        report_entry.delete(0, tk.END)
        report_entry.insert(0, file)

def run_scripts():
    downloads_path = downloads_entry.get().strip()
    report_path = report_entry.get().strip()

    if not os.path.isdir(downloads_path):
        messagebox.showerror("Invalid Folder", "Please select a valid Downloads folder.")
        return

    if not os.path.isfile(report_path):
        messagebox.showerror("Invalid File", "Please select a valid Report file.")
        return

    try:
        main_script = os.path.join(SCRIPT_DIR, "Update_IPH_Telemarketing.py")
        working_days_script = os.path.join(SCRIPT_DIR, "Update_WorkingDays.py")

        subprocess.run(["python", main_script, downloads_path], check=True)
        subprocess.run(["python", working_days_script, downloads_path, report_path], check=True)

        save_settings(downloads_path, report_path)
        messagebox.showinfo("Success", "All scripts completed successfully!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"A script failed:\n{e}")
    except Exception as e:
        messagebox.showerror("Unexpected Error", str(e))

# GUI setup
root = tk.Tk()
root.title("IPH TELEMARKETING AUTOMATION")
root.geometry("500x300")

# Load saved settings
settings = load_settings()

tk.Label(root, text="Downloads Folder:").pack(pady=(10, 0))
downloads_entry = tk.Entry(root, width=60)
downloads_entry.insert(0, settings.get("downloads_path", ""))
downloads_entry.pack()
tk.Button(root, text="Browse...", command=browse_folder).pack(pady=5)

tk.Label(root, text="Target Report File:").pack(pady=(10, 0))
report_entry = tk.Entry(root, width=60)
report_entry.insert(0, settings.get("report_path", ""))
report_entry.pack()
tk.Button(root, text="Browse...", command=browse_report_file).pack(pady=5)

tk.Button(root, text="Run Scripts", command=run_scripts, bg="green", fg="white", font=("Helvetica", 12)).pack(pady=20)

root.mainloop()
