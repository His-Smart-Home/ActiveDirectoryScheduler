import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import urllib.request
import io
import pandas as pd
from datetime import datetime
from tkcalendar import DateEntry
import os
import json

SETTINGS_FILE = 'settings.json'
LOGO_URL = 'https://hissmarthome.uk/img/HisSmartHome-black.png'
ICON_URL = 'https://hissmarthome.uk/img/HisSmartHome-grey-tile.jpg'

class SchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Event Scheduler")
        self.data = pd.DataFrame(columns=["user", "action", "value", "datetime"])
        self.current_csv_path = None

        self.set_window_icon()
        self.create_widgets()
        self.load_last_csv()

    def set_window_icon(self):
        try:
            req = urllib.request.Request(ICON_URL, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as u:
                raw_data = u.read()
            icon_image = Image.open(io.BytesIO(raw_data))
            icon_image = icon_image.resize((32, 32), Image.Resampling.LANCZOS)
            self.icon = ImageTk.PhotoImage(icon_image)
            self.root.iconphoto(True, self.icon)
        except Exception as e:
            print(f"Error setting window icon: {e}")

    def create_widgets(self):
        # Logo
        try:
            req = urllib.request.Request(LOGO_URL, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as u:
                raw_data = u.read()
            im = Image.open(io.BytesIO(raw_data))
            im = im.resize((344, 64), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(im)
            logo_label = ttk.Label(self.root, image=self.logo, cursor="hand2")
            logo_label.pack(pady=(10, 0))
            logo_label.bind("<Button-1>", self.on_logo_click)

        except Exception as e:
            print(f"Error loading logo: {e}")

        headers = {"user": "Username", "action": "Action", "value": "Value", "datetime": "Scheduled Time"}
        self.tree = ttk.Treeview(self.root, columns=self.data.columns.tolist(), show='headings')
        for col in self.data.columns:
            self.tree.heading(col, text=headers[col])
            self.tree.column(col, width=150)
        self.tree.pack(fill=tk.BOTH, expand=True, pady=10)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        load_btn = ttk.Button(btn_frame, text="Load CSV", command=self.load_csv)
        load_btn.pack(side=tk.LEFT, padx=5)

        add_btn = ttk.Button(btn_frame, text="Add Event", command=self.add_event)
        add_btn.pack(side=tk.LEFT, padx=5)

        delete_btn = ttk.Button(btn_frame, text="Delete Event", command=self.delete_event)
        delete_btn.pack(side=tk.LEFT, padx=5)

        save_btn = ttk.Button(btn_frame, text="Save CSV", command=self.save_csv)
        save_btn.pack(side=tk.LEFT, padx=5)

    def on_logo_click(self, event=None):
        messagebox.showinfo("About", "Active Directory Scheduler\nPowered by His Smart Home\nOriginal idea by Ben Morrell\nCoded to life by NJ Bendall")


    def load_last_csv(self):
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                last_path = settings.get('last_csv_path', '')
                if os.path.exists(last_path):
                    self.load_csv(path=last_path)
                else:
                    self.prompt_save_new_csv()
        else:
            self.prompt_save_new_csv()

    def prompt_save_new_csv(self):
        messagebox.showinfo("CSV Not Found", "No existing CSV found. Please select a location to create one.")
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.data.to_csv(file_path, index=False)
            self.current_csv_path = file_path
            self.save_settings(file_path)

    def save_settings(self, path):
        with open(SETTINGS_FILE, 'w') as f:
            json.dump({'last_csv_path': path}, f)

    def load_csv(self, path=None):
        if not path:
            path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            self.data = pd.read_csv(path, parse_dates=['datetime'])
            self.current_csv_path = path
            self.refresh_tree()
            self.save_settings(path)

    def refresh_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for _, row in self.data.iterrows():
            self.tree.insert('', tk.END, values=row.tolist())

    def add_event(self):
        top = tk.Toplevel(self.root)
        top.title("Add New Event")

        headers = {"user": "Username", "action": "Action", "value": "Value", "datetime": "Scheduled Time"}
        entries = {}

        for i, field in enumerate(self.data.columns):
            if field != "datetime":
                ttk.Label(top, text=headers[field]).grid(row=i, column=0, padx=5, pady=5, sticky="e")
            if field == "action":
                combo = ttk.Combobox(top, values=["disable", "enable", "addtogroup", "removefromgroup"], state="readonly")
                combo.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                entries[field] = combo
            elif field == "datetime":
                ttk.Label(top, text=headers[field]).grid(row=i, column=0, padx=5, pady=5, sticky="ne")

                frame = ttk.Frame(top)
                frame.grid(row=i, column=1, columnspan=2, padx=5, pady=5, sticky="w")

                date_entry = DateEntry(frame, width=12, background='darkblue', foreground='white', borderwidth=2)
                date_entry.pack(side="top", anchor="w")

                time_var = tk.StringVar()
                time_options = [f"{h:02}:{m:02}" for h in range(24) for m in range(0, 60, 15)]
                time_menu = ttk.Combobox(frame, textvariable=time_var, values=time_options, state="readonly", width=6)
                time_menu.pack(side="top", anchor="e", pady=(5, 0))
                time_menu.set("00:00")

                entries[field] = (date_entry, time_menu)
            else:
                entry = ttk.Entry(top)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                entries[field] = entry

        def save_event():
            try:
                new_event = {}
                for field in self.data.columns:
                    if field == "datetime":
                        date_part = entries[field][0].get_date()
                        time_part = entries[field][1].get()
                        dt = datetime.strptime(f"{date_part} {time_part}", "%Y-%m-%d %H:%M")
                        new_event[field] = dt
                    else:
                        new_event[field] = entries[field].get()
                self.data.loc[len(self.data)] = new_event
                self.refresh_tree()
                top.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Invalid data: {e}")

        ttk.Button(top, text="Add", command=save_event).grid(row=len(self.data.columns), column=0, columnspan=3, pady=10)

    def delete_event(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            self.tree.delete(selected_item)
            self.data = self.data.drop(self.data.index[index]).reset_index(drop=True)
        else:
            messagebox.showwarning("No selection", "Please select an event to delete.")

    def save_csv(self):
        if self.current_csv_path:
            self.data.to_csv(self.current_csv_path, index=False)
            messagebox.showinfo("Saved", f"Events saved to {self.current_csv_path}")
        else:
            file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if file_path:
                self.data.to_csv(file_path, index=False)
                self.current_csv_path = file_path
                self.save_settings(file_path)
                messagebox.showinfo("Saved", f"Events saved to {file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SchedulerApp(root)
    root.mainloop()