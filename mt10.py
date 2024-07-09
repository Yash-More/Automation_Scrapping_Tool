#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# working code

import os
import time
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
import webbrowser
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import csv
import threading
import pyautogui
import keyboard
import cv2
import datetime

# Configuration
Download_folder = r"C:\Users\Admin\Downloads\FB groups data"
Google_sheets_key_file = r"C:\Users\Admin\Desktop\Merge\automated-rune-428805-i4-fca1e1fbbafa.json"
google_sheets_workbook_id = "1OJFfQqf7tvoXcr9d-GtOQluNBig4H4t1TW44ctLl1OA"
processed_files = set()

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(Google_sheets_key_file, scope)
client = gspread.authorize(creds)
workbook = client.open_by_key(google_sheets_workbook_id)

def upload_to_google_sheets(file_path):
    print(f"Uploading {file_path} to Google Sheets")
    try:
        df = pd.read_excel(file_path)
        sheet_name = os.path.basename(file_path).replace('.xlsx', '')

        # Clean NaN and Infinity values using numpy
        df.replace([np.inf, -np.inf], np.nan, inplace=True)
        df.fillna('NaN', inplace=True)

        # Add a new sheet and upload the data
        worksheet = workbook.add_worksheet(title=sheet_name, rows=str(df.shape[0]), cols=str(df.shape[1]))
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
        print(f"Uploaded {file_path} to Google Sheets as {sheet_name} time : {datetime.datetime.now()}")
        processed_files.add(file_path)
    except Exception as e:
        print(f"Failed to upload {file_path}: {e}")

def scan_directory():
    for file_name in os.listdir(Download_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(Download_folder, file_name)
            if file_path not in processed_files:
                upload_to_google_sheets(file_path)

class FacebookGroupOpener:
    def __init__(self, root):
        self.root = root
        self.root.title("Facebook Group Opener")
        self.root.geometry("800x600")

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TButton', font=('Helvetica', 12), foreground='black', background='#FFD700', padding=10)
        self.style.map('TButton', foreground=[('pressed', 'black'), ('active', 'blue')], background=[('pressed', '!disabled', 'lightgray'), ('active', 'cyan')])
        self.style.configure('TLabel', font=('Helvetica', 12), foreground='black')
        self.style.configure('TListbox', font=('Helvetica', 12), foreground='black', background='white')
        self.style.configure('TEntry', font=('Helvetica', 12))

        self.url_list = []
        self.current_url_index = 0

        self.url_listbox = tk.Listbox(root, width=70, height=10, font=('Helvetica', 12), selectmode=tk.MULTIPLE, bg='white')
        self.url_listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.add_url_entry = ttk.Entry(root, width=70, font=('Helvetica', 12))
        self.add_url_entry.pack(pady=10)

        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=10)

        self.add_url_button = ttk.Button(self.button_frame, text="Add URL", command=self.add_url)
        self.add_url_button.grid(row=0, column=0, padx=10, pady=10)

        self.delete_url_button = ttk.Button(self.button_frame, text="Delete URL", command=self.delete_url)
        self.delete_url_button.grid(row=0, column=1, padx=10, pady=10)

        self.clear_all_button = ttk.Button(self.button_frame, text="Clear All", command=self.clear_all)
        self.clear_all_button.grid(row=0, column=2, padx=10, pady=10)

        self.upload_button = ttk.Button(self.button_frame, text="Upload CSV", command=self.upload_csv)
        self.upload_button.grid(row=1, column=0, padx=10, pady=10)

        self.start_button = ttk.Button(self.button_frame, text="Start", command=self.start)
        self.start_button.grid(row=1, column=1, padx=10, pady=10)

        self.pause_button = ttk.Button(self.button_frame, text="Pause", command=self.pause)
        self.pause_button.grid(row=2, column=0, padx=10, pady=10)

        self.resume_button = ttk.Button(self.button_frame, text="Resume", command=self.resume)
        self.resume_button.grid(row=2, column=1, padx=10, pady=10)

        self.url_count_label = ttk.Label(root, text="URL Count: 0")
        self.url_count_label.pack(pady=10)

        self.url_opened_count_label = ttk.Label(root, text="URLs Opened: 0")
        self.url_opened_count_label.pack(pady=10)

        self.url_left_to_open_label = ttk.Label(root, text="URLs Left: 0")
        self.url_left_to_open_label.pack(pady=10)

        self.running = False
        self.paused = False
        self.pause_loop = threading.Event()
        self.end_loop = threading.Event()
        self.pause_loop.set()  # Start in the running state

    def add_url(self):
        url = self.add_url_entry.get().strip()
        if url:
            self.url_list.append(url)
            self.url_listbox.insert(tk.END, url)
            self.add_url_entry.delete(0, tk.END)
            self.update_url_count_label()

    def update_url_count_label(self):
        self.url_count_label.config(text=f"URL Count: {len(self.url_list)}")
        self.update_url_left_to_open_label()

    def delete_url(self):
        selected_indices = self.url_listbox.curselection()
        if selected_indices:
            selected_indices = list(selected_indices)
            selected_indices.sort(reverse=True)
            for index in selected_indices:
                del self.url_list[index]
                self.url_listbox.delete(index)
            self.update_url_count_label()
            self.update_url_opened_count_label()
            self.update_url_left_to_open_label()

    def clear_all(self):
        self.url_list.clear()
        self.url_listbox.delete(0, tk.END)
        self.update_url_count_label()
        self.update_url_opened_count_label()
        self.update_url_left_to_open_label()

    def upload_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, newline='') as csvfile:
                csvreader = csv.reader(csvfile)
                for row in csvreader:
                    if row:
                        self.url_list.append(row[0])
                        self.url_listbox.insert(tk.END, row[0])
            self.update_url_count_label()
            self.update_url_left_to_open_label()

    def open_facebook_group(self, url):
        webbrowser.open(url)
        time.sleep(10)
        self.update_url_opened_count_label()

    def update_url_opened_count_label(self):
        opened_count = self.current_url_index
        self.url_opened_count_label.config(text=f"URLs Opened: {opened_count}")
        self.update_url_left_to_open_label()

    def update_url_left_to_open_label(self):
        left_to_open_count = len(self.url_list) - self.current_url_index
        self.url_left_to_open_label.config(text=f"URLs Left: {left_to_open_count}")

    def start(self):
        if not self.running and self.url_list:
            self.running = True
            self.paused = False
            self.current_url_index = 0
            self.end_loop.clear()
            self.automation_thread = threading.Thread(target=self.automate_group_actions)
            self.automation_thread.start()

    def automate_group_actions(self):
        for url in self.url_list:
            if self.end_loop.is_set():
                break

            self.pause_loop.wait()

            self.root.after(0, self.update_status, f"Processing URL: {url}")
            self.root.after(0, self.open_facebook_group, url)

            self.root.after(0, self.update_status, "Scrolling in the page's discussion tab")
            pyautogui.moveTo(324, 333, duration=1)
            pyautogui.moveTo(468, 673, duration=1)
            time.sleep(15)
            pyautogui.scroll(-258)
            time.sleep(10)
            pyautogui.scroll(-945)
            time.sleep(5)
            pyautogui.scroll(-835)

            extension_icon_path = r'extension_icon.png'
            ok_button_path = r'button_ok.png'
            run_button_path = r'run_button.png'
            export_button_path = r'export_button.png'

            self.root.after(0, self.update_status, "Interacting with the extension")
            pyautogui.moveTo(1313, 293, duration=1)
            pyautogui.moveTo(1137, 306, duration=1)
            pyautogui.moveTo(631, 132, duration=1)
            pyautogui.moveTo(609, 85, duration=1)
            self.move_and_click(extension_icon_path)

            pyautogui.moveTo(441, 173, duration=1)
            self.move_and_click(ok_button_path)

            pyautogui.moveTo(441, 173, duration=1)
            self.move_and_click(ok_button_path)

            pyautogui.moveTo(617, 339, duration=1)
            self.move_and_click(run_button_path)

            self.root.after(0, self.update_status, "Waiting for the process to finish")

            # Wait for the export button to become visible and click it
            while not self.move_and_click(export_button_path):
                self.root.after(0, self.update_status, "Export button not found, waiting...")
                time.sleep(5)

            self.root.after(0, self.update_status, f"Finished processing URL: {url}")
            self.current_url_index += 1
            self.root.after(0, self.update_url_opened_count_label)
            self.root.after(0, self.update_url_left_to_open_label)

    def move_and_click(self, image_path):
        pos = self.find_img_on_screen(image_path)
        if pos:
            self.root.after(0, self.update_status, f"Found {image_path} at {pos}, moving and clicking.")
            pyautogui.moveTo(pos[0], pos[1], duration=1)
            pyautogui.click()
            return True
        else:
            self.root.after(0, self.update_status, f"Could not find {image_path} on screen.")
            return False

    def find_img_on_screen(self, image_path, confidence=0.8):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        template = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        w, h = template.shape[::-1]

        gray_screenshot = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        result = cv2.matchTemplate(gray_screenshot, template, cv2.TM_CCOEFF_NORMED)
        loc = np.where(result >= confidence)

        for pt in zip(*loc[::-1]):
            return (pt[0] + w // 2, pt[1] + h // 2)

        return None

    def update_status(self, status):
        self.url_opened_count_label.config(text=status)

    def pause(self):
        self.paused = True
        self.pause_loop.clear()

    def resume(self):
        self.paused = False
        self.pause_loop.set()

    def control_loop(self):
        while not self.end_loop.is_set():
            if keyboard.is_pressed('p'):
                self.pause()
                while not keyboard.is_pressed('r') and not keyboard.is_pressed('e'):
                    time.sleep(0.1)  # Wait until 'r' or 'e' is pressed

            if keyboard.is_pressed('r'):
                self.resume()

            if keyboard.is_pressed('e'):
                self.end_loop.set()
                self.pause_loop.set()  # Resume any paused threads so they can exit
                self.running = False
                break

            time.sleep(0.1)  # Small delay to prevent high CPU usage


if __name__ == '__main__':
    # Start Google Sheets scanning thread
    def google_sheets_thread():
        while True:
            scan_directory()
            time.sleep(240)  # Wait for 4 minutes before scanning again

    google_sheets_thread = threading.Thread(target=google_sheets_thread)
    google_sheets_thread.start()

    # Start Tkinter application
    root = tk.Tk()
    app = FacebookGroupOpener(root)
    control_thread = threading.Thread(target=app.control_loop)
    control_thread.start()
    root.mainloop()
    app.end_loop.set()
    control_thread.join()
    google_sheets_thread.join()

