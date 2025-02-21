import pandas as pd
import keyboard
import pyperclip
import time
import threading
import os
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime

# Dictionary to store user-selected settings
file_paths = {}
hotkeys = {}
columns = {}

lock = threading.Lock()

def select_file():
    column_name = simpledialog.askstring("Column Name", "Enter the column name:")
    if not column_name:
        return
    
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_paths[column_name] = file_path
        columns[column_name] = column_name
        hotkey = simpledialog.askstring("Hotkey", f"Enter hotkey for {column_name} (e.g., alt+j):")
        if hotkey:
            hotkeys[hotkey] = column_name
            try:
                keyboard.add_hotkey(hotkey, lambda c=column_name: on_key_press(c))
                print(f"Hotkey {hotkey} assigned to {column_name}")
            except ValueError as e:
                print(f"Error assigning hotkey {hotkey}: {e}")
        
        # Create file with headers if it doesn't exist
        if not os.path.exists(file_path):
            df = pd.DataFrame(columns=["Date", column_name])
            df.to_excel(file_path, index=False)
    print(f"File selected for {column_name}: {file_path}, assigned to hotkey {hotkey}")

def write_to_excel(column_name, data):
    file_path = file_paths.get(column_name, "")
    if not file_path:
        print(f"No file selected for {column_name}. Please select a file.")
        return
    
    with lock:
        df = pd.read_excel(file_path) if os.path.exists(file_path) else pd.DataFrame(columns=["Date", column_name])
        new_data = pd.DataFrame({"Date": [datetime.now().strftime('%Y-%m-%d %H:%M:%S')], column_name: [data]})
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(file_path, index=False)
        print(f"Added '{data}' to {column_name}")

def on_key_press(column_name):
    content = pyperclip.paste()
    write_to_excel(column_name, content)

def setup_gui():
    root = tk.Tk()
    root.title("Dynamic Excel Logger")
    
    tk.Button(root, text="Add File & Hotkey", command=select_file).pack(pady=10)
    tk.Button(root, text="Close", command=root.quit).pack(pady=10)
    root.mainloop()

def run_listener():
    keyboard.wait()

def main():
    gui_thread = threading.Thread(target=setup_gui)
    gui_thread.start()
    print("Listening for clipboard content... Press your assigned hotkeys to categorize data.")
    listener_thread = threading.Thread(target=run_listener, daemon=True)
    listener_thread.start()
    while True:
        time.sleep(1)  # Keep main thread alive

if __name__ == "__main__":
    main()
